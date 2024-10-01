using PdfiumViewer;
using XiaoFeng.Ofd.BasicStructure;
using XiaoFeng.Ofd.Enum;
using XiaoFeng.Ofd.Images;
using System;
using System.Drawing;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OFDConverter
{
    public class PdfConverter
    {
        private static readonly Dictionary<string, SizeF> FixedPageSizesInMM = new Dictionary<string, SizeF>
            {
                // A Series
                { "A0", new SizeF(841f, 1189f) }, { "A1", new SizeF(594f, 841f) }, { "A2", new SizeF(420f, 594f) },
                { "A3", new SizeF(297f, 420f) }, { "A4", new SizeF(210f, 297f) }, { "A5", new SizeF(148f, 210f) },
                { "A6", new SizeF(105f, 148f) }, { "A7", new SizeF(74f, 105f) }, { "A8", new SizeF(52f, 74f) },
                { "A9", new SizeF(37f, 52f) }, { "A10", new SizeF(26f, 37f) },
            
                // B Series
                { "B0", new SizeF(1000f, 1414f) }, { "B1", new SizeF(707f, 1000f) }, { "B2", new SizeF(500f, 707f) },
                { "B3", new SizeF(353f, 500f) }, { "B4", new SizeF(250f, 353f) }, { "B5", new SizeF(176f, 250f) },
                { "B6", new SizeF(125f, 176f) }, { "B7", new SizeF(88f, 125f) }, { "B8", new SizeF(62f, 88f) },
                { "B9", new SizeF(44f, 62f) }, { "B10", new SizeF(31f, 44f) },

                { "Letter", new SizeF(216, 279) }, // 8.5" x 11"
                { "Legal", new SizeF(216, 356) },  // 8.5" x 14"
                { "Tabloid", new SizeF(279, 432) }, // 11" x 17"
                { "Ledger", new SizeF(432, 279) },  // 17" x 11"
                { "Executive", new SizeF(184, 267) }, // 7.25" x 10.5"
                { "Postcard", new SizeF(100, 148) } // 4" x 6"
            };

        private PdfConverterConfig _config;

        public PdfConverterConfig Config
        {
            get
            {
                if (_config == null)
                {
                    _config = new PdfConverterConfig();
                }
                return _config;
            }
        }

        public PdfConverter()
        {
        }

        public PdfConverter(PdfConverterConfig config)
        {
            this._config = config ?? new PdfConverterConfig();
        }

        private void EnsureTempDir()
        {
            if (string.IsNullOrWhiteSpace(this.Config.TempDirectory))
            {
                this.Config.TempDirectory = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "temp");
                if (!Directory.Exists(this.Config.TempDirectory))
                    Directory.CreateDirectory(this.Config.TempDirectory);
            }
        }

        public void ToOfd(string sourcePdfFile, string destOfdFile)
        {
            if (!System.IO.File.Exists(sourcePdfFile))
                throw new FileNotFoundException(sourcePdfFile);
            EnsureTempDir();
            var docUUID = Guid.NewGuid().ToString("N");
            var pageImages = new List<string>();
            var pdfPageSizeInPoints = System.Drawing.SizeF.Empty;
            try
            {
                var dpi = this.Config.ImageDPI;
                if (dpi <= 72f)
                    dpi = 200f;
                using (var pdfDocument = PdfDocument.Load(sourcePdfFile))
                {
                    for (int pageIndex = 0; pageIndex < pdfDocument.PageCount; pageIndex++)
                    {
                        using (var image = pdfDocument.Render(pageIndex, dpi, dpi, PdfRenderFlags.CorrectFromDpi))
                        {
                            var imageFilePath = Path.Combine(this.Config.TempDirectory, $"{docUUID}_{pageIndex}.png");
                            image.Save(imageFilePath, System.Drawing.Imaging.ImageFormat.Png);
                            pageImages.Add(imageFilePath);
                        }
                    }
                    pdfPageSizeInPoints = pdfDocument.PageSizes.FirstOrDefault();
                }

                var pageAreaSizeInMM = GetPageSizeInMM(new SizeF(
                    (pdfPageSizeInPoints.Width * 25.4f) / 72f,
                    (pdfPageSizeInPoints.Height * 25.4f) / 72f));

                WriteToOfd(destOfdFile, docUUID, pageImages, pageAreaSizeInMM);
            }
            finally
            {
                DeleteFiles(pageImages);
            }
        }

        private void WriteToOfd(string destOfdFile, string docUUID, List<string> pageImages, SizeF pageAreaSizeInMM)
        {
            OFDWriterHelper writer = new OFDWriterHelper();
            var docBody = new DocBody()
            {
                DocInfo = new DocInfo()
                {
                    DocID = docUUID,
                    Title = "",   //文档标题
                    CreationDate = DateTime.Now,
                    ModDate = DateTime.Now,
                    CreatorVersion = "1.0",
                    Creator = "wukonggo",
                    Author = "wukonggo"
                },
                DocRoot = "Doc_0/Document.xml"
            };
            writer.Ofd.DocType = DocumentType.OFD;
            writer.Ofd.AddDocBody(docBody);

            var docStruct = new XiaoFeng.Ofd.Internal.DocumentStructure();
            var doc = new Document();
            var commonData = new CommonData
            {
                MaxUnitID = 0,
                PageArea = new PageArea
                {
                    PhysicalBox = new XiaoFeng.Ofd.BaseType.Box($"0 0 {pageAreaSizeInMM.Width} {pageAreaSizeInMM.Height}")
                }
            };
            doc.AddCommonData(commonData);

            docStruct.Document = doc;

            var pubRes = new PublicRes();
            var simSun = XiaoFeng.Ofd.Fonts.Font.SimSun((uint)doc.GetMaxUnitIDAndAdd());
            pubRes.AddFont(simSun);
            docStruct.PublicRes = pubRes;

            for (int i = 0; i < pageImages.Count; i++)
            {
                var page = new XiaoFeng.Ofd.BasicStructure.Page();
                var layer = new Layer()
                {
                    ID = doc.GetMaxUnitIDAndAdd(),
                    Type = LayerType.Body
                };
                var imageRes = new MultiMedia()
                {
                    ID = doc.GetMaxUnitIDAndAdd(),
                    Format = "PNG",
                    Type = MultiMediaType.Image,
                    MediaFile = $"Image_{i}.png",
                    Bytes = System.IO.File.ReadAllBytes(pageImages[i]),
                };
                var imgobj = new ImageObject();
                imgobj.ID = doc.GetMaxUnitIDAndAdd();
                imgobj.ResourceID = imageRes.ID;
                imgobj.Boundary = new XiaoFeng.Ofd.BaseType.Box($"0 0 {pageAreaSizeInMM.Width} {pageAreaSizeInMM.Height}");
                imgobj.CTM = new XiaoFeng.Ofd.BaseType.STArray($"{pageAreaSizeInMM.Width} 0 0 {pageAreaSizeInMM.Height} 0 0");
                layer.AddPageBlock(imgobj);

                page.AddLayer(layer);

                docStruct.AddPage(page);
                docStruct.DocumentRes.AddMultiMedia(imageRes);
            }
            writer.Documents.Add(docStruct);

            writer.SaveFile(destOfdFile);
        }

        private SizeF GetPageSizeInMM(SizeF originalInMM)
        {
            foreach (var pageSize in FixedPageSizesInMM.Values)
            {
                if (IsApproximatelyEqual(originalInMM.Width, pageSize.Width) && IsApproximatelyEqual(originalInMM.Height, pageSize.Height))
                {
                    return pageSize;
                }
            }
            return originalInMM;
        }

        private bool IsApproximatelyEqual(float value, float target)
        {
            float tolerance = (target * 0.5f / 100f);
            return Math.Abs(value - target) <= tolerance;
        }

        private void DeleteFiles(List<string> files)
        {
            if (files == null || files.Count == 0)
                return;
            foreach (var fileItem in files)
            {
                try
                {
                    if (System.IO.File.Exists(fileItem))
                    {
                        System.IO.File.Delete(fileItem);
                    }
                }
                catch { }
            }
        }
    }
}
