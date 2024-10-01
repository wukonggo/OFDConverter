using ICSharpCode.SharpZipLib.Zip;
using XiaoFeng.Ofd.BasicStructure;
using XiaoFeng.Ofd.Internal;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OFDConverter
{
    public class OFDWriterHelper
    {
        #region 构造器
        /// <summary>
        /// 初始化一个新实例
        /// </summary>
        /// <param name="filePath">文档路径</param>
        public OFDWriterHelper()
        {
            this.Ofd = new OFD();
            this.Documents = new List<DocumentStructure>();
        }
        #endregion

        #region 属性
        /// <summary>
        /// OFD 文档
        /// </summary>
        public OFD Ofd { get; set; }
        /// <summary>
        /// 文档结构
        /// </summary>
        public List<DocumentStructure> Documents { get; set; }
        #endregion

        #region 方法

        #endregion

        #region 保存
        /// <summary>
        /// 保存
        /// </summary>
        /// <param name="filePath">保存地址</param>

        public bool SaveFile(string filePath)
        {
            if (this.Ofd == null)
                throw new Exception("OFD入口文档写入出错");
            using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                using (ZipOutputStream zipStream = new ZipOutputStream(fs))
                {
                    zipStream.SetLevel(6); // Set the compression level

                    // Create and write OFD.xml
                    WriteEntry(zipStream, "OFD.xml", this.Ofd.GetBytes());

                    for (int i = 0; i < this.Ofd.DocBody.Count; i++)
                    {
                        var body = this.Ofd.DocBody[i];
                        var Structure = this.Documents[i];
                        string rootPath = body.DocRoot.ToString().Substring(0, body.DocRoot.ToString().LastIndexOf("/") + 1);

                        // Write Document.xml
                        WriteEntry(zipStream, body.DocRoot.ToString(), Structure.Document.GetBytes());

                        // Write DocumentRes.xml
                        WriteEntry(zipStream, rootPath + Structure.Document.CommonData.First().DocumentRes, Structure.DocumentRes.GetBytes());

                        // Write attachments
                        var medias = Structure.DocumentRes.MultiMedias;
                        if (medias != null && medias.Count > 0)
                        {
                            foreach (var m in medias)
                            {
                                string mediaPath = $"{rootPath + Structure.DocumentRes.BaseLoc}/{m.MediaFile}";
                                WriteEntry(zipStream, mediaPath, m.Bytes);
                            }
                        }

                        // Write PublicRes.xml
                        WriteEntry(zipStream, rootPath + Structure.Document.CommonData.First().PublicRes, Structure.PublicRes.GetBytes());

                        // Write Pages/Page_0/Content.xml
                        for (int j = 0; j < Structure.Document.Pages.Count; j++)
                        {
                            var p = Structure.Document.Pages[j];
                            WriteEntry(zipStream, rootPath + p.BaseLoc, Structure.Pages[j].GetBytes());
                        }

                        // Write template pages Tpls/Tpl_0/Content.xml
                        var templatePages = Structure.Document.CommonData.FirstOrDefault().TemplatePage;
                        if (templatePages != null && templatePages.Count > 0)
                        {
                            for (int j = 0; j < templatePages.Count; j++)
                            {
                                var t = templatePages[j];
                                WriteEntry(zipStream, rootPath + t.BaseLoc, Structure.TemplatePages[j].GetBytes());
                            }
                        }

                    }
                    zipStream.Finish();
                    zipStream.Close();
                } 
            } 
            return true;
        }

        private void WriteEntry(ZipOutputStream zipStream, string entryName, byte[] data)
        {
            var newEntry = new ZipEntry(entryName)
            {
                DateTime = DateTime.Now,
                Size = data.Length
            };

            zipStream.PutNextEntry(newEntry);
            zipStream.Write(data, 0, data.Length);
            zipStream.CloseEntry();
        } 
        #endregion 
    }
}