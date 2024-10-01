using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OFDConverter.Test
{
    public partial class FormDemo : Form
    {
        public FormDemo()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var pdffileName = txtfile.Text.Trim();
            var dpi = 200f;
            float.TryParse(this.txtDPI.Text.Trim(), out dpi);
            if (string.IsNullOrWhiteSpace(pdffileName))
                return;
            if (!File.Exists(pdffileName))
                return;
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            var fileName = Path.GetFileNameWithoutExtension(pdffileName);
            string outOfdFile = Path.Combine(Application.StartupPath, $"{fileName}.ofd");
            try
            {
                PdfConverter helper = new PdfConverter(new PdfConverterConfig()
                {
                    ImageDPI = dpi
                });
                helper.ToOfd(pdffileName, outOfdFile);
            }
            catch (Exception)
            {
                throw;
            }
            stopwatch.Stop();

            MessageBox.Show(this, $"Success! Elapsed:{stopwatch.ElapsedMilliseconds}ms");

            if (!string.IsNullOrWhiteSpace(outOfdFile))
                Process.Start(outOfdFile);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "pdf files (*.pdf)|*.pdf|All files (*.*)|*.*";
            openFileDialog.Title = "打开文件";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txtfile.Text = openFileDialog.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch();

            stopwatch.Start();
            var dpi = 200f;
            float.TryParse(this.txtDPI.Text.Trim(), out dpi);
            try
            {
                var pdffiles = Directory.GetFiles(@"TestPdfs", "*.pdf");
                PdfConverter helper = new PdfConverter(new PdfConverterConfig()
                {
                    ImageDPI = dpi
                });
                for (int i = 0; i < 1; i++)
                {
                    foreach (var pdf in pdffiles)
                    {
                        var pdffileName = Path.GetFileNameWithoutExtension(pdf);
                        var outpath = Path.Combine(Application.StartupPath, $"{pdffileName}_{i}.ofd");
                        helper.ToOfd(pdf, outpath);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            stopwatch.Stop();

            MessageBox.Show(this, $"Success! Elapsed:{stopwatch.ElapsedMilliseconds}ms");
        }
    }
}
