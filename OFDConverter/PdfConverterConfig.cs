using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OFDConverter
{
    public class PdfConverterConfig
    {
        /// <summary>
        /// 临时目录 用于转换文件时产生临时文件
        /// </summary>
        public string TempDirectory { get; set; }

        private float _imageDPI = 200f;
        /// <summary>
        /// 转换pdf为图片设置的dpi
        /// </summary>
        public float ImageDPI
        {
            get { return _imageDPI; }
            set
            {
                _imageDPI = value;
            }
        }
    }
}
