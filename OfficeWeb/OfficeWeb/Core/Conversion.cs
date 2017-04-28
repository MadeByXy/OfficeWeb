using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading.Tasks;

namespace OfficeWeb.Core
{
    public interface IConversion
    {
        /// <summary>
        /// 回调方法
        /// </summary>
        Action<int, string> CallBack { get; set; }
        /// <summary>
        /// Office转换成图片
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="resolution">图片质量</param>
        void ConvertToImage(string filePath, int resolution);
    }

    /// <summary>
    /// Office转换工具
    /// </summary>
    public class Conversion
    {
        private IConversion ConversionFunc { get; set; }
        /// <summary>
        /// 文件路径
        /// </summary>
        private string FilePath { get; set; }
        /// <summary>
        /// 图片转换质量
        /// </summary>
        private const int Resolution = 200;
        /// <summary>
        /// 实例化转换工具
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="fileExtension">文件扩展名</param>
        /// <param name="callBack">回调方法</param>
        public Conversion(string filePath, string fileExtension, Action<int, string> callBack)
        {
            switch (fileExtension.ToLower())
            {
                case ".doc":
                case ".docx":
                    ConversionFunc = new WordConversion();
                    break;
                case ".pdf":
                    ConversionFunc = new PdfConversion();
                    break;
                case ".xls":
                case ".xlsx":
                    ConversionFunc = new ExcelConversion();
                    break;
                case ".ppt":
                case ".pptx":
                    ConversionFunc = new PptConversion();
                    break;
                default:
                    throw new Exception("不支持的文件类型。");
            }
            FilePath = filePath;
            ConversionFunc.CallBack = callBack;
        }

        /// <summary>
        /// Office转换成图片
        /// </summary>
        public void ConvertToImage()
        {
            ConversionFunc.ConvertToImage(FilePath, Resolution);
            ConversionFunc.CallBack(0, "");
        }
    }

    class PdfConversion : IConversion
    {
        private List<Task> TaskList = new List<Task>();
        public Action<int, string> CallBack { get; set; }

        public void ConvertToImage(string filePath, int resolution)
        {
            Aspose.Pdf.Document doc = new Aspose.Pdf.Document(filePath);
            string imageName = Path.GetFileNameWithoutExtension(filePath);

            for (int i = 1; i <= doc.Pages.Count; i++)
            {
                int pageNum = i;
                string imgPath = string.Format("{0}_{1}.Jpeg", Path.Combine(Path.GetDirectoryName(filePath), imageName), i.ToString("000"));
                if (File.Exists(imgPath))
                {
                    InvokeCallBack(pageNum, imgPath);
                    continue;
                }

                using (MemoryStream stream = new MemoryStream())
                {
                    Aspose.Pdf.Devices.Resolution reso = new Aspose.Pdf.Devices.Resolution(resolution);
                    Aspose.Pdf.Devices.JpegDevice jpegDevice = new Aspose.Pdf.Devices.JpegDevice(reso, 100);
                    jpegDevice.Process(doc.Pages[i], stream);
                    using (Image image = Image.FromStream(stream))
                    {
                        new Bitmap(image).Save(imgPath, ImageFormat.Jpeg);
                    }
                }
                InvokeCallBack(pageNum, imgPath);
            }

            Task.WaitAll(TaskList.ToArray());
        }

        private void InvokeCallBack(int pageNum, string imagePath)
        {
            Task task = new Task(() => { CallBack.Invoke(pageNum, imagePath); });
            TaskList.Add(task);
            task.Start();
        }
    }

    class WordConversion : IConversion
    {
        private List<Task> TaskList = new List<Task>();
        public Action<int, string> CallBack { get; set; }

        public void ConvertToImage(string filePath, int resolution)
        {
            //先将Word转换为PDF文件
            string pdfPath;
            try
            {
                using(WordToPdf convert = new WordToPdf())
                {
                    pdfPath = convert.ToPdf(filePath);
                }
            }
            catch(ArgumentNullException)
            {
                pdfPath = Path.ChangeExtension(filePath, "pdf");
                Aspose.Words.Document doc = new Aspose.Words.Document(filePath);
                doc.Save(pdfPath, Aspose.Words.SaveFormat.Pdf);
            }

            //再将PDF转换为图片
            PdfConversion converter = new PdfConversion() { CallBack = CallBack };
            converter.ConvertToImage(pdfPath, resolution);
        }

        private void InvokeCallBack(int pageNum,string imagePath)
        {
            Task task = new Task(() => { CallBack.Invoke(pageNum, imagePath); });
            TaskList.Add(task);
            task.Start();
        }
    }

    class ExcelConversion : IConversion
    {
        public Action<int, string> CallBack { get; set; }

        public void ConvertToImage(string filePath, int resolution)
        {
            Aspose.Cells.Workbook doc = new Aspose.Cells.Workbook(filePath);

            //先将Excel转换为PDF临时文件
            string pdfPath = Path.ChangeExtension(filePath, "pdf");
            doc.Save(pdfPath, Aspose.Cells.SaveFormat.Pdf);

            //再将PDF转换为图片
            PdfConversion converter = new PdfConversion() { CallBack = CallBack };
            converter.ConvertToImage(filePath, resolution);
        }
    }

    class PptConversion : IConversion
    {
        public Action<int, string> CallBack { get; set; }

        public void ConvertToImage(string filePath, int resolution)
        {
            Aspose.Slides.Presentation doc = new Aspose.Slides.Presentation(filePath);

            //先将ppt转换为PDF文件
            string tmpPdfPath = Path.ChangeExtension(filePath, "pdf");
            doc.Save(tmpPdfPath, Aspose.Slides.Export.SaveFormat.Pdf);

            //再将PDF转换为图片
            PdfConversion converter = new PdfConversion() { CallBack = CallBack };
            converter.ConvertToImage(filePath, resolution);
        }
    }
}