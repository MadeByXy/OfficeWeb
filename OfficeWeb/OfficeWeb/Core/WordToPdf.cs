using System;
using System.IO;
using Word;

namespace OfficeWeb.Core
{
    /// <summary>
    /// 通过COM组件进行格式转换
    /// </summary>
    class WordToPdf : IDisposable
    {
        dynamic office;

        public WordToPdf()
        {
            Type type = Type.GetTypeFromProgID("KWps.Application") ?? Type.GetTypeFromProgID("wps.Application") ?? Type.GetTypeFromProgID("Word.Application");
            if (type == null)
            {
                throw new ArgumentNullException("未找到可用的COM组件。");
            }
            office = Activator.CreateInstance(type);
        }

        public string ToPdf(string wpsFilename, string pdfFilename = null)
        {
            if (wpsFilename == null) { throw new ArgumentNullException("wpsFilename"); }

            if (pdfFilename == null)
            {
                pdfFilename = Path.ChangeExtension(wpsFilename, "pdf");
            }
            if (!File.Exists(pdfFilename))
            {
                dynamic doc = office.Documents.Open(wpsFilename, Visible: false);
                doc.ExportAsFixedFormat(pdfFilename, WdExportFormat.wdExportFormatPDF);
                doc.Close();
            }

            return pdfFilename;
        }

        public void Dispose()
        {
            if (office != null) { office.Quit(); }
        }
    }
}
