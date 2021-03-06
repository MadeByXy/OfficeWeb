﻿using System;
using System.IO;
using System.Runtime.InteropServices;
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
            foreach (Type type in new Type[] { Type.GetTypeFromProgID("Word.Application"), Type.GetTypeFromProgID("KWps.Application"), Type.GetTypeFromProgID("wps.Application") })
            {
                if (type != null)
                {
                    try
                    {
                        office = Activator.CreateInstance(type);
                        return;
                    }
                    catch (COMException) { }
                }
            }
            throw new COMException("未找到可用的COM组件。");
        }

        public string ToPdf(string wpsFilename, string pdfFilename = null)
        {
            if (wpsFilename == null) { throw new COMException("未找到指定文件。"); }

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
