using System;
using System.IO;
using System.Net;

namespace OfficeWeb.Core
{
    /// <summary>
    /// 从URL获取文件
    /// </summary>
    public class Network
    {
        /// <summary>
        /// 文件本地路径
        /// </summary>
        public string FilePath { get; }
        /// <summary>
        /// 文件扩展名
        /// </summary>
        public string FileExtension
        {
            get
            {
                return FilePath.Substring(FilePath.LastIndexOf("."));
            }
        }

        public Network(string url)
        {
            FilePath = GetFileInfo(new Uri(url));
        }

        private static string GetFileInfo(Uri uri)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            request.ContentType = "application/x-www-form-urlencoded";
            string filePath = GetFilePath(uri);

            if (!File.Exists(filePath))
            {
                using (WebResponse response = request.GetResponse())
                {
                    int length = (int)response.ContentLength;
                    using (BinaryReader reader = new BinaryReader(response.GetResponseStream()))
                    {
                        using (FileStream stream = File.Create(filePath))
                        {
                            stream.Write(reader.ReadBytes(length), 0, length);
                        }
                    }
                }
            }
            return filePath;
        }

        /// <summary>
        /// 获取文件本地路径
        /// </summary>
        /// <param name="uri">网络请求地址</param>
        private static string GetFilePath(Uri uri)
        {
            string localPath = uri.LocalPath;
            string filePath = localPath.Substring(localPath.LastIndexOf("/") + 1);
            string folderPath = string.Format(@"{0}\FileInfo\{1}\{2}\", AppDomain.CurrentDomain.BaseDirectory, uri.Host, localPath.Remove(localPath.LastIndexOf("/")).Replace("/", @"\"));
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            return folderPath + filePath;
        }
    }
}
