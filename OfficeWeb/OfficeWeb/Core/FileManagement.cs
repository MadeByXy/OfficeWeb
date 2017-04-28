using System;
using System.Configuration;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeWeb.Core
{
    /// <summary>
    /// 文件管理
    /// </summary>
    public class FileManagement
    {
        /// <summary>
        /// 文件清理时间
        /// </summary>
        public static string FileClearDay
        {
            get { return ConfigurationManager.AppSettings[nameof(FileClearDay)] ?? 10.ToString(); }
            set
            {
                int day;
                if (int.TryParse(value, out day))
                {
                    ConfigurationManager.AppSettings[nameof(FileClearDay)] = value;
                }
            }
        }

        /// <summary>
        /// 是否自动清理
        /// </summary>
        public static string AutoClear
        {
            get { return ConfigurationManager.AppSettings[nameof(AutoClear)] ?? false.ToString(); }
            set
            {
                bool auto;
                if (bool.TryParse(value, out auto))
                {
                    ConfigurationManager.AppSettings[nameof(AutoClear)] = value;
                }
            }
        }
        private static DirectoryInfo FileDirectory { get { return new DirectoryInfo(string.Format(@"{0}\FileInfo\", AppDomain.CurrentDomain.BaseDirectory)); } }

        /// <summary>
        /// 清理过期缓存
        /// </summary>
        public static void Clear()
        {
            foreach (FileInfo file in FileDirectory.GetFiles("*", SearchOption.AllDirectories))
            {
                if ((DateTime.Now - file.CreationTime).TotalDays > double.Parse(FileClearDay))
                {
                    file.Delete();
                }
            }
        }

        /// <summary>
        /// 清理全部缓存
        /// </summary>
        public static void ClearAll()
        {
            FileDirectory.Delete(true);
        }

        private static readonly FileManagement Instance = new FileManagement();

        private FileManagement()
        {
            Task task = new Task(() =>
            {
                if (bool.Parse(AutoClear))
                {
                    Clear();
                }
                Thread.Sleep(24 * 60 * 60 * 1000);
            });
            task.Start();
        }
    }
}