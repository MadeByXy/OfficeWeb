using Newtonsoft.Json.Linq;
using OfficeWeb.Core;
using System;
using System.Threading.Tasks;

namespace OfficeWeb
{
    public partial class Office : System.Web.UI.Page
    {
        /// <summary>
        /// 页面ID
        /// </summary>
        public string SessionId { get; set; }

        protected void Page_Load(object sender, EventArgs e)
        {
            switch (Request["Mode"])
            {
                case "GetResult":
                    GetResult(Request["PageId"]);
                    return;
                default:
                    SessionId = Session.SessionID + Request["FileUrl"];
                    GetFile(Request["FileUrl"]); return;
            }
        }

        /// <summary>
        /// 开始加载Office文件
        /// </summary>
        /// <param name="fileUrl">文件Url</param>
        private void GetFile(string fileUrl)
        {
            //避免页面加载完成前重复刷新造成的重复获取
            if (!CacheData.ContainsKey(SessionId))
            {
                CacheData.Set(SessionId, new JObject() { { "status", "running" }, { "message", "" }, { "imageList", new JArray() } });

                new Task(() =>
                {
                    if (!string.IsNullOrEmpty(fileUrl))
                    {
                        try
                        {
                            Network network = new Network(fileUrl);
                            new Conversion(network.FilePath, network.FileExtension, CallBack).ConvertToImage();
                        }
                        catch (Exception ex) { AlertError(ex.Message); }
                    }
                    else
                    {
                        AlertError("请输入文件地址");
                    }
                }).Start();
            }
        }

        /// <summary>
        /// 返回处理结果
        /// </summary>
        /// <param name="pageId">页面唯一ID</param>
        private void GetResult(string pageId)
        {
            Response.Clear();
            Response.ContentType = "application/json";
            Response.Write(CacheData.Get(pageId));

            string status = CacheData.Get(pageId)["status"].ToString();
            if (status == "finish" || status == "error")
            {
                CacheData.Remove(pageId);
            }

            Response.End();
        }

        /// <summary>
        /// 获取图片地址
        /// </summary>
        /// <param name="pageNum">当前图片次序</param>
        /// <param name="imageUrl">图片Url</param>
        public void CallBack(int pageNum, string imageUrl)
        {
            lock (this)
            {
                JObject jObject = CacheData.Get(SessionId);
                if (jObject["status"].ToString() == "running")
                {
                    if (pageNum == 0)
                    {
                        jObject["status"] = "finish";
                    }
                    else
                    {
                        ((JArray)jObject["imageList"]).Add(new JObject() { { "pageNum", pageNum }, { "imageUrl", FromLocal(imageUrl) } });
                    }
                    CacheData.Set(SessionId, jObject);
                }
            }
        }

        /// <summary>
        /// 页码报错提示
        /// </summary>
        /// <param name="message">提示信息</param>
        private void AlertError(string message)
        {
            CacheData.Set(SessionId, new JObject() { { "status", "error" }, { "message", message }, { "imageList", new JArray() } });
        }

        /// <summary>
        /// 本地路径转Url
        /// </summary>
        private string FromLocal(string filePath)
        {
            string fileUrl = filePath.Replace(AppDomain.CurrentDomain.BaseDirectory, "");  //转换成相对路径
            fileUrl = "/" + fileUrl.Replace(@"\", @"/");
            return fileUrl;
        }
    }
}
