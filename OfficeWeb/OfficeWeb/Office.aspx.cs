using Newtonsoft.Json.Linq;
using OfficeWeb.Core;
using System;
using System.Threading.Tasks;

namespace OfficeWeb
{
    public partial class Office : System.Web.UI.Page
    {
        /// <summary>
        /// ҳ��ID
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
        /// ��ʼ����Office�ļ�
        /// </summary>
        /// <param name="fileUrl">�ļ�Url</param>
        private void GetFile(string fileUrl)
        {
            //����ҳ��������ǰ�ظ�ˢ����ɵ��ظ���ȡ
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
                        AlertError("�������ļ���ַ");
                    }
                }).Start();
            }
        }

        /// <summary>
        /// ���ش�����
        /// </summary>
        /// <param name="pageId">ҳ��ΨһID</param>
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
        /// ��ȡͼƬ��ַ
        /// </summary>
        /// <param name="pageNum">��ǰͼƬ����</param>
        /// <param name="imageUrl">ͼƬUrl</param>
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
        /// ҳ�뱨����ʾ
        /// </summary>
        /// <param name="message">��ʾ��Ϣ</param>
        private void AlertError(string message)
        {
            CacheData.Set(SessionId, new JObject() { { "status", "error" }, { "message", message }, { "imageList", new JArray() } });
        }

        /// <summary>
        /// ����·��תUrl
        /// </summary>
        private string FromLocal(string filePath)
        {
            string fileUrl = filePath.Replace(AppDomain.CurrentDomain.BaseDirectory, "");  //ת�������·��
            fileUrl = "/" + fileUrl.Replace(@"\", @"/");
            return fileUrl;
        }
    }
}
