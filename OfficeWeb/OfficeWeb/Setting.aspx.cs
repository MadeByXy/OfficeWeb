using Newtonsoft.Json.Linq;
using OfficeWeb.Core;
using System;

namespace OfficeWeb
{
    public partial class Setting : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                switch (Request["Action"])
                {
                    case "Property":
                        typeof(FileManagement).GetProperty(Request["Name"])?.SetValue(null, Request["Value"], null);
                        return;
                    case "Function": typeof(FileManagement).GetMethod(Request["Name"])?.Invoke(null, null); return;
                    default: return;
                }
            }
            catch (Exception ex)
            {
                Response.Clear();
                Response.Write(ex.Message);
                Response.End();
            }
        }

        /// <summary>
        /// COM环境检测
        /// </summary>
        public JArray EnvironmentCheck
        {
            get
            {
                JArray wordArray = new JArray();
                wordArray.Add(new JObject() { { "name", "WPS" }, { "value", Type.GetTypeFromProgID("KWps.Application") != null || Type.GetTypeFromProgID("wps.Application") != null } });
                wordArray.Add(new JObject() { { "name", "OFFICE" }, { "value", Type.GetTypeFromProgID("Word.Application") != null } });
                JObject wordCheck = new JObject() { { "name", "Word" }, { "data", wordArray } };

                return new JArray() { wordCheck };
            }
        }
    }
}