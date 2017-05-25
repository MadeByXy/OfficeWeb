using Newtonsoft.Json.Linq;
using OfficeWeb.Core;
using System;
using System.Runtime.InteropServices;

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
                JObject wordCheck = new JObject() { { "name", "Word: 检测COM组件是否存在" }, { "data", new JArray() {
                    new JObject() { { "name", "WPS" }, { "value", Type.GetTypeFromProgID("KWps.Application") != null || Type.GetTypeFromProgID("wps.Application") != null } },
                    new JObject() { { "name", "OFFICE" }, { "value", Type.GetTypeFromProgID("Word.Application") != null } }
                } } };

                bool[] array = new bool[2];
                Type[] typeArray = new Type[] { Type.GetTypeFromProgID("KWps.Application") ?? Type.GetTypeFromProgID("wps.Application"), Type.GetTypeFromProgID("Word.Application") };
                for (int i = 0; i < typeArray.Length; i++)
                {
                    try
                    {
                        Activator.CreateInstance(typeArray[i]);
                        array[i] = true;
                    }
                    catch (COMException) { }
                }

                JObject canUseCheck = new JObject() { { "name", "Word: 检测COM组件是否可用" }, { "data", new JArray() {
                    new JObject() { { "name", "WPS" }, { "value", array[0] } },
                    new JObject() { { "name", "OFFICE" }, { "value", array[1] } }
                } } };

                return new JArray() { wordCheck, canUseCheck };
            }
        }
    }
}