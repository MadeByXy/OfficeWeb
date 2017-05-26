using Newtonsoft.Json.Linq;
using OfficeWeb.Core;
using System;
using System.Collections.Generic;
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
                Dictionary<string, CheckModel[]> checks = new Dictionary<string, CheckModel[]>();
                checks.Add("Word: 检测COM组件是否存在", new CheckModel[] {
                    new CheckModel() { Name = "WPS", Status = Type.GetTypeFromProgID("KWps.Application") != null || Type.GetTypeFromProgID("wps.Application") != null },
                    new CheckModel() { Name = "OFFICE", Status = Type.GetTypeFromProgID("Word.Application") != null }
                });

                checks.Add("Word: 检测COM组件是否可用", new Func<CheckModel[]>(() =>
                {
                    CheckModel[] checkModels = new CheckModel[] { new CheckModel() { Name = "WPS" }, new CheckModel() { Name = "OFFICE" } };
                    Type[] typeArray = new Type[] { Type.GetTypeFromProgID("KWps.Application") ?? Type.GetTypeFromProgID("wps.Application"), Type.GetTypeFromProgID("Word.Application") };
                    for (int i = 0; i < typeArray.Length; i++)
                    {
                        try
                        {
                            Activator.CreateInstance(typeArray[i]);
                            checkModels[i].Status = true;
                        }
                        catch (COMException e) { checkModels[i].Message = e.Message; }
                    }
                    return checkModels;
                }).Invoke());

                JArray resultArray = new JArray();
                foreach (string check in checks.Keys)
                {
                    JObject result = new JObject();
                    result.Add("name", check);
                    foreach (var checkModel in checks[check])
                    {
                        JObject item = new JObject();
                        item.Add("name", checkModel.Name);
                        item.Add("value", checkModel.Status);
                        item.Add("message", checkModel.Message);
                        result.Add("data", new JArray() { item });
                    }
                    resultArray.Add(result);
                }

                return resultArray;
            }
        }

        public class CheckModel
        {
            public string Name { get; set; }
            public bool Status { get; set; } = false;
            public string Message { get; set; }
        }
    }
}