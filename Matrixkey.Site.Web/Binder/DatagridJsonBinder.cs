using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Matrixkey.Site.Web
{
    /// <summary>
    /// 对DataGrid的data的json数据进行反序列化处理的绑定类
    /// 处理rows时，返回DataTable
    /// </summary>
    public class DatagridRowsJsonBinder : IModelBinder
    {
        public object BindModel(ControllerContext controllerContext, ModelBindingContext bindingContext)
        {
            //从请求中获取提交的参数数据 
            var json = controllerContext.HttpContext.Request.Form[bindingContext.ModelName] as string;
            //提交参数是数组 传入的rows和columns都是数组，故不做对象处理
            if (json.StartsWith("[") && json.EndsWith("]"))
            {
                if (json.StartsWith("[[") && json.EndsWith("]]"))
                {
                    //处理columns
                }
                else
                {
                    //处理rows，返回DataTable
                    JArray jsonRsp = JArray.Parse(json);
                    DataTable dt = new DataTable();
                    if (jsonRsp != null)
                    {
                        dt = (DataTable)JsonConvert.DeserializeObject<DataTable>(json);
                    }

                    return dt;
                }
            }
            return null;
        }
    }


    /// <summary>
    /// 对DataGrid的column的json数据进行反序列化处理的绑定类
    /// 处理columns时，返回DataTable[]
    /// </summary>
    public class DatagridColumnsJsonBinder : IModelBinder
    {
        public object BindModel(ControllerContext controllerContext, ModelBindingContext bindingContext)
        {
            //从请求中获取提交的参数数据 
            var json = controllerContext.HttpContext.Request.Form[bindingContext.ModelName] as string;
            //提交参数是数组 传入的rows和columns都是数组，故不做对象处理
            if (json.StartsWith("[") && json.EndsWith("]"))
            {
                if (json.StartsWith("[[") && json.EndsWith("]]"))
                {
                    //处理columns，返回DataTable[]
                    //columns可能是合并列
                    JArray jsonRsp = JArray.Parse(json);
                    IList<DataTable> list = new List<DataTable>();

                    if (jsonRsp != null)
                    {
                        JsonSerializer js = new JsonSerializer();
                        for (int i = 0; i < jsonRsp.Count; i++)
                        {
                            object obj = js.Deserialize(jsonRsp[i].CreateReader(), typeof(DataTable));
                            list.Add((DataTable)obj);
                        }
                    }

                    var temp = list.ToArray();
                    List<DataTable> dts = new List<DataTable>();
                    if (temp.Length == 0) { return dts; }

                    string fieldColumnName = "field", titleColumnName = "title", rowspanColumnName = "rowspan", colspanColumnName = "colspan";
                    foreach (var item in temp)
                    {
                        bool fieldExist = item.Columns.Contains(fieldColumnName);
                        bool titleExist = item.Columns.Contains(titleColumnName);

                        DataTable dt = new DataTable();
                        dt.Columns.Add(fieldColumnName, typeof(string));
                        dt.Columns.Add(titleColumnName, typeof(string));
                        dt.Columns.Add(rowspanColumnName, typeof(string));
                        dt.Columns.Add(colspanColumnName, typeof(string));
                        foreach (DataRow row in item.Rows)
                        {
                            var newRow = dt.NewRow();
                            newRow[fieldColumnName] = fieldExist ? row[fieldColumnName] : "";
                            newRow[titleColumnName] = titleExist ? row[titleColumnName] : "";
                            newRow[rowspanColumnName] = row[rowspanColumnName];
                            newRow[colspanColumnName] = row[colspanColumnName];
                            dt.Rows.Add(newRow);
                        }
                        dts.Add(dt);
                    }

                    return dts.ToArray();
                }
            }
            return null;
        }
    }
}