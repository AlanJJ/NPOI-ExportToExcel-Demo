using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;

using Matrixkey.Utility.Extensions;
using Matrixkey.Utility.Helper;

namespace Matrixkey.Site.Web.Controllers
{
    public class TestController : Controller
    {
        public ActionResult Index(string type = "4")
        {
            if (!new string[] { "1", "2", "3", "4" }.Contains(type)) { type = "4"; }
            ViewBag.InitType = type;

            return View();
        }

        /// <summary>
        /// datagrid导出数据
        /// </summary>
        /// <param name="data">datagrid的data数据的json格式，将被反序列化成 DataTable </param>
        /// <param name="columns">datagrid的columns数据的json格式，将被反序列化成 DataTable[]，其中列是固定的“field、title、rowspan、colspan” </param>
        /// <param name="fileName">excel文件名</param>
        /// <returns></returns>
        public FileResult ExportDataByDataGrid([ModelBinder(typeof(DatagridRowsJsonBinder))]System.Data.DataTable data, [ModelBinder(typeof(DatagridColumnsJsonBinder))]System.Data.DataTable[] columns, string fileName)
        {
            if (columns.Length == 0) { throw new Exception("列不存在，无法导出！"); }

            IEnumerable<Matrixkey.Utility.Helper.MergerCellParam> rules = null;
            DataTable dt = Matrixkey.Utility.Helper.ExportHelper.ParseColumnsAndRows(columns, data, out rules);

            var content = Matrixkey.Utility.Helper.ExportHelper.ExportToExcel(dt, false, rules, true, true, false);
            if (string.IsNullOrWhiteSpace(fileName)) { fileName = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + "导出"; }
            return File(content, "application/ms-excel", Url.Encode(fileName + ".xls"));
        }
    }
}