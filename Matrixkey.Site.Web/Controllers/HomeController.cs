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
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public FileResult Test1(int rc = 51, int cc = 10, string fileName = "测试DataTable无合并方式导出到Excel")
        {
            DataTable dt = GetDt(rc, cc);

            var content = Matrixkey.Utility.Helper.ExportHelper.ExportToExcel(dt, true, null, true, true, false);

            return File(content, "application/ms-excel", Url.Encode(fileName + ".xls"));
        }

        public FileResult Test2(int rc = 51, int cc = 10, string fileName = "测试DataTable，以不含列名、无合并方式导出到Excel")
        {
            DataTable dt = GetDt(rc, cc);

            var content = Matrixkey.Utility.Helper.ExportHelper.ExportToExcel(dt, false, null, true, true, false);

            return File(content, "application/ms-excel", Url.Encode(fileName + ".xls"));
        }

        public FileResult Test3(string fileName = "测试DataSet无合并方式导出到Excel")
        {
            DataTable dt1 = GetDt(2, 10);
            DataTable dt2 = GetDt(10, 10);
            DataSet ds = new DataSet();
            ds.Tables.Add(dt1); ds.Tables.Add(dt2);

            var content = Matrixkey.Utility.Helper.ExportHelper.ExportToExcel(ds, null, true, true, false);

            return File(content, "application/ms-excel", Url.Encode(fileName + ".xls"));
        }

        public FileResult Test4(int rc = 51, int cc = 10, string fileName = "测试DataTable带合并方式导出到Excel")
        {
            DataTable dt = GetDt(rc, cc);

            List<MergerCellParam> rules = new List<MergerCellParam>() { 
                new MergerCellParam { RowStartIndex = 1, ColStartIndex = 0, RowEndIndex = 3, ColEndIndex = 0, Repeat = false, IntervalRowCount = 0, IntervalColCount = -1 }
            };

            var content = Matrixkey.Utility.Helper.ExportHelper.ExportToExcel(dt, true, rules, true, true, false);

            return File(content, "application/ms-excel", Url.Encode(fileName + ".xls"));
        }

        public FileResult Test5(int rc = 51, int cc = 10, string fileName = "测试DataTable带合并方式导出到Excel")
        {
            DataTable dt = GetDt(rc, cc);

            List<MergerCellParam> rules = new List<MergerCellParam>() { 
                new MergerCellParam { RowStartIndex = 1, ColStartIndex = 0, RowEndIndex = 3, ColEndIndex = 0, Repeat = true, IntervalRowCount = 0, IntervalColCount = -1 }
            };

            var content = Matrixkey.Utility.Helper.ExportHelper.ExportToExcel(dt, true, rules, true, true, false);

            return File(content, "application/ms-excel", Url.Encode(fileName + ".xls"));
        }

        public FileResult Test6(int rc = 51, int cc = 10, string fileName = "测试DataTable带合并方式导出到Excel")
        {
            DataTable dt = GetDt(rc, cc);

            List<MergerCellParam> rules = new List<MergerCellParam>() { 
                new MergerCellParam { RowStartIndex = 1, ColStartIndex = 0, RowEndIndex = 3, ColEndIndex = 0, Repeat = true, IntervalRowCount = 1, IntervalColCount = -1 }
            };

            var content = Matrixkey.Utility.Helper.ExportHelper.ExportToExcel(dt, true, rules, true, true, false);

            return File(content, "application/ms-excel", Url.Encode(fileName + ".xls"));
        }

        public FileResult Test7(string fileName = "测试DataSet带合并方式导出到Excel")
        {
            DataTable dt1 = GetDt(2, 10);
            DataTable dt2 = GetDt(10, 10);
            DataSet ds = new DataSet();
            ds.Tables.Add(dt1); ds.Tables.Add(dt2);

            List<MergerCellParam> rules = new List<MergerCellParam>() { 
                new MergerCellParam { RowStartIndex = 0, ColStartIndex = 1, RowEndIndex = 0, ColEndIndex = 3, Repeat = false },
                new MergerCellParam { RowStartIndex = 0, ColStartIndex = 5, RowEndIndex = 0, ColEndIndex = 6, Repeat = false },
                new MergerCellParam { RowStartIndex = 0, ColStartIndex = 7, RowEndIndex = 0, ColEndIndex = 9, Repeat = false }
            };

            var content = Matrixkey.Utility.Helper.ExportHelper.ExportToExcel(ds, rules, true, true, false);

            return File(content, "application/ms-excel", Url.Encode(fileName + ".xls"));
        }

        public FileResult Test8(int rc = 51, int cc = 10, string fileName = "测试DataTable带合并方式导出到Excel")
        {
            DataTable dt = GetDt(rc, cc);

            List<MergerCellParam> rules = new List<MergerCellParam>() { 
                new MergerCellParam { RowStartIndex = 1, ColStartIndex = 0, RowEndIndex = 1, ColEndIndex = 1, Repeat = true, IntervalRowCount = -1, IntervalColCount = 3 }
            };

            var content = Matrixkey.Utility.Helper.ExportHelper.ExportToExcel(dt, true, rules, true, true, false);

            return File(content, "application/ms-excel", Url.Encode(fileName + ".xls"));
        }

        public FileResult Test9(int rc = 51, int cc = 10, string fileName = "测试DataTable带合并方式导出到Excel")
        {
            DataTable dt = GetDt(rc, cc);

            List<MergerCellParam> rules = new List<MergerCellParam>() { 
                new MergerCellParam { RowStartIndex = 1, ColStartIndex = 0, RowEndIndex = 3, ColEndIndex = 1, Repeat = true, IntervalRowCount = 1, IntervalColCount = 1 }
            };

            var content = Matrixkey.Utility.Helper.ExportHelper.ExportToExcel(dt, true, rules, true, true, false);

            return File(content, "application/ms-excel", Url.Encode(fileName + ".xls"));
        }

        public FileResult Test10(int rc = 51, int cc = 10, string fileName = "测试DataTable带合并方式导出到Excel")
        {
            DataTable dt = GetDt(rc, cc);

            List<MergerCellParam> rules = new List<MergerCellParam>() { 
                new MergerCellParam { RowStartIndex = 1, ColStartIndex = 0, RowEndIndex = 3, ColEndIndex = 1, Repeat = true, IntervalRowCount = 0, IntervalColCount = 0 }
            };

            var content = Matrixkey.Utility.Helper.ExportHelper.ExportToExcel(dt, true, rules, true, true, false);

            return File(content, "application/ms-excel", Url.Encode(fileName + ".xls"));
        }

        public DataTable GetDt(int rowCount, int colCount)
        {
            DataTable dt = new DataTable();
            for (int c = 0; c < colCount; c++)
            {
                dt.Columns.Add("第" + c + "个列", typeof(string));
            }

            for (int r = 0; r < rowCount; r++)
            {
                DataRow row = dt.NewRow();
                row.SetValue("第" + r + "行数据", 0);
                dt.Rows.Add(row);
            }
            return dt;
        }
    }
}