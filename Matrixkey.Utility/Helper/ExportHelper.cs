using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;
using System.Data;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

using Matrixkey.Utility.Extensions;

namespace Matrixkey.Utility.Helper
{
    public static class ExportHelper
    {
        #region 一组对外公开的导出数据到Excel的Api方法

        #region DataTable、DataSet导出到Excel的通用方法（支持复杂表头）

        /// <summary>
        /// 将DataTable导出到Excel
        /// </summary>
        /// <param name="data">要导出的DataTable对象</param>
        /// <param name="drawColumnName">画Excel时是否包含DataTable的列名</param>
        /// <param name="mergerRules">单元格合并规则集合，默认null</param>
        /// <param name="showBorder">每个单元格是否显示边框，默认是</param>
        /// <param name="autoFitColumn">是否自动计算列宽，默认是</param>
        /// <param name="ignoreMergeRuleWhenFitColumn">是否自动计算列宽时，是否忽略单元格合并规则，默认否。当不忽略时，当单元格包含在合并单元格集合中时，就不会自动计算列宽；当忽略时，则总是会自动计算列宽。该参数仅在<see cref="autoFitColumn"/>为true时有效。/></param>
        /// <returns></returns>
        public static byte[] ExportToExcel(DataTable data, bool drawColumnName, IEnumerable<MergerCellParam> mergerRules = null, bool showBorder = true, bool autoFitColumn = true, bool ignoreMergeRuleWhenFitColumn = false)
        {
            //创建excel
            HSSFWorkbook book = new HSSFWorkbook();

            //创建工作薄
            ISheet sheet1 = book.CreateSheet("Sheet1");

            int rowTotal = data.Rows.Count;
            int colTotal = data.Columns.Count;

            #region 解析合并规则

            IEnumerable<NPOI.SS.Util.CellRangeAddress> regions = ParseMergeRules(rowTotal, colTotal, mergerRules);

            #endregion

            #region 将DataTable画到Excel的工作薄中

            long[] cellByteLength = new long[colTotal]; //用来存储每个列的数据字节长度
            bool needMerge = !mergerRules.IsNullOrEmpty(); //是否需要合并单元格
            int rowIndex = 0;
            //每个单元格的样式，居中、边框
            var cellStyle0 = GetCellStyle(book, null, FillPattern.NoFill, null, null, HorizontalAlignment.Center, VerticalAlignment.Center, showBorder);

            if (drawColumnName)
            {
                //包含DataTable的列名
                IRow row = sheet1.CreateRow(rowIndex);
                row.HeightInPoints = 25;
                for (int k = 0; k < colTotal; k++)
                {
                    var col = data.Columns[k];
                    var cell = row.CreateCell(k);

                    string temp = col.ColumnName;
                    cell.SetCellValue(temp);
                    cell.CellStyle = cellStyle0;
                    if (autoFitColumn && (!needMerge || (!ignoreMergeRuleWhenFitColumn && !CheckInColumnMergeRegion(rowIndex, k, regions))))
                    {
                        TrySetSheetColumnWidth(sheet1, cellByteLength, k, temp.ToString());
                    }
                }
                rowIndex++;
            }

            for (int k = 0; k < rowTotal; k++)
            {
                IRow row = sheet1.CreateRow(rowIndex);
                row.HeightInPoints = 25;

                var dr = data.Rows[k];
                for (int m = 0; m < colTotal; m++)
                {
                    var dc = data.Columns[m];
                    var dv = Convert.ChangeType(dr[m], dc.DataType);

                    var cell = row.CreateCell(m);
                    if (dc.DataType.IsNumeric())
                    {
                        cell.SetCellType(CellType.Numeric);
                        if (dc.DataType == typeof(int))
                        {
                            int temp = dv.ToString().ToInt();
                            cell.SetCellValue(temp);

                            if (autoFitColumn && (!needMerge || (!ignoreMergeRuleWhenFitColumn && !CheckInColumnMergeRegion(rowIndex, m, regions))))
                            {
                                TrySetSheetColumnWidth(sheet1, cellByteLength, m, temp.ToString());
                            }
                        }
                        else if (dc.DataType == typeof(decimal) || dc.DataType == typeof(double))
                        {
                            double temp = (double)(dv.ToString().ToDecimal());
                            cell.SetCellValue(temp);

                            if (autoFitColumn && (!needMerge || (!ignoreMergeRuleWhenFitColumn && !CheckInColumnMergeRegion(rowIndex, m, regions))))
                            {
                                TrySetSheetColumnWidth(sheet1, cellByteLength, m, temp.ToString());
                            }
                        }
                        else
                        {
                            double temp = (double)(dv.ToString().ToDecimal());
                            cell.SetCellValue(temp);

                            if (autoFitColumn && (!needMerge || (!ignoreMergeRuleWhenFitColumn && !CheckInColumnMergeRegion(rowIndex, m, regions))))
                            {
                                TrySetSheetColumnWidth(sheet1, cellByteLength, m, temp.ToString());
                            }
                        }
                    }
                    else
                    {
                        cell.SetCellType(CellType.String);
                        string temp = dc.DataType == typeof(DateTime) ? ((DateTime)dv).ToString("yyyy-MM-dd") : dv.ToString();
                        cell.SetCellValue(temp);

                        if (autoFitColumn && (!needMerge || (!ignoreMergeRuleWhenFitColumn && !CheckInColumnMergeRegion(rowIndex, m, regions))))
                        {
                            TrySetSheetColumnWidth(sheet1, cellByteLength, m, temp);
                        }
                    }
                    cell.CellStyle = cellStyle0;
                }
                rowIndex++;
            }

            #endregion

            #region 合并单元格

            if (!regions.IsNullOrEmpty())
            {
                //设置合并后的单元格的对齐方式
                var cellStyle1 = GetCellStyle(book, null, FillPattern.NoFill, null, null, HorizontalAlignment.Center, VerticalAlignment.Center, showBorder);
                foreach (var item in regions)
                {
                    sheet1.AddMergedRegion(item);

                    for (int i = item.FirstRow; i <= item.LastRow; i++)
                    {
                        IRow row = NPOI.HSSF.Util.HSSFCellUtil.GetRow(i, sheet1 as HSSFSheet);
                        for (int j = item.FirstColumn; j <= item.LastColumn; j++)
                        {
                            ICell singleCell = NPOI.HSSF.Util.HSSFCellUtil.GetCell(row, (short)j);
                            singleCell.CellStyle = cellStyle1;
                        }
                    }
                }
            }

            #endregion

            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            book.Write(ms);
            ms.Seek(0, SeekOrigin.Begin);

            byte[] tt = ms.ToArray();
            ms.Close();
            return tt;
        }

        /// <summary>
        /// 将DataSet导出到Excel。使用DataSet是考虑到复杂列头部分和数据体部分的DataTable的数据类型是不同的，可以把复杂列头、数据体分别以DataTable格式放入DataSet中。
        /// </summary>
        /// <param name="data">要导出的DataSet对象</param>
        /// <param name="mergerRules">单元格合并规则集合，默认null</param>
        /// <param name="showBorder">每个单元格是否显示边框，默认是</param>
        /// <param name="autoFitColumn">是否自动计算列宽，默认是</param>
        /// <param name="ignoreMergeRuleWhenFitColumn">是否自动计算列宽时，是否忽略单元格合并规则，默认否。当不忽略时，当单元格包含在合并单元格集合中时，就不会自动计算列宽；当忽略时，则总是会自动计算列宽。该参数仅在<see cref="autoFitColumn"/>为true时有效。/></param>
        /// <returns></returns>
        public static byte[] ExportToExcel(DataSet data, IEnumerable<MergerCellParam> mergerRules = null, bool showBorder = true, bool autoFitColumn = true, bool ignoreMergeRuleWhenFitColumn = false)
        {
            //创建excel
            HSSFWorkbook book = new HSSFWorkbook();

            //创建工作薄
            ISheet sheet1 = book.CreateSheet("Sheet1");

            int dtTotal = data.Tables.Count;

            #region 解析合并规则

            int rowsCount = 0, colsCount = 0;
            for (int n = 0; n < dtTotal; n++)
            {
                var dt = data.Tables[n];
                rowsCount += dt.Rows.Count;
                colsCount = Math.Max(colsCount, dt.Columns.Count);
            }

            IEnumerable<NPOI.SS.Util.CellRangeAddress> regions = ParseMergeRules(rowsCount, colsCount, mergerRules);

            #endregion

            #region 将DataSet中每个DataTable画到Excel的工作薄中

            if (dtTotal > 0)
            {
                int rowIndex = 0;
                //找出各DataTable中列最多的数目
                int maxColCount = 0;
                for (int n = 0; n < dtTotal; n++)
                {
                    var dt = data.Tables[n];
                    int colTotal = dt.Columns.Count;
                    if (colTotal > maxColCount) { maxColCount = colTotal; }
                }
                long[] cellByteLength = new long[maxColCount]; //用来存储每个列的数据字节长度
                bool needMerge = !mergerRules.IsNullOrEmpty(); //是否需要合并单元格
                //每个单元格的样式，居中、边框
                var cellStyle0 = GetCellStyle(book, null, FillPattern.NoFill, null, null, HorizontalAlignment.Center, VerticalAlignment.Center, showBorder);
                for (int n = 0; n < dtTotal; n++)
                {
                    var dt = data.Tables[n];

                    int rowTotal = dt.Rows.Count;
                    int colTotal = dt.Columns.Count;
                    for (int k = 0; k < rowTotal; k++)
                    {
                        IRow row = sheet1.CreateRow(rowIndex);
                        row.HeightInPoints = 25;

                        var dr = dt.Rows[k];
                        for (int m = 0; m < colTotal; m++)
                        {
                            var dc = dt.Columns[m];
                            var dv = Convert.ChangeType(dr[m], dc.DataType);

                            var cell = row.CreateCell(m);
                            if (dc.DataType.IsNumeric())
                            {
                                cell.SetCellType(CellType.Numeric);
                                if (dc.DataType == typeof(int))
                                {
                                    int temp = dv.ToString().ToInt();
                                    cell.SetCellValue(temp);

                                    if (autoFitColumn && (!needMerge || (!ignoreMergeRuleWhenFitColumn && !CheckInColumnMergeRegion(rowIndex, m, regions))))
                                    {
                                        TrySetSheetColumnWidth(sheet1, cellByteLength, m, temp.ToString());
                                    }
                                }
                                else if (dc.DataType == typeof(decimal) || dc.DataType == typeof(double))
                                {
                                    double temp = (double)(dv.ToString().ToDecimal());
                                    cell.SetCellValue(temp);

                                    if (autoFitColumn && (!needMerge || (!ignoreMergeRuleWhenFitColumn && !CheckInColumnMergeRegion(rowIndex, m, regions))))
                                    {
                                        TrySetSheetColumnWidth(sheet1, cellByteLength, m, temp.ToString());
                                    }
                                }
                                else
                                {
                                    double temp = (double)(dv.ToString().ToDecimal());
                                    cell.SetCellValue(temp);

                                    if (autoFitColumn && (!needMerge || (!ignoreMergeRuleWhenFitColumn && !CheckInColumnMergeRegion(rowIndex, m, regions))))
                                    {
                                        TrySetSheetColumnWidth(sheet1, cellByteLength, m, temp.ToString());
                                    }
                                }
                            }
                            else
                            {
                                cell.SetCellType(CellType.String);
                                string temp = dc.DataType == typeof(DateTime) ? ((DateTime)dv).ToString("yyyy-MM-dd") : dv.ToString();
                                cell.SetCellValue(temp);

                                if (autoFitColumn && (!needMerge || (!ignoreMergeRuleWhenFitColumn && !CheckInColumnMergeRegion(rowIndex, m, regions))))
                                {
                                    TrySetSheetColumnWidth(sheet1, cellByteLength, m, temp);
                                }
                            }
                            cell.CellStyle = cellStyle0;
                        }
                        rowIndex++;
                    }
                }
            }

            #endregion

            #region 合并单元格

            if (!regions.IsNullOrEmpty())
            {
                //设置合并后的单元格的对齐方式
                var cellStyle1 = GetCellStyle(book, null, FillPattern.NoFill, null, null, HorizontalAlignment.Center, VerticalAlignment.Center, showBorder);
                foreach (var item in regions)
                {
                    sheet1.AddMergedRegion(item);

                    for (int i = item.FirstRow; i <= item.LastRow; i++)
                    {
                        IRow row = NPOI.HSSF.Util.HSSFCellUtil.GetRow(i, sheet1 as HSSFSheet);
                        for (int j = item.FirstColumn; j <= item.LastColumn; j++)
                        {
                            ICell singleCell = NPOI.HSSF.Util.HSSFCellUtil.GetCell(row, (short)j);
                            singleCell.CellStyle = cellStyle1;
                        }
                    }
                }
            }

            #endregion

            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            book.Write(ms);
            ms.Seek(0, SeekOrigin.Begin);

            byte[] tt = ms.ToArray();
            ms.Close();
            return tt;
        }

        #endregion

        #region 可枚举集合导出到Excel的方法（不支持复杂表头）

        /// <summary>
        /// 将可枚举的数据集合导出到Excel
        /// </summary>
        /// <typeparam name="T">要导出的数据的数据类型</typeparam>
        /// <param name="data">要导出的数据</param>
        /// <param name="exportPropertities">要导出的属性名称集合，请保证属性名称在<see cref="data"/>中存在。</param>
        /// <param name="sheetSize">每个工作薄显示多少条数据。若设置为0或小于0，则表示在一张工作表中显示所有数据</param>
        /// <param name="userDisplayName">是否以属性的Display特性值来当做列头</param>
        public static byte[] ExportToExcel<T>(IEnumerable<T> data, IEnumerable<string> exportPropertities, int sheetSize, bool userDisplayName)
        {
            using (System.IO.MemoryStream ms = BuildExeclStruct(data, exportPropertities, sheetSize, userDisplayName))
            {
                byte[] tt = ms.ToArray();

                return tt;
            }
        }

        #endregion

        #endregion

        #region 导出easyui-datagrid时，对标题列集合、数据行集合处理的Api方法

        /// <summary>
        /// 处理标题列集合、数据行集合，返回完整规则的DataTable对象，并将解析所得的单元格合并规则以输出参数形式返回
        /// </summary>
        /// <param name="columns">标题列集合</param>
        /// <param name="data">数据行集合</param>
        /// <param name="rules">输出参数，根据标题列集合解析所得的单元格合并规则</param>
        /// <returns>完整规则的DataTable对象</returns>
        public static DataTable ParseColumnsAndRows(DataTable[] columns, DataTable data, out IEnumerable<Matrixkey.Utility.Helper.MergerCellParam> rules)
        {
            #region 组建目标DataTable结构

            int columnsRowCount = columns.Length, columnsColCount = columns.Length > 0 ? columns[0].Rows.Count : 0;
            System.Data.DataTable dt = new System.Data.DataTable();
            for (int i = 0; i < columnsColCount; i++)
            {
                System.Data.DataRow mapRow = null;
                foreach (var item in columns)
                {
                    var tempRow = item.Rows[i];
                    if (tempRow["field"] != null && !string.IsNullOrWhiteSpace(tempRow["field"].ToString()))
                    {
                        mapRow = tempRow; break;
                    }
                }
                if (mapRow != null) { dt.Columns.Add(mapRow["field"].ToString(), typeof(string)); }
                else { dt.Columns.Add("列" + i, typeof(string)); }
            }

            #endregion

            #region 填充标题列数据

            foreach (var item in columns)
            {
                System.Data.DataRow newRow = dt.NewRow();
                for (int i = 0; i < columnsColCount; i++)
                {
                    newRow[i] = item.Rows[i]["title"];
                }
                dt.Rows.Add(newRow);
            }

            #endregion

            #region 标题列合并规则计算

            rules = new List<Matrixkey.Utility.Helper.MergerCellParam>();
            Dictionary<string, Matrixkey.Utility.Helper.MergerCellParam> dicRules = new Dictionary<string, Matrixkey.Utility.Helper.MergerCellParam>();

            for (int k = 0; k < columns.Length; k++)
            {
                System.Data.DataTable itemFc = columns[k];
                for (int i = 0; i < itemFc.Rows.Count; i++)
                {
                    System.Data.DataRow itemRow = itemFc.Rows[i];
                    if (itemRow["colspan"].ToString() != "1")
                    {
                        string key = itemRow["field"].ToString() + "_" + itemRow["title"].ToString();
                        Matrixkey.Utility.Helper.MergerCellParam tempRule = null;
                        if (dicRules.TryGetValue(key, out tempRule))
                        {
                            tempRule.RowEndIndex = k;
                            tempRule.ColEndIndex = i;
                        }
                        else
                        {
                            tempRule = new Matrixkey.Utility.Helper.MergerCellParam();
                            tempRule.RowStartIndex = k;
                            tempRule.ColStartIndex = i;
                            tempRule.Repeat = false;

                            dicRules.Add(key, tempRule);
                        }
                    }
                    if (itemRow["rowspan"].ToString() != "1")
                    {
                        string key = itemRow["field"].ToString() + "_" + itemRow["title"].ToString();
                        Matrixkey.Utility.Helper.MergerCellParam tempRule = null;
                        if (dicRules.TryGetValue(key, out tempRule))
                        {
                            tempRule.RowEndIndex = k;
                            tempRule.ColEndIndex = i;
                        }
                        else
                        {
                            tempRule = new Matrixkey.Utility.Helper.MergerCellParam();
                            tempRule.RowStartIndex = k;
                            tempRule.ColStartIndex = i;
                            tempRule.Repeat = false;

                            dicRules.Add(key, tempRule);
                        }
                    }
                }
            }

            rules = dicRules.Select(s => s.Value);

            #endregion

            #region 填充主体数据

            for (int i = 0; i < data.Rows.Count; i++)
            {
                System.Data.DataRow newRow = dt.NewRow();
                for (int j = 0; j < data.Columns.Count; j++)
                {
                    string columnName = data.Columns[j].ColumnName; //Field名
                    if (!dt.Columns.Contains(columnName)) { continue; }
                    object cellObject = data.Rows[i][j];//单元格内容

                    newRow[columnName] = cellObject;
                }
                dt.Rows.Add(newRow);
            }

            #endregion

            return dt;
        }

        #endregion

        #region 内部使用的构建excel有关的定向方法

        /// <summary>
        /// 组建excel结构，返回流对象
        /// </summary>
        /// <typeparam name="T">要导出的数据的数据类型</typeparam>
        /// <param name="data">要导出的数据</param>
        /// <param name="exportPropertities">要导出的属性名称集合，请保证属性名称在<see cref="data"/>中存在。</param>
        /// <param name="sheetSize">每个工作薄显示多少条数据。若设置为0或小于0，则表示在一张工作表中显示所有数据</param>
        /// <param name="userDisplayName">是否以属性的Display特性值来当做列头</param>
        /// <returns></returns>
        private static System.IO.MemoryStream BuildExeclStruct<T>(IEnumerable<T> data, IEnumerable<string> exportPropertities, int sheetSize, bool userDisplayName)
        {
            int dataLen = data.Count();
            if (dataLen == 0) { throw new ArgumentException("要导出的数据对象不能为空。"); }

            //创建Excel文件的对象
            HSSFWorkbook book = new HSSFWorkbook();
            AddBookPropertyInfo(book);

            //标题行样式
            ICellStyle headStyle = book.CreateCellStyle();
            var font = book.CreateFont();
            font.FontHeightInPoints = 20;
            font.Boldweight = 700;
            headStyle.SetFont(font);

            PropertyInfo[] properties = typeof(T).GetProperties();
            var props = exportPropertities.Count() > 0 ? properties.Where(w => exportPropertities.Contains(w.Name)).ToArray() : properties;
            long[] headCellByteLength = new long[props.Length];

            int sheetCount = 1;
            if (sheetSize > 0)
            {
                sheetCount = dataLen / sheetSize;
                if (dataLen % sheetSize > 0)
                {
                    sheetCount++;
                }
            }

            for (int k = 0; k < sheetCount; k++)
            {
                ISheet sheetChild = book.CreateSheet("Sheet" + (k + 1));
                if (k == 0)
                {
                    BuildExcelHead(sheetChild, headStyle, props, userDisplayName, out headCellByteLength);
                }
                else
                {
                    BuildExcelHead(sheetChild, headStyle, props, userDisplayName, headCellByteLength);
                }
                if (sheetSize > 0)
                {
                    AppendItemData(sheetChild, props, headCellByteLength, data.Skip(k * sheetSize).Take(sheetSize));
                }
                else
                {
                    AppendItemData(sheetChild, props, headCellByteLength, data);
                }
            }

            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            book.Write(ms);
            ms.Seek(0, SeekOrigin.Begin);

            return ms;
        }

        /// <summary>
        /// 组装Excel标题列
        /// </summary>
        /// <param name="excelSheet">工作薄对象</param>
        /// <param name="headRowStyle">标题行的样式</param>
        /// <param name="properties">要导出的数据的类型的属性集合</param>
        /// <param name="userDisplayName">是否以属性的Display特性值来当做列头</param>
        /// <param name="headCellByteLength">输出参数，存储了标题行每列title字节长度的数组</param>
        private static void BuildExcelHead(ISheet excelSheet, ICellStyle headRowStyle, PropertyInfo[] properties, bool userDisplayName, out long[] headCellByteLength)
        {
            int len = properties.Length;
            headCellByteLength = new long[len];

            if (len == 0)
            { return; }

            IRow headRow = excelSheet.CreateRow(0);
            headRow.HeightInPoints = 25;
            if (headRowStyle != null)
            {
                headRow.RowStyle = headRowStyle;
            }

            string title = string.Empty;
            for (int k = 0; k < len; k++)
            {
                if (userDisplayName)
                {
                    var attr = properties[k].ToDescription();
                    title = string.IsNullOrWhiteSpace(attr) ? properties[k].Name : attr;
                }
                else
                {
                    title = properties[k].Name;
                }
                headCellByteLength[k] = title.GetByteLength();
                ICell headCellItem = headRow.CreateCell(k);
                headCellItem.SetCellValue(title);
                excelSheet.SetColumnWidth(k, 256 * ((int)headCellByteLength[k] + 1));
            }
        }

        /// <summary>
        /// 组装Excel标题列
        /// </summary>
        /// <param name="excelSheet">工作薄对象</param>
        /// <param name="headRowStyle">标题行的样式</param>
        /// <param name="properties">要导出的数据的类型的属性集合</param>
        /// <param name="userDisplayName">是否以属性的Display特性值来当做列头</param>
        /// <param name="headCellByteLength">存储了标题行每列title字节长度的数组</param>
        private static void BuildExcelHead(ISheet excelSheet, ICellStyle headRowStyle, PropertyInfo[] properties, bool userDisplayName, long[] headCellByteLength)
        {
            int len = properties.Length;
            if (len == 0)
            { return; }

            IRow headRow = excelSheet.CreateRow(0);
            headRow.HeightInPoints = 25;
            if (headRowStyle != null)
            {
                headRow.RowStyle = headRowStyle;
            }

            string title = string.Empty;
            for (int k = 0; k < len; k++)
            {
                if (userDisplayName)
                {
                    var attr = properties[k].ToDescription();
                    title = string.IsNullOrWhiteSpace(attr) ? properties[k].Name : attr;
                }
                else
                {
                    title = properties[k].Name;
                }
                ICell headCellItem = headRow.CreateCell(k);
                headCellItem.SetCellValue(title);
                excelSheet.SetColumnWidth(k, 256 * ((int)headCellByteLength[k] + 1));
            }
        }

        /// <summary>
        /// 组装数据行
        /// </summary>
        /// <typeparam name="T">要导出的数据的数据类型</typeparam>
        /// <param name="excelSheet">工作薄对象</param>
        /// <param name="properties">要导出的数据的类型的属性集合</param>
        /// <param name="headCellByteLength">存储了标题行每列title字节长度的数组，用来比较同列下的数据长度，最终决定列宽</param>
        /// <param name="data">要导出的数据</param>
        private static void AppendItemData<T>(ISheet excelSheet, PropertyInfo[] properties, long[] headCellByteLength, IEnumerable<T> data)
        {
            int propertyLen = properties.Length;
            int dataLen = data.Count();
            for (int i = 0; i < dataLen; i++)
            {
                IRow rowtemp = excelSheet.CreateRow(i + 1);
                for (int k = 0; k < propertyLen; k++)
                {
                    var obj = properties[k].GetValue(data.ElementAt(i), null);
                    string val = obj == null ? "" : obj.ToString();
                    rowtemp.CreateCell(k).SetCellValue(val);
                    if (val.GetByteLength() > headCellByteLength[k])
                    {
                        headCellByteLength[k] = val.GetByteLength();
                        excelSheet.SetColumnWidth(k, 256 * ((int)headCellByteLength[k] + 1));
                    }
                }
            }
        }

        #endregion

        #region 内部使用的构建excel有关的通用方法

        /// <summary>
        /// 添加Excel文件属性信息
        /// </summary>
        /// <param name="book"></param>
        private static void AddBookPropertyInfo(HSSFWorkbook book)
        {
            NPOI.HPSF.DocumentSummaryInformation dsi = NPOI.HPSF.PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "PM";
            book.DocumentSummaryInformation = dsi;

            NPOI.HPSF.SummaryInformation si = NPOI.HPSF.PropertySetFactory.CreateSummaryInformation();
            si.Author = "品茗造价"; //填加xls文件作者信息
            si.ApplicationName = "NPOI"; //填加xls文件创建程序信息
            si.LastAuthor = "品茗造价"; //填加xls文件最后保存者信息
            si.Comments = "作者信息"; //填加xls文件作者信息
            si.Title = "导出的Excel"; //填加xls文件标题信息
            si.Subject = "数据导出";//填加文件主题信息
            si.CreateDateTime = System.DateTime.Now;
            book.SummaryInformation = si;
        }

        /// <summary>
        /// 获取单元格样式
        /// </summary>
        /// <param name="hssfworkbook">Excel操作类</param>
        /// <param name="font">单元格字体</param>
        /// <param name="fillPattern">图案样式</param>
        /// <param name="fillForegroundColor">图案的颜色</param>
        /// <param name="fillBackgroundColor">单元格背景</param>
        /// <param name="ha">垂直对齐方式</param>
        /// <param name="va">垂直对齐方式</param>
        /// <param name="showBorder">是否显示边框</param>
        /// <returns></returns>
        private static ICellStyle GetCellStyle(HSSFWorkbook hssfworkbook, IFont font, FillPattern fillPattern, NPOI.HSSF.Util.HSSFColor fillForegroundColor, NPOI.HSSF.Util.HSSFColor fillBackgroundColor, HorizontalAlignment ha, VerticalAlignment va, bool showBorder)
        {
            ICellStyle cellstyle = hssfworkbook.CreateCellStyle();
            cellstyle.FillPattern = fillPattern;
            cellstyle.Alignment = ha;
            cellstyle.VerticalAlignment = va;
            if (fillForegroundColor != null)
            {
                cellstyle.FillForegroundColor = fillForegroundColor.Indexed;
            }
            if (fillBackgroundColor != null)
            {
                cellstyle.FillBackgroundColor = fillBackgroundColor.Indexed;
            }
            if (font != null)
            {
                cellstyle.SetFont(font);
            }
            if (showBorder)
            {
                cellstyle.BorderBottom = BorderStyle.Thin;
                cellstyle.BorderLeft = BorderStyle.Thin;
                cellstyle.BorderRight = BorderStyle.Thin;
                cellstyle.BorderTop = BorderStyle.Thin;
                cellstyle.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
                cellstyle.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
                cellstyle.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
                cellstyle.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            }
            return cellstyle;
        }

        /// <summary>
        /// 获取字体样式
        /// </summary>
        /// <param name="hssfworkbook">Excel操作类</param>
        /// <param name="fontname">字体名</param>
        /// <param name="fontcolor">字体颜色</param>
        /// <param name="fontsize">字体大小</param>
        /// <returns></returns>
        public static IFont GetFontStyle(HSSFWorkbook hssfworkbook, string fontfamily, NPOI.HSSF.Util.HSSFColor fontcolor, int fontsize)
        {
            IFont font1 = hssfworkbook.CreateFont();
            if (string.IsNullOrEmpty(fontfamily))
            {
                font1.FontName = fontfamily;
            }
            if (fontcolor != null)
            {
                font1.Color = fontcolor.Indexed;
            }
            font1.IsItalic = true;
            font1.FontHeightInPoints = (short)fontsize;
            return font1;
        }

        /// <summary>
        /// 判定指定行索引指定列索引的单元格是否在跨列的合并区域中
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <param name="colIndex">列索引</param>
        /// <param name="rules">合并规则集合</param>
        /// <returns></returns>
        private static bool CheckInColumnMergeRegion(int rowIndex, int colIndex, IEnumerable<NPOI.SS.Util.CellRangeAddress> regions)
        {
            bool result = false;
            foreach (var item in regions)
            {
                if (item.FirstColumn == item.LastColumn) { continue; }
                if (rowIndex >= item.FirstRow && rowIndex <= item.LastRow && colIndex >= item.FirstColumn && colIndex <= item.LastColumn)
                {
                    result = true;
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// 尝试设置工作薄指定列的宽度
        /// </summary>
        /// <param name="sheet">工作薄对象</param>
        /// <param name="his">各列内容最大字节长度的集合</param>
        /// <param name="colIndex">列索引</param>
        /// <param name="val">单元格内容</param>
        private static void TrySetSheetColumnWidth(ISheet sheet, long[] his, int colIndex, string val)
        {
            long cellValueByteLength = val.GetByteLength();
            if (cellValueByteLength > his[colIndex])
            {
                his[colIndex] = cellValueByteLength;
                sheet.SetColumnWidth(colIndex, 256 * ((int)his[colIndex] + 1));
            }
        }

        /// <summary>
        /// 解析单元格合并规则，返回要合并的区域集合
        /// </summary>
        /// <param name="rowsCount">数据行总数</param>
        /// <param name="colsCount">数据列总数</param>
        /// <param name="rules">合并规则集合</param>
        /// <returns></returns>
        private static IEnumerable<NPOI.SS.Util.CellRangeAddress> ParseMergeRules(int rowsCount, int colsCount, IEnumerable<MergerCellParam> rules)
        {
            List<NPOI.SS.Util.CellRangeAddress> regions = new List<NPOI.SS.Util.CellRangeAddress>();
            if (rules.IsNullOrEmpty()) { return regions; }
            foreach (var item in rules)
            {
                NPOI.SS.Util.CellRangeAddress region = new NPOI.SS.Util.CellRangeAddress(item.RowStartIndex, item.RowEndIndex, item.ColStartIndex, item.ColEndIndex);
                if (item.Repeat && !(item.IntervalRowCount < 0 && item.IntervalColCount < 0))
                {

                    // 行间隔 大于0，列间隔 小于0 A1
                    // 行间隔 大于0，列间隔 等于0 C => A1
                    // 行间隔 等于0，列间隔 大于0 D => B1
                    // 行间隔 等于0，列间隔 小于0 A2
                    // 行间隔 小于0，列间隔 大于0 B1
                    // 行间隔 小于0，列间隔 等于0 B2
                    if ((item.IntervalRowCount > 0 && item.IntervalColCount > 0) || (item.IntervalRowCount == 0 && item.IntervalColCount == 0))
                    {
                        // 都大于0 或 都等于0
                        int intervalRow = item.RowEndIndex - item.RowStartIndex + 1;//合并区域的跨行数
                        int intervalCol = item.ColEndIndex - item.ColStartIndex + 1;//合并区域的跨列数
                        for (int kr = item.RowEndIndex + 1 + item.IntervalRowCount, kc = item.ColEndIndex + 1 + item.IntervalColCount; kr < rowsCount + 1 && kc < colsCount + 1; kr += (intervalRow + item.IntervalRowCount), kc += (intervalCol + item.IntervalColCount))
                        {
                            regions.Add(new NPOI.SS.Util.CellRangeAddress(kr, kr + intervalRow - 1, kc, kc + intervalCol - 1));
                        }
                    }
                    else
                    {
                        if (item.IntervalRowCount > 0 && item.IntervalColCount == 0)
                        {
                            // 行间隔 大于0，列间隔 等于0 C => A1
                            item.IntervalColCount = -1;
                        }
                        else if (item.IntervalRowCount == 0 && item.IntervalColCount > 0)
                        {
                            // 行间隔 等于0，列间隔 大于0 D => B1
                            item.IntervalRowCount = -1;
                        }

                        if (item.IntervalColCount < 0)
                        {
                            // 行间隔 大于0，列间隔 小于0 A1
                            // 行间隔 等于0，列间隔 小于0 A2
                            //间隔列小于0，表示在“主体规则的列范围”中重复合并单元格
                            int intervalRow = item.RowEndIndex - item.RowStartIndex + 1;//合并区域的跨行数
                            //item.IntervalRowCount;//重复合并区域之间的跨行数
                            for (int k = item.RowEndIndex + 1 + item.IntervalRowCount; k < rowsCount + 1; k += (intervalRow + item.IntervalRowCount))
                            {
                                regions.Add(new NPOI.SS.Util.CellRangeAddress(k, k + intervalRow - 1, item.ColStartIndex, item.ColEndIndex));
                            }
                        }
                        else if (item.IntervalRowCount < 0)
                        {
                            // 行间隔 小于0，列间隔 大于0 B1
                            // 行间隔 小于0，列间隔 等于0 B2
                            //间隔行小于0，表示在“主体规则的行范围”中重复合并单元格
                            int intervalCol = item.ColEndIndex - item.ColStartIndex + 1;//合并区域的跨列数
                            //item.IntervalColCount;//重复合并区域之间的跨列数
                            for (int k = item.ColEndIndex + 1 + item.IntervalColCount; k < colsCount + 1; k += (intervalCol + item.IntervalColCount))
                            {
                                regions.Add(new NPOI.SS.Util.CellRangeAddress(item.RowStartIndex, item.RowEndIndex, k, k + intervalCol - 1));
                            }
                        }
                    }
                }
                regions.Add(region);
            }

            return regions;
        }

        //设置单元格内显示数据的格式
        //        ICell cell = row.CreateCell(1);
        //ICellStyle cellStyleNum = Excel.GetICellStyle(book);
        //IDataFormat formatNum = book.CreateDataFormat();
        //cellStyleNum.DataFormat = formatNum.GetFormat("0.00E+00");//设置单元格的格式为科学计数法cell.CellStyle = cellStyleNum;

        #endregion
    }

    /// <summary>
    /// 单元格合并规则模型
    /// </summary>
    public class MergerCellParam
    {
        /// <summary>
        /// 开始行的索引
        /// </summary>
        public int RowStartIndex { get; set; }

        /// <summary>
        /// 开始列的索引
        /// </summary>
        public int ColStartIndex { get; set; }

        /// <summary>
        /// 结束行的索引
        /// </summary>
        public int RowEndIndex { get; set; }

        /// <summary>
        /// 结束列的索引
        /// </summary>
        public int ColEndIndex { get; set; }

        /// <summary>
        /// 是否重复
        /// </summary>
        public bool Repeat { get; set; }

        /// <summary>
        /// 间隔行数，小于0则表示在 主体规则的行范围内重复合并
        /// </summary>
        public int IntervalRowCount { get; set; }

        /// <summary>
        /// 间隔列数，小于0则表示在 主体规则的列范围内重复合并
        /// </summary>
        public int IntervalColCount { get; set; }
    }
}
