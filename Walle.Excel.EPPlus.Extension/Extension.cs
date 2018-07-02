using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Walle.Excel.Core;
using Walle.Excel.Core.Attributes;

namespace Walle.Excel.EPPlus.Extension
{
    public static class Extension
    {
        public static void ToExcel<T>(this IEnumerable<T> list, string filePath, string sheetName = "Sheet1") where T : class, ISheetRow, new()
        {
            FileInfo file = new FileInfo(filePath);
            if (file.Exists)
            {
                file.Delete();
                file = new FileInfo(filePath);
            }
            using (ExcelPackage package = new ExcelPackage(file))
            {
                //创建sheet
                package.Workbook.Worksheets.Add(sheetName).FromList<T>(list);
                package.Save();
            }
        }

        public static byte[] ToExcelContent<T>(this IEnumerable<T> list, string sheetName = "Sheet1") where T : class, ISheetRow, new()
        {
            using (var ms = new MemoryStream())
            {
                ExcelPackage package = new ExcelPackage(ms);

                package.Workbook.Worksheets.Add(sheetName).FromList(list);
                package.Save();
                return ms.ToArray();
            }
        }

        #region private
        private static ExcelWorksheet FromList<T>(this ExcelWorksheet sheet, IEnumerable<T> list) where T : class, ISheetRow, new()
        {
            sheet.CreateTitleRow<T>();
            int index = 2;
            foreach (var item in list)
            {
                sheet.CreateContentRow(item, index);
                index++;
            }
            return sheet;
        }

        private static void CreateTitleRow<T>(this ExcelWorksheet sheet) where T : class, ISheetRow, new()
        {
            var rowIndex = 1;
            var type = typeof(T);
            var properties = type.GetProperties();
            List<Column> columns = new List<Column>();
            foreach (var property in properties)
            {
                var attributes = property.GetCustomAttributes(typeof(Column), true);
                if (attributes != null && attributes.Length > 0)
                {
                    var attr = attributes[attributes.Length - 1];
                    if (attr is Column)
                    {
                        var col = attr as Column;
                        if (col.Ignore)
                        {
                            continue;
                        }
                        columns.Add(col);
                    }
                }
            }
            sheet.InsertRow(rowIndex, 1);
            var titles = columns.OrderBy(p => p.Index).Select(p => p.Title);
            var index = 1;
            foreach (var title in titles)
            {
                sheet.SetValue(rowIndex, index, title);
                index++;
            }
            sheet.Cells.AutoFitColumns();
        }

        private static void CreateContentRow<T>(this ExcelWorksheet sheet, T item, int index) where T : class, ISheetRow, new()
        {
            var type = typeof(T);
            var properties = type.GetProperties();
            List<Column> columns = new List<Column>();
            foreach (var property in properties)
            {
                var attributes = property.GetCustomAttributes(typeof(Column), true);
                if (attributes != null && attributes.Length > 0)
                {
                    var attr = attributes[attributes.Length - 1];
                    if (attr is Column)
                    {
                        var col = attr as Column;
                        if (col.Ignore)
                        {
                            continue;
                        }
                        var value = property.GetValue(item);
                        col.SetValue(value);
                        columns.Add(col);
                    }
                }
            }
            sheet.InsertRow(index, 1);
            var values = columns.OrderBy(p => p.Index).Select(p => p.Value);
            var colIndex = 1;
            foreach (var title in values)
            {
                sheet.SetValue(index, colIndex, title);
                colIndex++;
            }
            sheet.Cells.AutoFitColumns();
        }

        private static void SetValue(this Column col, object value)
        {
            if (value is null || value == null)
            {
                value = col.DefaultValue;
            }
            if (value is string && string.IsNullOrWhiteSpace(value.ToString()) && col.DefaultValue != null)
            {
                col.Value = col.DefaultValue.ToString();
                return;
            }
            if (value is DateTime)
            {
                col.Value = ((DateTime)value).ToString(col.DateFormat);
                return;
            }
            if (value is DateTime?)
            {
                var v = (value as DateTime?);
                if (v.HasValue)
                {
                    col.Value = v.Value.ToString(col.DateFormat);
                    return;
                }
                else
                {
                    col.Value = string.Empty;
                    return;
                }
            }
            col.Value = value.ToString();
        }
        #endregion

    }
}
