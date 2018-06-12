using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Walle.Excel.Core;
using Walle.Excel.Core.Attributes;

namespace Walle.Excel.NPOI.Extension
{
    public static class Extension
    {
        public static void ToExcel<T>(this IEnumerable<T> list, string filePath, string sheetName = "Sheet1") where T : class, ISheetRow, new()
        {
            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                var book = new XSSFWorkbook();
                var sheet = book.CreateSheet(sheetName);
                sheet = sheet.FromList<T>(list);
                book.Write(fs);
            }
        }

        public static byte[] ToExcelContent<T>(this IEnumerable<T> list, string sheetName = "Sheet1") where T : class, ISheetRow, new()
        {
            using (var ms = new MemoryStream())
            {
                var book = new XSSFWorkbook();
                var sheet = book.CreateSheet(sheetName).FromList(list);
                book.Write(ms);
                return ms.ToArray();
            }
        }

        #region private
        private static ISheet FromList<T>(this ISheet sheet, IEnumerable<T> list) where T : class, ISheetRow, new()
        {
            sheet.CreateTitleRow<T>();
            int index = 1;
            foreach (var item in list)
            {
                sheet.CreateContentRow(item, index);
                index++;
            }
            return sheet;
        }
        private static void CreateTitleRow<T>(this ISheet sheet) where T : class, ISheetRow, new()
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
                        columns.Add(col);
                    }
                }
            }
            var row = sheet.CreateRow(0);
            var titles = columns.OrderBy(p => p.Index).Select(p => p.Title);
            var index = 0;
            foreach (var title in titles)
            {
                var cell = row.CreateCell(index);
                cell.SetCellValue(title);
                index++;
            }
        }

        private static void CreateContentRow<T>(this ISheet sheet, T item, int index) where T : class, ISheetRow, new()
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
            var row = sheet.CreateRow(index);
            var values = columns.OrderBy(p => p.Index).Select(p => p.Value);
            var colIndex = 0;
            foreach (var title in values)
            {
                var cell = row.CreateCell(colIndex);
                cell.SetCellValue(title);
                colIndex++;
            }
        }

        private static void SetValue(this Column col, object value)
        {
            if (value is null || value == null)
            {
                value = col.DefaultValue;
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
