using System;
using System.Collections.Generic;
using System.Text;

namespace Walle.NPOI.Extension.Attributes
{
    /// <summary>
    /// 
    /// </summary>
    public class Column : Attribute
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Index">列索引</param>
        /// <param name="Title">列名</param>
        /// <param name="DateFormat">日期格式</param>
        /// <param name="Ignore">是否忽略本行</param>
        public Column(int Index = 0, string Title = "", string DateFormat = "yyyy-MM-dd HH:mm:ss", object DefaultValue = null, bool Ignore = false)
        {
            this.Index = Index;
            this.Title = Title;
            this.DateFormat = DateFormat;
            this.Ignore = Ignore;
            this.DefaultValue = DefaultValue;
        }

        public string Title { get; set; } = string.Empty;
        public string DateFormat { get; set; } = "yyyy-MM-dd HH:mm:ss";
        public int Index { get; set; } = 0;
        public bool Ignore { get; set; } = false;
        public object DefaultValue { get; set; } = null;
        internal string Value { get; set; } = string.Empty;
    }
}
