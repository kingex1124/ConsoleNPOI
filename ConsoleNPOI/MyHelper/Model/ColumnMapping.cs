using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleNPOI.MyHelper.Model
{
    public class ColumnMapping
    {
        public string ModelFieldName { get; set; }
        /// <summary>
        /// 若以範本初始化 Excel 則此欄可不填
        /// 欄位的值(表頭用的)
        /// </summary>
        public string ExcelColumnName { get; set; }
        public NpoiDataType DataType { get; set; }
        /// <summary>
        /// 如果是 String 則這個欄位不生效
        /// </summary>
        public string Format { get; set; }
    }
}
