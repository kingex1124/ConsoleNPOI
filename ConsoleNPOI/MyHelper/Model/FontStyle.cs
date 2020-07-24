using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleNPOI.MyHelper.Model
{
    public class FontStyle
    {
        /// <summary>
        /// 字體名稱
        /// </summary>
        public string FontName { get; set; }
        /// <summary>
        /// 字體大小
        /// </summary>
        public short? FontHeightInPoints { get; set; }

        public FontStyle()
        {
            FontName = "新細明體";
            FontHeightInPoints = 12;
        }
    }
}
