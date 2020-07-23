using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleNPOI.Helper.Model
{
    public class FontStyle
    {
        public string FontName { get; set; }
        public short? FontHeightInPoints { get; set; }

        public FontStyle()
        {
            FontName = null;
            FontHeightInPoints = null;
        }
    }
}
