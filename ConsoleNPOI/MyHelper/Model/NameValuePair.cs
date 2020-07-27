using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleNPOI.MyHelper.Model
{
    public class NameValuePair<T>
    {
        public string Name { get; set; }
        public T Value { get; set; }

        public NameValuePair()
        {
        }

        public NameValuePair(string name, T value)
        {
            this.Name = name;
            this.Value = value;
        }
    }
}
