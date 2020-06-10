using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace iCAT.Objects
{

    [ComVisible(true)]
    [Guid("18B999A2-FC40-42D6-A4CD-815AD147BC5E")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("iCAT.TNumeric")]
    public class TNumeric : IObject
    {
        public object Value { get; set; }

        public void sub()
        {
            throw new NotImplementedException();
        }

        public string Test()
        {
            return this.GetType().ToString();
        }
        public TNumeric() { }
        public TNumeric(string value)
        {
            this.Value = value;
        }
    }
}
