using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace iCAT.Interopt
{
    [ComVisible(true)]
    [Guid("18727C24-0875-4DEE-80BB-38BD6D4457A7"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface INumeric
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string Test();
        object Value { get; set; }
        void sub();
    }

    [ComVisible(true)]
    [Guid("18B999A2-FC40-42D6-A4CD-815AD147BC5E"),ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgID + ".TNumeric")]
    public class TNumeric : INumeric
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
