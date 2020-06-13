using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Skynet.Objects
{
    [ComVisible(true)]
    [Guid("18727C24-0875-4DEE-80BB-38BD6D4457A7"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface INumeric
    {
        string Test();
        int Value { get;  }
    }

    [ComVisible(true)]
    [Guid("18B999A2-FC40-42D6-A4CD-815AD147BC5E"),ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgID + ".TNumeric")]
    public class TNumeric : INumeric,IValue
    {
        public int Value { get; }
        public IObject Object => (IObject)this;
        object IValue.Value => (object)Value;
        public string Test()
        {
            return this.GetType().ToString();
        }
        public bool Equals(IObject other)
        {
            throw new NotImplementedException();
        }

        public CompareResult CompareTo(IObject other)
        {
            throw new NotImplementedException();
        }

        public string ToString(string format, IFormatProvider formatProvider)
        {
            throw new NotImplementedException();
        }

        public TNumeric() { }
        public TNumeric(int Value)
        {
            this.Value = Value;
        }
    }
}
