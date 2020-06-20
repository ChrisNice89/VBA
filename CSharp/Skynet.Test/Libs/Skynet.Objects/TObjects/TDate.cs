using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Skynet.Objects
{
    [ComVisible(true)]
    [Guid("724EDBBB-C612-4B49-AA38-A4120549AAE2"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDate
    {
        string Test();
        DateTime Value { get; }
    }

    [ComVisible(true)]
    [Guid("D5A3D24A-C15F-44CA-8183-13E0708A98FE"), ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgID + ".TNumeric")]
    public class TDate : IDate
    {
        public DateTime Value { get; }
        public IObject Object =>  (IObject)this;

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

        public TDate() { }
        public TDate(DateTime Value)
        {
            _ = Value;
        }
    }
}
