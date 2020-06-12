using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace iCAT.Interopt
{
    [ComVisible(true)]
    [Guid("724EDBBB-C612-4B49-AA38-A4120549AAE2"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDate
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string Test();
        IObject Object { get; }

    }

    [ComVisible(true)]
    [Guid("D5A3D24A-C15F-44CA-8183-13E0708A98FE"), ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgID + ".TNumeric")]
    public class TDate : IDate,IObject<TDate>,IValue<TDate>
    {
   
        public object  Value { get; }

        object IValue<TDate>.Value => throw new NotImplementedException();

        public IObject Object =>  (IObject)this;

        public string Test()
        {
            return this.GetType().ToString();
        }

        public override string ToString()
        {
            return ToString(null, System.Globalization.CultureInfo.CurrentCulture);
        }

        [return: MarshalAs(UnmanagedType.Bool)]
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

        bool IObject<TDate>.Equals(TDate other)
        {
            throw new NotImplementedException();
        }

        CompareResult IObject<TDate>.CompareTo(TDate other)
        {
            throw new NotImplementedException();
        }

        string IFormattable.ToString(string format, IFormatProvider formatProvider)
        {
            throw new NotImplementedException();
        }

        IObject<TDate> IValue<TDate>.Object()
        {
            throw new NotImplementedException();
        }

        public TDate() { }
        public TDate(string value)
        {
            this.Value = value;
        }


    }
}
