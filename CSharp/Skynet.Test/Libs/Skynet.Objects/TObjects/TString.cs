using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;


namespace Skynet.Objects
{
    [ComVisible(true)]
    [Guid("70165784-CDD8-4973-AADB-2EDB018DF3DE"),ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgID +".TString")]
    public class TString: IString, IValue
    {
        public IObject Object =>  (IObject) this;
        object IValue.Value => (object)Value;

        public string Value { get; }

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

        public TString(){ }
        public TString(string Value)
        {
            _ =Value;
        }
    }
    [ComVisible(true)]
    [Guid("1493456F-A1F8-4627-A8B3-1CB81BB198BC"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IString
    {
        string Test();
        string Value { get;  }
    }
}

