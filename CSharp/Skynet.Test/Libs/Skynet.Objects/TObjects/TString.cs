
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;

namespace Skynet.Objects
{
    [Guid("70165784-CDD8-4973-AADB-2EDB018DF3DE"),
    ProgId(Constants.ProgID + ".TString"),
    ClassInterface(ClassInterfaceType.None),
    ComDefaultInterface(typeof(IString)),
    ComVisible(true),
    ComSourceInterfaces(typeof(IStringEvents))]
    public class TString : TBase<TString, String>, IString
    {
        #region ComEvents
        public event Action<string> Created;
        #endregion ComEvents
        #region constructors
        public TString() : base("") {  }
        public TString(string Value) : base(Value)
        {
            if (Created != null)
                Created(Value);
        }
        #endregion constructors
        #region implicit operator 
        public static implicit operator TString(string value) { return new TString(value); }
        public static implicit operator string(TString TCustom) { return TCustom._value; }
        #endregion implicit operator 
        #region Com
        string IString.Value => this._value;
        IObject IString.Object => this;
        #region IObject
        bool IObject.Equals(IObject other) => base.Equals(other);
        CompareResult IObject.CompareTo(IObject other) => base.CompareTo(other);
        bool IObject.IsRelatedTo(IObject other) => base.IsRelatedTo(other);
        string IObject.ToString() => base.ToString();
        #endregion IObject
        #endregion com
    }

    [Guid("1493456F-A1F8-4627-A8B3-1CB81BB198BC"), ComVisible(true),
    InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IString:IObject
    {
        IObject Object { get; }
        string Value { get; }
    }
    
    // Events interface Database_COMObjectEvents 
    [Guid("67bd8422-9641-4675-acda-3dfc3c911a07"),ComVisible(true),
    InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IStringEvents
    {
        void Created(string state);
    }

    //[ComImport]
    //[Guid("79eac9d0-baf9-11ce-8c82-00aa004ba90b")]
    //[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    //public interface IConstructor
    //{
    //    [PreserveSig]
    //    int Authenticate(
    //        [In, Out] ref IntPtr phwnd,
    //        [In, Out, MarshalAs(UnmanagedType.LPWStr)] ref string pszUsername,
    //        [In, Out, MarshalAs(UnmanagedType.LPWStr)] ref string pszPassword);
    //}
}

