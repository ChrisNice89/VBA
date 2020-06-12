using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;


namespace iCAT.Interopt
{
    [ComVisible(true)]
    [Guid("70165784-CDD8-4973-AADB-2EDB018DF3DE"),ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgID +".TString")]
    public class TString: IString
    {
        private object value;
        public object Value 
        {
            get => value;
            set => _ = value;
        }

        public void sub()
        {
            throw new NotImplementedException();
        }

        public string Test()
        {
            return this.GetType().ToString();
        }
        public TString(){ }
        public TString(string value)
        {
            this.Value = value;
        }
    }
    [ComVisible(true)]
    [Guid("1493456F-A1F8-4627-A8B3-1CB81BB198BC"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IString
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string Test();
        object Value { get; set; }
        void sub();
    }
}

