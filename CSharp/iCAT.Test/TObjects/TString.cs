using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;


namespace iCAT.Objects
{

    [ComVisible(true)]
    [Guid("70165784-CDD8-4973-AADB-2EDB018DF3DE")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("iCAT.TString")]
    public class TString: IObject
    {
        private object value;
        public object Value 
        {
            get => value;
            set => _ = value;
        }

        object IObject.Value { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        void IObject.sub()
        {
            throw new NotImplementedException();
        }

        public string Test()
        {
            return this.GetType().ToString();
        }
        public TString() { }
        public TString(string value)
        {
            this.Value = value;
        }
    }
}

