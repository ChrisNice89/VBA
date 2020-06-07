using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;


namespace SqlTools.Test
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)] // intelsense
    [Guid("F23B9801-1BF1-454C-88B4-3D2DB36D4B5C")]
    public interface ImyTestClass
    {
        string Test();
        string prop { get; set; }
        void sub();
    }

    [ComVisible(true)]
    [Guid("18B999A2-FC40-42D6-A4CD-815AD147BC5E")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("SqlToolsTest.myTestClass")]
    public class mytestclass: ImyTestClass
    {
        private string myProp;
        public string prop { get { return "property"; } set { myProp = value; } }

        public void sub()
        {
            throw new NotImplementedException();
        }
        public string Test()
        {
            return "done";
        }
    }
}
