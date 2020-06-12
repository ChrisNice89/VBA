using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;


namespace iCAT.Interopt
{
    [ComVisible(true)]
    [Guid("0A577329-8441-427C-843F-5632C49F0763"), ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgID + ".Factory")]
    public class Factory : IFactory
    {
        public Constants Constant()
        {
            return new Constants();
        }

        public TDate TDate(object value)
        {
            return new TDate((string)value);
        }

        public string Test()
        {
            return this.GetType().ToString();
        }

        public TNumeric TNumeric(object value)
        {
            return new TNumeric((string)value);
        }

        public TString TString(object value)
        {
            return new TString((string)value);
        }
    }
    [ComVisible(true)]
    [Guid("471A3ABD-92DF-468C-B8B3-40963F31A236"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IFactory
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string Test();
        Constants Constant();
        TString TString(object value);
        TDate TDate(object value);
        TNumeric TNumeric(object value);


    }

    [ComVisible(false)]
    [Guid("20999B4B-E0CC-4BA9-AFBC-04E7F72217B2"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IFactoryTObjets
    { 
        TString TString(object value);
        TDate TDate(object value);
        TNumeric TNumeric(object value);
    }
}

