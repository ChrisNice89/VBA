using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Configuration;
using Skynet.Objects.Enums;

namespace Skynet.Objects
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

        public string CreateConnection(Connectiontype T)
        {
            string Key = T.GetDescription();
            if (Key != null){
                return ConfigurationManager.ConnectionStrings[Key].ConnectionString;
            }
          
            switch (T)
            {
                case Connectiontype.EXCEL:
                    {
                      return ConfigurationManager.ConnectionStrings["EXCEL"].ConnectionString;
                    }
                case Connectiontype.SQL:
                    {
                        return ConfigurationManager.ConnectionStrings["SQL"].ConnectionString;
                    }
                default: break;

            }
            return "";
        }

        public TDate TDate(DateTime Value)
        {
            return new TDate(Value);
        }

        public string Test()
        {
            return this.GetType().ToString();
        }

        public TNumeric TNumeric(int Value)
        {
            return new TNumeric( Value);
        }

        public TString TString(string Value)
        {
            return new TString(Value);
        }
    }
    [ComVisible(true)]
    [Guid("471A3ABD-92DF-468C-B8B3-40963F31A236"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IFactory
    {
        string Test();
        Constants Constant();
        TString TString(string Value);
        TDate TDate(DateTime value);
        TNumeric TNumeric(int value);

        string CreateConnection(Connectiontype T);
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

