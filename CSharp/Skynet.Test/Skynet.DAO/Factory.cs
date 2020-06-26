using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Configuration;
using Skynet.DAO.Enums;
using System.Diagnostics;
using RGiesecke.DllExport;

namespace Skynet.DAO
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

        public string CreateConnection(ConnectionType T)
        {
            string Key = T.GetDescription();
            if (Key != null)
            {
                return ConfigurationManager.ConnectionStrings[Key].ConnectionString;
            }

            switch (T)
            {
                case ConnectionType.Excel:
                    {
                        return ConfigurationManager.ConnectionStrings["EXCEL"].ConnectionString;
                    }
                case ConnectionType.SQL:
                    {
                        return ConfigurationManager.ConnectionStrings["SQL"].ConnectionString;
                    }
                default: break;

            }
            return "";
        }

      
    }
    [ComVisible(true)]
    [Guid("471A3ABD-92DF-468C-B8B3-40963F31A236"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IFactory
    {
        string CreateConnection(ConnectionType T);
    }
    static class Exports
    {
        [DllExport]
        [return: MarshalAs(UnmanagedType.IDispatch)]
        public static Object Factory()
        {
            return new Factory();
        }
    }
}
