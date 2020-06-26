using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Configuration;
using Skynet.Objects.Enums;
using System.Diagnostics;
using RGiesecke.DllExport;

namespace Skynet.Objects
{
    [Guid("00452587-C641-4F1F-92B6-483A4E146611"),
    ProgId(Constants.ProgID + ".Factory"),
    ClassInterface(ClassInterfaceType.None),
    ComDefaultInterface(typeof(IFactory)),
    ComVisible(true),
    ComSourceInterfaces(typeof(IFactoryEvents))]
    public class Factory : IFactory
    {
        public event Action<IObject> Created;
        public Constants Constant()
        {
            return new Constants();
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
    }

    [Guid("2FA865CE-7B95-40C5-8471-AC5C8306C51C"), ComVisible(true),
    InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IFactoryEvents
    {
        void Created(IObject instance);
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

