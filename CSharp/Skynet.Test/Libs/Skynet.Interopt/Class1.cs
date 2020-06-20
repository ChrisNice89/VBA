using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using RGiesecke.DllExport;
using Skynet.Objects;

namespace ClassLibrary3
{
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public class Class1
    {
        public string Text
        {
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
            [param: MarshalAs(UnmanagedType.BStr)]
            set;
        }

        [return: MarshalAs(UnmanagedType.BStr)]
        public string TestMethod()
        {
            return Text + "...";
        }
      
        public Factory CreateFactory()
        {
            return new Factory();
        }
    }

    [ComVisible(true)]
    [Guid("A8656E7F-BDC2-4FB6-9914-EACD80DC8E66"),InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISample
    {
        // without MarshalAs(UnmanagedType.BStr), .Net will marshal these strings as single-byte Ansi!
        // BStr is equivalent to Delphi's WideString
        String Name
        {
            // this is how to add attributes to a getter's result parameter
            [return: MarshalAs(UnmanagedType.BStr)]
            get;
            // this is how to add attributes to a setter's value parameter
            [param: MarshalAs(UnmanagedType.BStr)]
            set;
        }

        int DoSomething(int value);
    }
    [ComVisible(true)]
    [Guid("F872621E-5D3F-45BB-A172-0BEE494F301A"), ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgID + ".Sample")]
    public class Sample : ISample
    {
        public String Name { get; set; }
        public int DoSomething(int value)
        {
            return value + 1;
        }
    }

    static class Exports
    {
        [DllExport]
        [return: MarshalAs(UnmanagedType.IDispatch)]
        public static Object CreateDotNetObject(String text)
        {
            //return new Class1 { Text = "CreateDotNetObject" + text };
            return new Factory();
        }

        [DllExport]
        [return: MarshalAs(UnmanagedType.IDispatch)]
        public static Object Factory()
        {
           return new Class1 { Text = "CreateDotNetObject" };
        }

        [DllExport]
        [return: MarshalAs(UnmanagedType.IDispatch)]
        public static Object CreateSample(String text)
        {
            return new Sample { Name = text };
        }
    }
}