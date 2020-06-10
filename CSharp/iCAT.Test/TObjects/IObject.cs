using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace iCAT.Objects
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)] // intelsense
    [Guid("F23B9801-1BF1-454C-88B4-3D2DB36D4B5C")]
    public interface IObject
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string Test();
        object Value { get; set; }
        void sub();
    }
}
