using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Skynet.Objects.Enums;

namespace Skynet.Objects
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)] // intelsense
    [Guid("268AE869-9119-47FE-976D-AD48AA382131")]
    public interface IObject: IFormattable
    {
        Boolean Equals(IObject other);
        CompareResult CompareTo(IObject other);
    }

    [ComVisible(false)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)] // intelsense
    [Guid("F23B9801-1BF1-454C-88B4-3D2DB36D4B5C")]
    public interface IObject<in T>: IFormattable
    {
        [return: MarshalAs(UnmanagedType.Bool)]
        Boolean Equals(T other);
        CompareResult CompareTo(T other);
    }

}
