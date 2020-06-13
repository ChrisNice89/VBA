using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Skynet.Objects
{

    [ComVisible(true)]
    [Guid("501D2B36-B470-4B36-8A84-B099914DB3BD"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IValue:IObject
    {
        IObject Object { get; }
        object Value { get; }
    }

    [ComVisible(true)]
    [Guid("E095D85C-348D-4922-AAB9-ADAFD48755ED"),InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IValue<in T>
    {
        IObject<T> Object();
        object Value { get; }
    }
}