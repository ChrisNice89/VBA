using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace iCAT.Interopt
{
    [ComVisible(true)]
    [Guid("E095D85C-348D-4922-AAB9-ADAFD48755ED"),InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IValue<in T>
    {
        IObject<T> Object();
        object Value { get; }
    }
}