using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace iCAT.Objects
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)] // intelsense
    [Guid("23E64AEC-14DE-4609-8F89-F7A890921C0C")]
    public interface IFactory
    {
        [DispId(1)]
        TString TString(string value);
        [DispId(2)]
        TNumeric TNumeric(string value);
    }
}
