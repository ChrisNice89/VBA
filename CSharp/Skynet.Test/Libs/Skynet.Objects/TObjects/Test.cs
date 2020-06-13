using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Skynet.Objects
{
    [Guid("6F428F16-B90A-4E4D-AE6F-C5F53CB1A59B"),ClassInterface(ClassInterfaceType.None),
       ComSourceInterfaces(typeof(IComEvents))]
    [ProgId(Constants.ProgID + ".Test")]
    class Test : ICom
    {
        public object Value { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public void ComCall()
        {
            throw new NotImplementedException();
        }

    }

    [ComVisible(true)]
    [Guid("8BE0CA44-A1A8-4AAA-BF10-CC497C1CD3D2"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICom
    {
     void ComCall();
     }

     [Guid("BA522FED-4BF3-4BDC-8C6B-CA83D7246175"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IComEvents
     {
     }
}

