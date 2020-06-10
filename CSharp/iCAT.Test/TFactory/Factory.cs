using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace iCAT.Objects
{
    [ComVisible(true)]
    [Guid("F0328978-58FF-4907-9864-692B8C94157D"),ClassInterface(ClassInterfaceType.None)]
    [ProgId("iCAT.Factory")]
    class Factory : IFactory
    {
       
        public TNumeric TNumeric(string value)
        {
            return new Objects.TNumeric(value);
        }

        public TString TString(string value)
        {
            return new Objects.TString(value);
        }
    }
}
