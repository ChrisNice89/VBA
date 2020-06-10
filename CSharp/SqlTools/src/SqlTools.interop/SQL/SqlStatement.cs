using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [Guid("D9D30156-0CEC-48F8-9106-8FF0DC9887B5")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".SqlStatement")]
    public class SqlStatement : Sql.SqlStatement, ISqlStatement
    {
    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("3D5EF742-C31F-43C8-AFED-C686A5B12BDC")]
    public interface ISqlStatement : Common.Sql.ISqlStatement
    {
    }
}