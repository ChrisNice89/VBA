using System.Runtime.InteropServices;
using AccessCodeLib.Data.SqlTools.Converter.Common.Ansi92;

namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [Guid("3AC48E51-9B13-44B0-856A-F1B117B968F8")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".Ansi92SqlConverter")]
    public class Ansi92SqlConverter : SqlConverter, IAnsi92SqlConverterComInterface
    {
        public string GenerateSqlString(ISqlStatement sqlStatement)
        {
            return base.GenerateSqlString(sqlStatement);
        }
    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("7624599B-4CA1-408B-8F71-6DEC760691D3")]
    public interface IAnsi92SqlConverterComInterface : ISqlConverter
    {
        new string GenerateSqlString(ISqlStatement SqlStatement);
    }
}