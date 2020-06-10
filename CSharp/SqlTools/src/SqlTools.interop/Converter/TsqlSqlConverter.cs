using System.Runtime.InteropServices;
using AccessCodeLib.Data.SqlTools.Converter.Mssql;

namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [Guid("F024506B-EFE4-479E-A93C-C9BEE909C422")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".TsqlSqlConverter")]
    public class TsqlSqlConverter : SqlConverter, ITsqlSqlConverterComInterface
    {
        public string GenerateSqlString(ISqlStatement sqlStatement)
        {
            return base.GenerateSqlString(sqlStatement);
        }
    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("9284A7FE-DD59-4806-91CE-80AB5E129E0A")]
    public interface ITsqlSqlConverterComInterface : ISqlConverter
    {
        new string GenerateSqlString(ISqlStatement SqlStatement);
    }
}