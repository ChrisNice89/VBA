using System.Runtime.InteropServices;
using AccessCodeLib.Data.SqlTools.Converter.Jet.Dao;

namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [Guid("C6DACF41-14C6-4829-B1AE-0075203D29A5")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".DaoSqlConverter")]
    public class DaoSqlConverter : SqlConverter, IDaoSqlConverterComInterface
    {
        public string GenerateSqlString(ISqlStatement sqlStatement)
        {
            return base.GenerateSqlString(sqlStatement);
        }
    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("24BB9B97-0636-4A17-BD40-048BE78C3546")]
    public interface IDaoSqlConverterComInterface : ISqlConverter
    {
        new string GenerateSqlString(ISqlStatement SqlStatement);
    }
}