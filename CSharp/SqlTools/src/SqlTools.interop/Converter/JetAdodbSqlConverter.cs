using System.Runtime.InteropServices;
using AccessCodeLib.Data.SqlTools.Converter.Jet.Oledb;

namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [Guid("2D5A405A-1D36-4C70-8417-0469B57087DE")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".JetAdodbSqlConverter")]
    public class JetAdodbSqlConverter : SqlConverter, IJetAdodbSqlConverterComInterface
    {
        public string GenerateSqlString(ISqlStatement sqlStatement)
        {
            return base.GenerateSqlString(sqlStatement);
        }
    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("9E17807B-B12A-4BEB-BDA0-79269BCDDA93")]
    public interface IJetAdodbSqlConverterComInterface : ISqlConverter
    {
        new string GenerateSqlString(ISqlStatement SqlStatement);
    }
}