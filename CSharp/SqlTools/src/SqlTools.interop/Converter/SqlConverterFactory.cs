using System.Runtime.InteropServices;

namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [Guid("27444DA3-BE19-4425-9771-BACD47B20321")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".SqlConverterFactory")]
    public class SqlConverterFactory : ISqlConverterFactory
    {
        public ISqlConverter Ansi92SqlConverter() { return new Ansi92SqlConverter(); }
        public ISqlConverter DaoSqlConverter() { return new DaoSqlConverter(); }
        public ISqlConverter TsqlSqlConverter() { return new TsqlSqlConverter(); }
        public ISqlConverter JetAdodbSqlConverter() { return new JetAdodbSqlConverter(); } 
    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("58064D7A-3CC0-458C-AC9E-C0E1EA790602")]
    public interface ISqlConverterFactory
    {
        ISqlConverter Ansi92SqlConverter();
        ISqlConverter DaoSqlConverter();
        ISqlConverter TsqlSqlConverter();
        ISqlConverter JetAdodbSqlConverter();
    }
}
