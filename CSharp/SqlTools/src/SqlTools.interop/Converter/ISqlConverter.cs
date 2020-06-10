using System.Runtime.InteropServices;

namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("E76E1F36-72E9-4CCA-95C9-5711E17E69F4")]
    public interface ISqlConverter : Common.Sql.Converter.ISqlConverter
    {
        string GenerateSqlString(ISqlStatement SqlStatement);
    }
}