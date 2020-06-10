using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Converter.Mssql
{
    internal class NameConverter : NameConverterBase
    {
        protected override string GetCheckedSourceNameString(INamedSource source)
        {
            return SqlConverterTools.CheckedSourceNameString(source);
        }

        protected override string GetCheckedItemNameString(string name)
        {
            return SqlConverterTools.CheckedItemNameString(name);
        }
    }
}