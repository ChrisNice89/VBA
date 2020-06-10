using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Converter.Common.Ansi92
{
    internal class NameConverter : NameConverterBase
    {
        protected override string GetCheckedSourceNameString(INamedSource name)
        {
            return SqlConverterTools.CheckedItemNamedSourceString(name);
        }

        protected override string GetCheckedItemNameString(string name)
        {
            return SqlConverterTools.CheckedItemNameString(name);
        }
    }
}