using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Converter.Jet
{
    public class NameConverter : NameConverterBase
    {
        protected override string GetCheckedItemNameString(string name)
        {
            return SqlConverterTools.CheckedItemNameString(name);
        }

        protected override string GetCheckedSourceNameString(INamedSource source)
        {
            return SqlConverterTools.CheckedSourceNameString(source);
        }
    }
}