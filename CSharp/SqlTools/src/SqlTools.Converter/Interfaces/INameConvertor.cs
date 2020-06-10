using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Converter
{
    public interface INameConverter
    {
        string GenerateAliasNameString(IAlias alias);
        string GenerateSourceNameString(ISource source);
        string GenerateFieldString(IField field);
        string GenerateSelectFieldString(IField field);
        string GenerateFieldNameString(ISource source, string fieldName);
    }
}