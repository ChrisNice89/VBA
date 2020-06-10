namespace AccessCodeLib.Data.SqlTools
{
    public interface ISqlGenerator
    {
        ISqlGenerator Select(params string[] fieldNames);
        ISqlGenerator SelectAlias(string fieldName, string alias);
        ISqlGenerator SelectAll();
        ISqlGenerator From(string source);
        ISqlGenerator Where(string whereString);
        ISqlGenerator GroupBy(params string[] fieldNames);
        ISqlGenerator Having(string havingString);
        ISqlGenerator OrderBy(params string[] fieldNames);
        string ToDaoString();
    }
}