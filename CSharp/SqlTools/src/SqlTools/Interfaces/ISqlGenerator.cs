using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools
{
    public interface ISqlGenerator
    {
// ReSharper disable UnusedMethodReturnValue.Global
        ISqlGenerator Select(params IField[] fields);
        ISqlGenerator Select(params string[] fields);
        ISqlGenerator SelectAll();
        ISqlGenerator SelectField(string fieldName, ISource source = null, string alias = "");

        ISqlGenerator From(ISource source);
        ISqlGenerator From(string sourceName);

        ISqlGenerator Join(ISource source, ICondition condition, JoinOperator op = JoinOperator.Inner);
        
        ISqlGenerator Where(ICondition condition);
        ISqlGenerator Where(IField field, RelationalOperators op, object value);
        ISqlGenerator Where(string whereString);

        ISqlGenerator GroupBy(params IField[] fieldNames);
        ISqlGenerator GroupBy(params string[] fieldNames);

        ISqlGenerator Having(ICondition condition);
        ISqlGenerator Having(IField field, RelationalOperators op, object value);
        ISqlGenerator Having(string havingString);

        ISqlGenerator OrderBy(params IField[] fieldNames);
        ISqlGenerator OrderBy(params string[] fieldNames);

// ReSharper restore UnusedMethodReturnValue.Global
        ISqlStatement SqlStatement { get; }
        string ToString();
    }
}