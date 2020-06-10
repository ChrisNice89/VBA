namespace AccessCodeLib.Data.Common.Sql
{
    public interface ICondition
    {
    }

    public interface IFieldCondition : ICondition
    {
        IField Field { get; }
        RelationalOperators Operator { get; }
        object Value { get; }
    }

    public interface IConditionString : ICondition
    {
        string Value { get; }
    }

}