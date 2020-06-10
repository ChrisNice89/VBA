namespace AccessCodeLib.Data.Common.Sql
{
    public interface IConditionStringStatement : IStatement
    {
        string Condition { get; }
    }
}