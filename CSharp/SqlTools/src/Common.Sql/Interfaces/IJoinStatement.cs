namespace AccessCodeLib.Data.Common.Sql
{
    public interface IJoinStatement : IFromStatement
    {
        JoinOperator JoinOperator { get; }
        ICondition Condition { get; }
    }
}