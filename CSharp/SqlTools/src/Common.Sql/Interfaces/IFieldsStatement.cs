namespace AccessCodeLib.Data.Common.Sql
{
    public interface IFieldsStatement : IStatement
    {
        IFieldList Fields { get; }
    }
}
