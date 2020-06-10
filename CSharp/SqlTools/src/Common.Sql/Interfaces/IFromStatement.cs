namespace AccessCodeLib.Data.Common.Sql
{
    public interface IFromStatement : IStatement
    {
        ISource Source { get; }
    }
}