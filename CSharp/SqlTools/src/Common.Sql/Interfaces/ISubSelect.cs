namespace AccessCodeLib.Data.Common.Sql
{
    public interface ISubSelect : ISource
    {
        ISqlStatement SqlStatement { get; }
    }
}