namespace AccessCodeLib.Data.Common.Sql
{
    public interface INamedSource : ISource
    {
        string Schema { get; }
        string Name { get; }
    }
}