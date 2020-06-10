namespace AccessCodeLib.Data.Common.Sql
{
    public interface ISourceAlias : ISource, IAlias
    {
        ISource Source { get; }
    }
}