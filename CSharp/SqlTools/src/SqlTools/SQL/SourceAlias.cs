using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class SourceAlias : ISourceAlias
    {
        public SourceAlias(ISource source, string alias)
        {
            Source = source;
            Alias = alias;
        }

        public string Alias { get; set; }
        public ISource Source { get; private set; }
    }
}