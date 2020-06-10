using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class NamedSource : INamedSource
    {
        public NamedSource(string tablename, string schema = null)
        {
            Name = tablename;
            Schema = schema;
        }

        public string Schema { get; private set; }
        public string Name { get; private set; }
    }
}
