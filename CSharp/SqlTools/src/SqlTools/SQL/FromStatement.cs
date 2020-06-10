using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class FromStatement : IFromStatement
    {
        public FromStatement()
        {
        }

        public FromStatement(ISource source)
        {
            Source = source;
        }

        public FromStatement(string source)
        {
            Source = new NamedSource(source);
        }

        public string Key
        {
            get { return StatementKeys.From; }
        }

        public ISource Source { get; set; }
    }
}
