using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class SubSelectSource : ISubSelect
    {
        public SubSelectSource(ISqlStatement sqlStatement)
        {
            SqlStatement = sqlStatement;
        }

        public ISqlStatement SqlStatement { get; private set; }
    }
}
