using System.Collections.Generic;

namespace AccessCodeLib.Data.Common.Sql
{
    public interface ISqlStatement : IList<IStatement>
    {
        IEnumerable<IStatement> Find(string key);
    }
}