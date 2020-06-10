using System;
using System.Collections.Generic;
using System.Linq;
using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class SqlStatement : List<IStatement>, ISqlStatement
    {
        public IEnumerable<IStatement> Find(string key)
        {
            return this.Where(s => s.Key.Equals(key, StringComparison.InvariantCultureIgnoreCase)).ToList();
        }
    }
}
