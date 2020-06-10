using System.Collections.Generic;

namespace AccessCodeLib.Data.Common.Sql.Converter
{
    public interface ISqlConverter
    {
        string GenerateSqlString(ISqlStatement sqlStatement);
        string GenerateConditionString(IConditionGroup conditionGroup);
    }
}