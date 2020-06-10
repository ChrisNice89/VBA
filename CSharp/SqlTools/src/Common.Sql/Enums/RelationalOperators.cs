using System;

namespace AccessCodeLib.Data.Common.Sql
{
    [Flags]
    // Diese Einträge können kombiniert werden. Beispiel: ">=" ... op = RelationalOperators.Equal | RelationalOperators.GreaterThan)
    public enum RelationalOperators
    {
        Not = 1,
        Equal = 2,          // Is Null ... RelationalOperators = Equal, Value = Null ?
        LessThan = 4,
        GreaterThan = 8,
        Like = 256,
        Between = 512,
        In = 1024,
        AddWildcardSuffix = 2048,
        AddWildcardPrefix = 4096
    }
}