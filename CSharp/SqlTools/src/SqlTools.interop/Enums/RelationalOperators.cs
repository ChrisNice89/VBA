using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.Data.SqlTools.interop
{
    [Flags]
    [ComVisible(true)]
    [Guid("3C823FAF-8EBF-43C4-B0A6-184DAD808B97")]
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