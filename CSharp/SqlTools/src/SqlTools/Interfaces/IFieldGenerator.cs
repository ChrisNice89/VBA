using System.Collections.Generic;
using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools
{
    public interface IFieldGenerator
    {
// ReSharper disable ReturnTypeCanBeEnumerable.Global
        IField[] FromString(string fieldsString, char deliminator = ',');
        IField[] FromArray(IEnumerable<string> fieldNames);
// ReSharper restore ReturnTypeCanBeEnumerable.Global
    }
}
