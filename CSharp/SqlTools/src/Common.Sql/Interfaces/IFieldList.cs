using System.Collections.Generic;

namespace AccessCodeLib.Data.Common.Sql
{
    public interface IFieldList : IList<IField>
    {
        void Add(params string[] fieldNames);
        void Add(params IField[] fields);
    }
}