using System.Collections.Generic;
using System.Linq;
using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class FieldList : List<IField>, IFieldList
    {
        public void Add(params string[] fieldNames)
        {
            foreach (var fieldName in fieldNames.Where(fieldName => !string.IsNullOrEmpty(fieldName) && !string.IsNullOrEmpty(fieldName.Trim())))
            {
                Add(new Field(fieldName));
            }
        }

        public void Add(params IField[] fields)
        {
            AddRange(fields.Where(field => field != null && !string.IsNullOrEmpty(field.Name) && !string.IsNullOrEmpty(field.Name.Trim())));
        }
    }
}
