using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public abstract class FieldsStatement : IFieldsStatement
    {
        protected FieldsStatement()
        {
            Fields = new FieldList();
        }

        protected FieldsStatement(params IField[] fields)
        {
            Fields = new FieldList {fields};
        }

        protected FieldsStatement(params string[] fieldNames)
        {
            Fields = new FieldList {fieldNames};
        }

        public abstract string Key { get; }
        public IFieldList Fields { get; set; }
    }
}
