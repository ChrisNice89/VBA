using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class FieldAlias : IFieldAlias
    {
        public FieldAlias(IField field, string alias)
        {
            Field = field;
            Alias = alias;
        }

        public FieldAlias(string name, ISource source, string alias)
        {
            Field = new Field(name, source);
            Alias = alias;
        }

        public string Alias { get; private set; }
        public IField Field { get; private set; }

        public string Name { get { return Field.Name; } }
        public ISource Source { get { return Field.Source; } }
        public FieldDataType DataType  { get { return Field.DataType; } }
    }
}
