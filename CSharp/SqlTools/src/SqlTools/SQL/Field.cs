using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class Field : IField
    {
        protected Field()
        {
            DataType = FieldDataType._Unspecified;
        }

        public Field(string name, ISource source = null, FieldDataType dataType = FieldDataType._Unspecified)
        {
            Name = name;
            Source = source;
            DataType = dataType;
        }

        public string Name { get; private set; }
        public ISource Source { get; private set; }
        public FieldDataType DataType { get; set; }
    }
}
