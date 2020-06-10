using System.Runtime.InteropServices;
using AccessCodeLib.Data.SqlTools.Sql;
using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [Guid("217E3F77-7EC2-4D61-A29E-FA9681714251")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".Field")]
    public class Field : Sql.Field, IField
    {
        public Field()
        {
        }

        public Field(string name, object source = null, FieldDataType dataType = FieldDataType._Unspecified) 
            : base(name, source is string ? new NamedSource((string)source) : (ISource)source, (Common.Sql.FieldDataType) dataType)
        {
        }

        public Field(Common.Sql.IField field) : base(field.Name, field.Source, field.DataType)
        {
        }

        public new object Source
        {
            get { return base.Source; }
        }

        public new FieldDataType DataType
        {
            get { return (FieldDataType)base.DataType; }
            set { base.DataType = (Common.Sql.FieldDataType) value; }
        }
    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("E47C0214-563B-472D-8335-039F55A2839F")]
    public interface IField : Common.Sql.IField
    {
        new string Name { get; }
        new object Source { get; }
        new FieldDataType DataType { get; set; }
    }
}
