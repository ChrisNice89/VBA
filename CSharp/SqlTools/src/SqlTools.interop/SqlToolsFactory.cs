using System.Runtime.InteropServices;

namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [Guid("A3559212-7730-42CF-B00F-C670C0647F8B")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".SqlToolsFactory")]
    public class SqlToolsFactory : ISqlToolsFactory
    {
        public SqlGenerator SqlGenerator(ISqlConverter converter = null)
        {
            return new SqlGenerator(converter);
        }

        private static readonly ISqlConverterFactory _sqlConverters = new SqlConverterFactory();
        public ISqlConverterFactory SqlConverters
        {
            get { return _sqlConverters; }
        }

        public FieldGenerator FieldGenerator()
        {
            return new FieldGenerator();
        }

        public ConditionGenerator ConditionGenerator()
        {
            return new ConditionGenerator();
        }

        public ConditionStringBuilder ConditionStringBuilder(ISqlConverter converter = null)
        {
            return new ConditionStringBuilder(converter);
        }

    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("39B72296-C4A1-47EA-AD65-A643356D12D8")]
    public interface ISqlToolsFactory
    {
        SqlGenerator SqlGenerator(ISqlConverter Converter = null);
        ISqlConverterFactory SqlConverters { get; }
        FieldGenerator FieldGenerator();
        ConditionGenerator ConditionGenerator();
        ConditionStringBuilder ConditionStringBuilder(ISqlConverter Converter = null);
    }
}
