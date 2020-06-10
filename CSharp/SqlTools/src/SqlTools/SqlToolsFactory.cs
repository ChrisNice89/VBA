namespace AccessCodeLib.Data.SqlTools
{
    public static class SqlToolsFactory
    {
        public static ISqlGenerator SqlGenerator
        {
            get { return new SqlGenerator();}
        }

        public static FieldGenerator FieldGenerator
        {
            get { return new FieldGenerator(); }
        }
    }
}
