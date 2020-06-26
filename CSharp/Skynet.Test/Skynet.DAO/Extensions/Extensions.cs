using Skynet.DAO.Attributes;
using Skynet.DAO.Reflection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Skynet.DAO.Extensions
{
    static class QueryResult
    {
        public static IList<T> ToList<T>(this SqlDataReader reader) 
            where T : class, new()
        {
            //Action<T, PropertyInfo, object> Inject = (instance, property, value) => property.SetValue(instance, Convert.ChangeType(value, property.PropertyType), null);
            var columns = Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).ToArray();
            PropertyInfo[] modelProperties = New<T>.GetProperties<SqlField>();
            PropertyInfo[] properties = new PropertyInfo[reader.FieldCount];

            //var props = typeof(T).GetProperties().Select(p => new { p, attr = p.GetCustomAttribute<SqlField>() }).Where(p => p.attr != null);
            //var props = typeof(T).GetProperties().Select(p => p.GetCustomAttribute<SqlField>()).Where(p => p.Name!=null);

            for (var i = 0; i < columns.Length; ++i)
            {
                var property = modelProperties.SingleOrDefault(p=> p.GetCustomAttribute<SqlField>().Name.Equals(columns[i], StringComparison.InvariantCultureIgnoreCase));
                if (property != null)
                {
                    properties[i] = property;
                }
            }

            IList<T> list = new List<T>(256);
            while (reader.Read())
            {
                var instance = New<T>.Instance();
                var values = new object[reader.FieldCount];
                reader.GetValues(values);
                
                for (var i = 0; i < values.Length; ++i)
                {
                    if (values[i] == DBNull.Value)
                    {
                        values[i] = null;
                    }
                    instance.Inject(properties[i], values[i]);
                }
                list.Add(instance);
            }
            return list;
        }
        public static IList<T> SqlQuery<T>(this SqlConnection connection, string Sql)
            where T : class, new()
        {
            return new Query<T>(connection.ConnectionString).Execute(Sql);
        }
        public static Query<T> CreateQuery<T>(this SqlConnection connection)
            where T : class, new()
        {
            return new Query<T>(connection.ConnectionString);
        }
    }

    public class Query<T>
        where T : class, new()
    {
        internal readonly string Sql;
        protected readonly string _connectionString;
        public Query(string connectionString)
        {
            _connectionString = connectionString;
        }
        public Query(string Sql, string connectionString)
        {
            this.Sql = Sql;
            _connectionString = connectionString;
        }

        internal void Prepare()
        {
            New<T>.Compile();
        }
        internal IList<T> Execute()
        {
            return Execute(Sql);
        }
        internal IList<T> Execute(string query)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                using (var command = new SqlCommand(query, connection))
                {
                    return new ReadOnlyCollection<T>(command.ExecuteReader().ToList<T>());
                }
            }
        }
    }
}
static class DbTypeConverter
    {
        private static Dictionary<SqlDbType, Type> DataTypes;
        private static Dictionary<Type, SqlDbType> SqlTypes;
        static DbTypeConverter()
        {
            DataTypes = new Dictionary<SqlDbType, Type>();
            DataTypes.Add(SqlDbType.BigInt, typeof(Int64));
            DataTypes.Add(SqlDbType.Binary, typeof(Byte[]));
            DataTypes.Add(SqlDbType.Bit, typeof(Boolean));
            DataTypes.Add(SqlDbType.Char, typeof(String));
            DataTypes.Add(SqlDbType.Date, typeof(DateTime));
            DataTypes.Add(SqlDbType.DateTime, typeof(DateTime));
            DataTypes.Add(SqlDbType.DateTime2, typeof(DateTime));
            DataTypes.Add(SqlDbType.DateTimeOffset, typeof(DateTimeOffset));
            DataTypes.Add(SqlDbType.Decimal, typeof(Decimal));
            DataTypes.Add(SqlDbType.Float, typeof(Double));
            DataTypes.Add(SqlDbType.Image, typeof(Byte[]));
            DataTypes.Add(SqlDbType.Int, typeof(Int32));
            DataTypes.Add(SqlDbType.Money, typeof(Decimal));
            DataTypes.Add(SqlDbType.NChar, typeof(String));
            DataTypes.Add(SqlDbType.NText, typeof(String));
            DataTypes.Add(SqlDbType.NVarChar, typeof(String));
            DataTypes.Add(SqlDbType.Real, typeof(Single));
            DataTypes.Add(SqlDbType.SmallDateTime, typeof(DateTime));
            DataTypes.Add(SqlDbType.SmallInt, typeof(Int16));
            DataTypes.Add(SqlDbType.SmallMoney, typeof(Decimal));
            DataTypes.Add(SqlDbType.Text, typeof(String));
            DataTypes.Add(SqlDbType.Time, typeof(TimeSpan));
            DataTypes.Add(SqlDbType.Timestamp, typeof(Byte[]));
            DataTypes.Add(SqlDbType.TinyInt, typeof(Byte));
            DataTypes.Add(SqlDbType.UniqueIdentifier, typeof(Guid));
            DataTypes.Add(SqlDbType.VarBinary, typeof(Byte[]));
            DataTypes.Add(SqlDbType.VarChar, typeof(String));

            SqlTypes = new Dictionary<Type, SqlDbType>();
            SqlTypes.Add(typeof(Boolean), SqlDbType.Bit);
            SqlTypes.Add(typeof(String), SqlDbType.NVarChar);
            SqlTypes.Add(typeof(DateTime), SqlDbType.DateTime);
            SqlTypes.Add(typeof(Int16), SqlDbType.Int);
            SqlTypes.Add(typeof(Int32), SqlDbType.Int);
            SqlTypes.Add(typeof(Int64), SqlDbType.Int);
            SqlTypes.Add(typeof(Decimal), SqlDbType.Float);
            SqlTypes.Add(typeof(Double), SqlDbType.Float);
        }
        public static Type ToDataType(this SqlDbType sqlType)
        {
            Type datatype = null;
            if (DataTypes.TryGetValue(sqlType, out datatype))
                return datatype;
            throw new TypeLoadException(string.Format("Can not load CLR Type from {0}", sqlType));
        }

        public static SqlDbType ToSqlType(Type SysType)
        {
            SqlDbType datatype = SqlDbType.NVarChar;
            if (SqlTypes.TryGetValue(SysType, out datatype))
                return datatype;
            throw new TypeLoadException(string.Format("Can not load Sql Type from {0}", SysType));
        }

        public static Type GetClrType(SqlDbType sqlType, bool isNullable)
        {
            switch (sqlType)
            {
                case SqlDbType.BigInt:
                    return isNullable ? typeof(long?) : typeof(long);

                case SqlDbType.Binary:
                case SqlDbType.Image:
                case SqlDbType.Timestamp:
                case SqlDbType.VarBinary:
                    return typeof(byte[]);

                case SqlDbType.Bit:
                    return isNullable ? typeof(bool?) : typeof(bool);

                case SqlDbType.Char:
                case SqlDbType.NChar:
                case SqlDbType.NText:
                case SqlDbType.NVarChar:
                case SqlDbType.Text:
                case SqlDbType.VarChar:
                case SqlDbType.Xml:
                    return typeof(string);

                case SqlDbType.DateTime:
                case SqlDbType.SmallDateTime:
                case SqlDbType.Date:
                case SqlDbType.Time:
                case SqlDbType.DateTime2:
                    return isNullable ? typeof(DateTime?) : typeof(DateTime);

                case SqlDbType.Decimal:
                case SqlDbType.Money:
                case SqlDbType.SmallMoney:
                    return isNullable ? typeof(decimal?) : typeof(decimal);

                case SqlDbType.Float:
                    return isNullable ? typeof(double?) : typeof(double);

                case SqlDbType.Int:
                    return isNullable ? typeof(int?) : typeof(int);

                case SqlDbType.Real:
                    return isNullable ? typeof(float?) : typeof(float);

                case SqlDbType.UniqueIdentifier:
                    return isNullable ? typeof(Guid?) : typeof(Guid);

                case SqlDbType.SmallInt:
                    return isNullable ? typeof(short?) : typeof(short);

                case SqlDbType.TinyInt:
                    return isNullable ? typeof(byte?) : typeof(byte);

                case SqlDbType.Variant:
                case SqlDbType.Udt:
                    return typeof(object);

                case SqlDbType.Structured:
                    return typeof(DataTable);

                case SqlDbType.DateTimeOffset:
                    return isNullable ? typeof(DateTimeOffset?) : typeof(DateTimeOffset);

                default:
                    throw new ArgumentOutOfRangeException("sqlType");
            }
        }
}
