using Skynet.DAO.Attributes;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace Skynet.DAO.Reflection
{
    public static class New<T>
        where T : class, new()
    {
        public static readonly Func<T> Instance = Constructor();
        //public static Func<T> Instance { get; }// = Constructor();
        //public readonly Func<T> Instance = Constructor();
        static Func<T> Constructor()
        {
            Type t = typeof(T);
            if (t == typeof(string))
                return Expression.Lambda<Func<T>>(Expression.Constant(string.Empty)).Compile();

            if (t.HasDefaultConstructor())
                return Expression.Lambda<Func<T>>(Expression.New(t)).Compile();
            return () => (T)FormatterServices.GetUninitializedObject(t);
        }
        static New(){}
        public static void Compile() { }
        public static PropertyInfo[] GetProperties<A>()
            where A : Attribute
        {
            return typeof(T).GetType().GetProperties().Where(x => Attribute.IsDefined(x, typeof(A), false)).ToArray();
        }
    }
    public static class TypeExtensions
    {
        internal static void Inject<T>(this T instance, PropertyInfo property, object value)
               where T : class
        {
            property.SetValue(instance, Convert.ChangeType(value, property.PropertyType), null);
        }
        
        internal static bool HasDefaultConstructor(this Type t)
        {
            return t.IsValueType || t.GetConstructor(Type.EmptyTypes) != null;
        }
    }
}
