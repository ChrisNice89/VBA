using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.CSharp.RuntimeBinder;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;

namespace Skynet.Objects
{
    public abstract class TBase<TCustom, TValue> 
        where TCustom : class,IObject
    {
        protected readonly TValue _value;
        public TBase(TValue value)
        {
            _value = value;
        }

        #region mathematical operators
        public static bool operator <(TBase<TCustom, TValue> a, TBase<TCustom, TValue> b)
        {
            return Comparer<TValue>.Default.Compare(a._value, b._value) < 0;
        }
        public static bool operator >(TBase<TCustom, TValue> a, TBase<TCustom, TValue> b)
        {
            return !(a < b);
        }
        public static bool operator <=(TBase<TCustom, TValue> a, TBase<TCustom, TValue> b)
        {
            return (a < b) || (a == b);
        }
        public static bool operator >=(TBase<TCustom, TValue> a, TBase<TCustom, TValue> b)
        {
            return (a > b) || (a == b);
        }
        public static TCustom operator +(TBase<TCustom, TValue> a, TBase<TCustom, TValue> b)
        {
            return (dynamic)a._value + b._value;
        }
        public static TCustom operator -(TBase<TCustom, TValue> a, TBase<TCustom, TValue> b)
        {
            return ((dynamic)a._value - b._value);
        }
        public static TCustom operator *(TBase<TCustom, TValue> a, TBase<TCustom, TValue> b)
        {
            return ((dynamic)a._value * b._value);
        }
        public static TCustom operator /(TBase<TCustom, TValue> a, TBase<TCustom, TValue> b)
        {
            return ((dynamic)a._value / b._value);
        }
        #endregion mathematical operators
        #region logical operators
        public static bool operator ==(TBase<TCustom, TValue> a, TBase<TCustom, TValue> b)
        {
            return a.Equals((object)b);
        }
        public static bool operator !=(TBase<TCustom, TValue> a, TBase<TCustom, TValue> b)
        {
            return !(a == b);
        }
        protected bool Equals(TBase<TCustom, TValue> other)
        {
            return EqualityComparer<TValue>.Default.Equals(_value, other._value);
        }
        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((TBase<TCustom, TValue>)obj);
        }
        protected bool IsRelatedTo(object other)
        {
            return typeof(TCustom).Equals(other.GetType());
            //return other.GetType().IsInstanceOfType(typeof(TCustom));
        }
        #endregion logical operators
        #region CompareTo
        protected CompareResult CompareTo(TBase<TCustom, TValue> other)
        {
            return (CompareResult)Comparer<TValue>.Default.Compare(this._value, other._value);
        }
        protected CompareResult CompareTo(object obj)
        {
            //var other = obj as TCustom;
            //if (other != null)
            //{
            //    return CompareTo(other);
            //}
            //else
            //{
            //    return CompareResult.IsLower;
            //}
            return CompareTo(TryConvert<TBase<TCustom, TValue>>(obj));
        }
        protected static T TryConvert<T>(object input)
        {
            if (input is T)
            {
                return (T)input;
            }
            try
            {
                return (T)Convert.ChangeType(input, typeof(T));
            }
            catch (InvalidCastException)
            {
                return default(T);
            }
        }
        #endregion CompareTo
        public override string ToString()
        {
            return _value.ToString();
        }
        public override int GetHashCode()
        {
            return EqualityComparer<TValue>.Default.GetHashCode(_value);
        }
    }
    [ComVisible(true)]
    [Guid("25911AF5-3292-4676-9545-A9D7D965A20C"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IInterface
    {
        int X { get; set; }
        int Y { get; set; }
    }

    public static class IInterfaceTHelper
    {
        public static IInterface Add<T>(this IInterface a, IInterface b)
            where T : new()
        {
            var ret = (IInterface)new T();
            ret.X = a.X + b.X;
            ret.Y = a.Y + b.Y; 
            return ret;
        }
        public static T Add<T>(this T me, T other)
           where T : TBase<TString, String>, new ()
        {
            var ret = new T();
            return ret;
        }
    }
    class Foo : IInterface
    {
        public int X { get; set; }
        public int Y { get; set; }

        public static IInterface operator +(Foo a, IInterface b)
        {
            return a.Add<Foo>(b);
        }
    }

    class Bar : IInterface
    {
        public int X { get; set; }
        public int Y { get; set; }

        public static IInterface operator +(Bar a, IInterface b)
        {
            return a.Add<Bar>(b);
        }
    }

}
