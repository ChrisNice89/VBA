using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Skynet.Objects
{
    public class TDate : TBase<TDate, DateTime>, IDate
    {
        #region constructors
        public TDate() : base(DateTime.Now) { }
        public TDate(DateTime Value) : base(Value) { }
        #endregion constructors
        #region implicit operator 
        public static implicit operator TDate(DateTime value) { return new TDate(value); }
        public static implicit operator DateTime(TDate TCustom) { return TCustom._value; }
        #endregion implicit operator 
        #region Com
        DateTime IDate.Value => this._value;
        IObject IDate.Object => this;
        #region IObject
        bool IObject.Equals(IObject other) => base.Equals(other);
        CompareResult IObject.CompareTo(IObject other) => base.CompareTo(other);
        bool IObject.IsRelatedTo(IObject other) => base.IsRelatedTo(other);
        string IObject.ToString() => base.ToString();
        #endregion IObject
        #endregion com
    }
    [Guid("C306CA30-61BF-4DA9-AB90-ABE0EFBCAAEC"), ComVisible(true),
    InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDate : IObject
    {
        IObject Object { get; }
        DateTime Value { get; }
    }
}
