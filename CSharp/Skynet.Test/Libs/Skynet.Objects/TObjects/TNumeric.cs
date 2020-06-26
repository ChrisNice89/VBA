using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Skynet.Objects
{
    public class TNumeric : TBase<TNumeric, int>, INumeric
    {
    #region constructors
    public TNumeric() : base(0) { }
    public TNumeric(int Value) : base(Value){ }
    #endregion constructors
    #region implicit operator 
    public static implicit operator TNumeric(int value) { return new TNumeric(value); }
    public static implicit operator int(TNumeric TCustom) { return TCustom._value; }
    #endregion implicit operator 
    #region Com
    int INumeric.Value => this._value;
    IObject INumeric.Object => this;
    #region IObject
    bool IObject.Equals(IObject other) => base.Equals(other);
    CompareResult IObject.CompareTo(IObject other) => base.CompareTo(other);
    bool IObject.IsRelatedTo(IObject other) => base.IsRelatedTo(other);
    string IObject.ToString() => base.ToString();
        #endregion IObject
        #endregion com
    }
    [Guid("F3FDC204-C03C-42B8-B64B-3B57917B17F7"), ComVisible(true),
    InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface INumeric : IObject
    {
        IObject Object { get; }
        int Value { get; }
    }
}
