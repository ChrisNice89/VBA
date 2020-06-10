using System.Collections;
using System.Linq;
using System.Runtime.InteropServices;

namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [Guid("1B57EBDE-B250-4FAC-B50D-405FBDAE8418")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IFieldList))]
    [ProgId(Constants.ProgIdLibName + ".FieldList")]
    public class FieldList : Sql.FieldList, IFieldList
    {
        public void Add(IField field)
        {
            base.Add(field);
        }

        public void Add(params IField[] fields)
        {
            base.Add(fields);
        }

        public new object[] ToArray()
        {
            /*
            var fields = new object[Count];
            var i = 0;
            foreach (var fld in this)
            {
                fields[i++] = fld as IField;
            }
            return fields;
            */
            return this.Cast<object>().ToArray();
        }

        public IField Item(int index)
        {
            return (IField)this[index];
        }

        public IEnumerable Items()
        {
            return this.Cast<IField>();
        }

    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid("02226951-20A9-4FCF-8F38-317559F205EC")]
    public interface IFieldList : Common.Sql.IFieldList
    {
        [DispId(0)]
        IField Item(int Index);

        [DispId(-4)]
        IEnumerable Items();

        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_DISPATCH)]
        object[] ToArray();

        void Add(IField Field);
    }

}
