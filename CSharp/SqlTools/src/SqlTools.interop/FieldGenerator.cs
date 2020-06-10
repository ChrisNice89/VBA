using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace AccessCodeLib.Data.SqlTools.interop
{
    [ComVisible(true)]
    [Guid("F0181F29-D1F2-434C-9493-8B039E4B34DD")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".FieldGenerator")]
    public class FieldGenerator : SqlTools.FieldGenerator, IFieldGenerator
    {
        public IFieldList FromString(string fieldsString, string deliminator = ",")
        {
            try
            {
                var delim = deliminator.ToCharArray(0, 1)[0];
                var fields = FromString(fieldsString, delim);
                return ConvertFieldsToCom(fields);
            }
            catch (Exception ex)
            {
                throw new Exception("FromString:" + ex.Message, ex);
            }
        }

        public IFieldList FromArray(object fieldArray)
        {
            var objectArray = fieldArray as object[];
            if (objectArray == null)
                throw new ArrayTypeMismatchException("Array expected");

            if (objectArray is IField[])
            {
                var l = new FieldList();
                l.AddRange((IField[]) objectArray);
                return l;
            }

            try
            {
                var stringArray = objectArray.Select(field => field).Cast<string>().ToArray();
                var fields = base.FromArray(stringArray);
                return ConvertFieldsToCom(fields);
            }
            catch (Exception ex)
            {
                throw new ArrayTypeMismatchException(fieldArray + " not implemented\n" + ex);
            }
        }

        private static IFieldList ConvertFieldsToCom(IEnumerable<Common.Sql.IField> fields)
        {
            var fieldList = new FieldList { fields.Select(field => new AccessCodeLib.Data.SqlTools.interop.Field(field)).Cast<IField>().ToArray() };
            return fieldList;
        }

        public IFieldList FromNames(
            string field1, string field2 = null, string field3 = null, string field4 = null, string field5 = null,
            string field6 = null, string field7 = null, string field8 = null, string field9 = null, string field10 = null,
            string field11 = null, string field12 = null, string field13 = null, string field14 = null, string field15 = null,
            string field16 = null, string field17 = null, string field18 = null, string field19 = null, string field20 = null
            )
        {
            return FromArray(new [] {
                        field1, field2, field3, field4, field5, field6, field7, field8, field9, field10,
                        field11, field12, field13, field14, field15, field16, field17, field18, field19, field20});
        }

        internal static IField GetFieldFromObject(object field, FieldDataType dataType = FieldDataType._Unspecified)
        {
            return field is IField ? (IField)field : new Field((string)field, null, dataType);
        }

        public IField Field(string name, string source = "", FieldDataType dataType = FieldDataType._Unspecified)
        {
            return new Field(name, string.IsNullOrEmpty(source) ? null : new Sql.NamedSource(source), dataType);
        }

    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("F147176D-8B3B-40C4-984F-1BB4AC0552B1")]
    public interface IFieldGenerator : SqlTools.IFieldGenerator
    {
        IFieldList FromString(string FieldsString, string Deliminator = ",");

        IFieldList FromArray(object FieldNameArray);

        IFieldList FromNames(
            string Field1, string Field2 = null, string Field3 = null, string Field4 = null, string Field5 = null,
            string Field6 = null, string Field7 = null, string Field8 = null, string Field9 = null, string Field10 = null,
            string Field11 = null, string Field12 = null, string Field13 = null, string Field14 = null, string Field15 = null,
            string Field16 = null, string Field17 = null, string Field18 = null, string Field19 = null, string Field20 = null
            );

        IField Field(string name, string source = "", FieldDataType dataType = FieldDataType._Unspecified);
    }

}
