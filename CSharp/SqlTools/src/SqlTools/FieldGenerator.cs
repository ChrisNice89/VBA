using System;
using System.Collections.Generic;
using System.Linq;
using AccessCodeLib.Data.Common.Sql;
using AccessCodeLib.Data.SqlTools.Sql;

namespace AccessCodeLib.Data.SqlTools
{
    public class FieldGenerator : IFieldGenerator
    {
    	/// <summary>
    	/// \~english Creates an IField from a single String containing FieldNames delimited with delimiter
    	/// \~german  Erzeugt ein IField-Array aus einem String, der die Feldnamen mit Trennzeichen getrennt enthält \~
    	/// </summary>
    	/// <param name="fieldsString">String with FieldNames</param>
    	/// <param name="delimiter">default: ','</param>
    	/// <returns>IField[]</returns>
        public IField[] FromString(string fieldsString, char delimiter = ',')
        {
            var fieldNames = fieldsString.Split(new[] { delimiter }, StringSplitOptions.RemoveEmptyEntries);
            return FromArray(fieldNames.Select(fieldName => fieldName.Trim()).Where(n => !string.IsNullOrEmpty(n)).ToArray());
        }

        public IField[] FromArray(IEnumerable<string> fieldNames)
        {
            var fieldList = new FieldList();

            foreach (var fieldName in fieldNames)
            {
                fieldList.Add(fieldName);
            }
            return fieldList.ToArray();
        }
    }
}
