using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Converter
{
    public class NameConverterBase : INameConverter
    {
// ReSharper disable MemberCanBePrivate.Global
        protected const string SqlSelectAsStringFormat = " As ";
        protected const string SqlSourceAsAliasStringFormat = " As ";
// ReSharper restore MemberCanBePrivate.Global

        protected virtual string GetCheckedItemNameString(string name)
        {
            return name;
        }

        protected virtual string GetCheckedSourceNameString(INamedSource source)
        {
            return source.Name;
        }

// ReSharper disable VirtualMemberNeverOverriden.Global
        protected virtual string GetCheckedSourceNameString(ISourceAlias source)
        {
            return source.Source is INamedSource
                ? string.Concat(GetCheckedSourceNameString((INamedSource)(source.Source)), SqlSourceAsAliasStringFormat, GetCheckedItemNameString(source.Alias))
                : source.Alias;
        }

        public virtual string GenerateAliasNameString(IAlias alias)
        {
            return GetCheckedItemNameString(alias.Alias);
        }

        public virtual string GenerateSourceNameString(ISource source)
        {
            return GetCheckedSourceNameString(source);
        }

        public virtual string GenerateFieldNameString(ISource source, string fieldName)
        {
            var sourceString = GetCheckedFieldSourceNameString(source);

            return string.IsNullOrEmpty(sourceString) ? fieldName : string.Format("{0}.{1}", sourceString, fieldName);
        }

        protected virtual string GetCheckedFieldSourceNameString(ISource source)
        {
            if (source is ISourceAlias)
                return GetCheckedItemNameString(((ISourceAlias)source).Alias);

            if (source is INamedSource)
                return GetCheckedSourceNameString((INamedSource)source);

            if (source is ISubSelect) // SubSelect without Alias => null
                return null;

            throw new NotSupportedSourceException(source);
        }

        protected virtual string GetCheckedSourceNameString(ISource source)
        {
            if (source == null)
                return null;

            if (source is ISourceAlias)
                return GetCheckedSourceNameString((ISourceAlias)source);

            if (source is INamedSource)
                return GetCheckedSourceNameString((INamedSource)source);

            if (source is ISubSelect) // SubSelect without Alias => null
                return null;

            throw new NotSupportedSourceException(source);
        }

        public virtual string GenerateSelectFieldString(IField field)
        {
            var fieldString = GenerateFieldString(field);
            if (field is IAlias)
                fieldString = string.Concat(fieldString, SqlSelectAsStringFormat, GetCheckedAliasString((IAlias)field));

            return fieldString;
        }

        protected virtual string GetCheckedAliasString(IAlias alias)
        {
            return GetCheckedItemNameString(alias.Alias);
        }

        public virtual string GenerateFieldString(IField field)
        {
            return field.Source != null
                       ? GenerateFieldNameString(field.Source, GetCheckedItemNameString(field.Name))
                       : GetCheckedItemNameString(field.Name);
        }
// ReSharper restore VirtualMemberNeverOverriden.Global
    }
}
