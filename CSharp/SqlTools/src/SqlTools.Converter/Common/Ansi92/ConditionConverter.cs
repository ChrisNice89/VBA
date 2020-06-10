using System.Text.RegularExpressions;

namespace AccessCodeLib.Data.SqlTools.Converter.Common.Ansi92
{
    class ConditionConverter : Converter.ConditionConverter
    {
        public ConditionConverter(INameConverter nameConvertor, IValueConverter valueConverter)
            : base(nameConvertor, valueConverter)
        {
        }

        private static readonly Regex ConditionWildcardMReplaceRegex = new Regex(@"(like [\'\""][^\[\*]*)\*", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private static readonly Regex ConditionWildcardMReReplaceRegex = new Regex(@"(like ['""][^\[])%", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private static readonly Regex ConditionWildcardSReplaceRegex = new Regex(@"(like [\'\""][^\]]*)\?", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private static readonly Regex ConditionWildcardSReReplaceRegex = new Regex(@"(like ['""][^\[])_", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        protected override string GetCheckedConditionString(string condition)
        {
            // * => %
            condition = ConditionWildcardMReReplaceRegex.Replace(condition, m => m.Groups[1] + "[%]" + m.Groups[2]);
            condition = ConditionWildcardMReplaceRegex.Replace(condition, m => m.Groups[1] + "%" + m.Groups[2]);

            // ? => _
            condition = ConditionWildcardSReReplaceRegex.Replace(condition, m => m.Groups[1] + "[_]" + m.Groups[2]);
            condition = ConditionWildcardSReplaceRegex.Replace(condition, m => m.Groups[1] + "_" + m.Groups[2]);

            return condition;
        }
    }
}