using AccessCodeLib.Data.Common.Sql;
using System;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class BetweenValue : IBetweenValue
    {
        public BetweenValue(IValue firstValue, IValue secondValue)
        {
            FirstValue = firstValue;
            SecondValue = secondValue;
        }

        public BetweenValue(string firstValue, string secondValue)
        {
            FirstValue = new TextValue(firstValue);
            SecondValue = new TextValue(secondValue);
        }

        public BetweenValue(int firstValue, int secondValue)
        {
            FirstValue = new NumericValue<int>(firstValue);
            SecondValue = new NumericValue<int>(secondValue);
        }
        
        public IValue FirstValue { get; private set; }
        public IValue SecondValue { get; private set;}

        public Type TypeOfValue { get { return typeof(IBetweenValue); } }
        object IValue.Value { get { return new [] { FirstValue, SecondValue }; } }
    }
}
