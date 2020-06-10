using System.Collections.Generic;

namespace AccessCodeLib.Data.Common.Sql.Interfaces
{
    interface IValueList : IList<IValue>
    {
        void Add(params IValue[] values);
    }
}
