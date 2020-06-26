using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Skynet.DAO
{
    using Skynet.DAO.Attributes;
    using Skynet.DAO.Extensions;
    using Skynet.DAO.Reflection;
    using System.Collections.ObjectModel;
    using System.Data.SqlClient;
    using System.Reflection;

    //void Main()
    //{
    //    var Sql = "select top 10 * from TVChannel";
    //    SqlConnection connection = new SqlConnection();
        
    //    var query = connection.CreateQuery<TVChannel>();
    //    query.Prepare();
    //    IList<TVChannel> records = query.Execute(Sql);
    //    IList<TVChannel> records2 = connection.SqlQuery<TVChannel>(Sql);
    //}

    public class TVChannel
    {
        [SqlField(Name = "number")]
        public string number { get; set; }
        [SqlField(Name = "title")]
        public string title { get; set; }
        [SqlField(Name = "favoriteChannel")]
        public string favoriteChannel { get; set; }
        [SqlField(Name = "description")]
        public string description { get; set; }
        [SqlField(Name = "packageid")]
        public string packageid { get; set; }
        [SqlField(Name = "format")]
        public string format { get; set; }
    }
}


