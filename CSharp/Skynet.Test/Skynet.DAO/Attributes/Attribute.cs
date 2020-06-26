using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Skynet.DAO.Attributes
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false)]
    [Serializable]
    public class SqlField : Attribute
    {
        public string Name = null;
        public string Alias = null;
    }

    public class Product
    {
        [SqlField(Name = "product_id")]
        public int ProductId { get; private set; }
        [SqlField(Name = "supplier_id")]
        public int SupplierId { get; private set; }
        [SqlField(Name = "name")]
        public string Name { get; private set; }
        [SqlField(Name = "price")]
        public decimal Price { get; private set; }
        [SqlField(Name = "total_stock")]
        public int Stock { get; private set; }
        [SqlField(Name = "pending_stock")]
        public int PendingStock { get; private set; }
    }
}

