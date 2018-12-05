using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "retailer_product_hierarchy")]
    public class RetailerProductHierarchy: IModel
    {
        [Order(1)] public string Retailer { get; set; }
        [Order(2)] public string Banner { get; set; }
        [Order(3)] public string Country { get; set; }
        [Order(4)] public string EAN { get; set; }
        [Order(5)] public string CategoryGroup { get; set; }
        [Order(6)] public string SubDivision { get; set; }
        [Order(7)] public string Category { get; set; }
        [Order(8)] public string Market { get; set; }
        [Order(9)] public string Sector { get; set; }
        [Order(10)] public string SubSector { get; set; }
        [Order(11)] public string Segment { get; set; }
        [Order(12)] public string ProductForm { get; set; }
        [Order(13)] public string CPG { get; set; }
        [Order(14)] public string BrandForm { get; set; }
        [Order(15)] public string SizePackForm { get; set; }
        [Order(16)] public string SizePackFormVariant { get; set; }

        public bool IsEmpty()
        {
            if (Retailer == null &&
                Banner == null &&
                Country == null &&
                EAN == null &&
                CategoryGroup == null &&
                SubDivision == null &&
                Category == null &&
                Market == null &&
                Sector == null &&
                SubSector == null &&
                Segment == null &&
                ProductForm == null &&
                CPG == null &&
                BrandForm == null &&
                SizePackForm == null &&
                SizePackFormVariant == null)
                return true;
            else
                return false;
        }
    }
}
