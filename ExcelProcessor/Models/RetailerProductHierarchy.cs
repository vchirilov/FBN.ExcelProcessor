using ExcelProcessor.Helpers;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "fbn_staging.retailer_product_hierarchy")]
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
            if (Retailer.IsNullOrEmpty() == true &&
                Banner.IsNullOrEmpty() == true &&
                Country.IsNullOrEmpty() == true &&
                EAN.IsNullOrEmpty() == true &&
                CategoryGroup.IsNullOrEmpty() == true &&
                SubDivision.IsNullOrEmpty() == true &&
                Category.IsNullOrEmpty() == true &&
                Market.IsNullOrEmpty() == true &&
                Sector.IsNullOrEmpty() == true &&
                SubSector.IsNullOrEmpty() == true &&
                Segment.IsNullOrEmpty() == true &&
                ProductForm.IsNullOrEmpty() == true &&
                CPG.IsNullOrEmpty() == true &&
                BrandForm.IsNullOrEmpty() == true &&
                SizePackForm.IsNullOrEmpty() == true &&
                SizePackFormVariant.IsNullOrEmpty() == true)
                return true;
            else
                return false;
        }
    }
}
