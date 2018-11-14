using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "cpg_product_hierarchy")]
    public class CpgProductHierarchy : IModel
    {
        [Order(1)]
        public string EAN { get; set; }
        [Order(2)]
        public string CategoryGroup { get; set; }
        [Order(3)]
        public string Subdivision { get; set; }
        [Order(4)]
        public string Category { get; set; }
        [Order(5)]
        public string Market { get; set; }
        [Order(6)]
        public string Sector { get; set; }
        [Order(7)]
        public string SubSector { get; set; }
        [Order(8)]
        public string Segment { get; set; }
        [Order(9)]
        public string ProductForm { get; set; }
        [Order(10)]
        public string CPG { get; set; }
        [Order(11)]
        public string BrandForm { get; set; }
        [Order(12)]
        public string SizePackForm { get; set; }
        [Order(13)]
        public string SizePackFormVariant { get; set; }

        public bool IsEmpty()
        {
            if (EAN == null
                && CategoryGroup == null
                && Subdivision == null
                && Category == null
                && Market == null
                && Sector == null
                && SubSector == null
                && Segment == null
                && ProductForm == null
                && CPG == null
                && BrandForm == null
                && SizePackForm == null
                && SizePackFormVariant == null)
                return true;
            else
                return false;
        }
    }
}
