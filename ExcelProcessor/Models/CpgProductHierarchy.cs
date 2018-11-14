using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "cpg_product_hierarchy")]
    public class CpgProductHierarchy : IModel
    {
        public string EAN { get; set; }
        public string CategoryGroup { get; set; }
        public string Subdivision { get; set; }
        public string Category { get; set; }
        public string Market { get; set; }
        public string Sector { get; set; }
        public string SubSector { get; set; }
        public string Segment { get; set; }
        public string ProductForm { get; set; }
        public string CPG { get; set; }
        public string BrandForm { get; set; }
        public string SizePackForm { get; set; }
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
