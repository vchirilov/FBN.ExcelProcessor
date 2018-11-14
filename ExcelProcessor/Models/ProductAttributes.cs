using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "product_attributes")]
    public class ProductAttributes: IModel
    {
        public string Type { get; set; }
        public string EAN { get; set; }
        public string Brand { get; set; }
        public decimal PackSize { get; set; }
        public string PackSizeUnit { get; set; }
        public string MultiPack { get; set; }
        public int UnitsPerPack { get; set; }
        public string PackageType { get; set; }
        public string Form { get; set; }
        public string TargetUser { get; set; }
        public string TargetArea { get; set; }
        public string Variant { get; set; }
        public string NielsenCategory { get; set; }
        public string Priority { get; set; }

        public bool IsEmpty()
        {
            if (Type == null 
                && EAN == null 
                && Brand == null 
                && PackSize == 0m
                && PackSizeUnit == null 
                && MultiPack == null 
                && UnitsPerPack == 0
                && PackageType == null
                && Form == null
                && TargetUser == null 
                && TargetArea == null
                && Variant == null
                && NielsenCategory == null 
                && Priority == null)
                return true;
            else
                return false;
        }
    }
}
