using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "product_attributes")]
    public class ProductAttributes: IModel
    {
        [Order(1)]
        public string Type { get; set; }
        [Order(2)]
        public string EAN { get; set; }
        [Order(3)]
        public string Brand { get; set; }
        [Order(4)]
        public decimal PackSize { get; set; }
        [Order(5)]
        public string PackSizeUnit { get; set; }
        [Order(6)]
        public string MultiPack { get; set; }
        [Order(7)]
        public int UnitsPerPack { get; set; }
        [Order(8)]
        public string PackageType { get; set; }
        [Order(9)]
        public string Form { get; set; }
        [Order(10)]
        public string TargetUser { get; set; }
        [Order(11)]
        public string TargetArea { get; set; }
        [Order(12)]
        public string Variant { get; set; }
        [Order(13)]
        public string NielsenCategory { get; set; }
        [Order(14)]
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
