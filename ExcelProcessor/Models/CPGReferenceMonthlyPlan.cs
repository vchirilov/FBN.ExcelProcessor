using ExcelProcessor.Helpers;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "cpg_reference_monthly_plan")]
    public class CPGReferenceMonthlyPlan: IModel
    {
        [Order(1)] public int Year { get; set; }
        [Order(2)] public string YearType { get; set; }
        [Order(3)] public string Retailer { get; set; }
        [Order(4)] public string Banner { get; set; }
        [Order(5)] public string Country { get; set; }
        [Order(6)] public string EAN { get; set; }
        [Order(7)] public decimal Jan { get; set; }
        [Order(8)] public decimal Feb { get; set; }
        [Order(9)] public decimal Mar { get; set; }
        [Order(10)] public decimal Apr { get; set; }
        [Order(11)] public decimal May { get; set; }
        [Order(12)] public decimal Jun { get; set; }
        [Order(13)] public decimal Jul { get; set; }
        [Order(14)] public decimal Aug { get; set; }
        [Order(15)] public decimal Sep { get; set; }
        [Order(16)] public decimal Oct { get; set; }
        [Order(17)] public decimal Nov { get; set; }
        [Order(18)] public decimal Dec { get; set; }

        public bool IsEmpty()
        {
            if (Year == 0
                && YearType.IsNullOrEmpty() == true
                && Retailer.IsNullOrEmpty() == true
                && Banner.IsNullOrEmpty() == true
                && Country.IsNullOrEmpty() == true
                && EAN.IsNullOrEmpty() == true
                && Jan == 0m
                && Feb == 0m
                && Mar == 0m
                && Apr == 0m
                && May == 0m
                && Jun == 0m
                && Jul == 0m
                && Aug == 0m
                && Sep == 0m
                && Oct == 0m
                && Nov == 0m
                && Dec == 0m)
                return true;
            else
                return false;
        }
    }
}
