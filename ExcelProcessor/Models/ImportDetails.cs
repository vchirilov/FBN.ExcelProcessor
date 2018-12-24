using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "fbn_import.import_details")]
    public class ImportDetails
    {
        [Order(0)] public Guid Uuid { get; set; }
        [Order(1)] public string FileName { get; set; }
        [Order(2)] public string User { get; set; }
        [Order(3)] public string ImportType { get; set; }
        [Order(4)] public int Year { get; set; }
        [Order(5)] public int Month { get; set; }
    }
}
