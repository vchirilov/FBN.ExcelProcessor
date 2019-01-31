using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [Model(Table = "bain_configuration.parameters")]
    public class Configuration
    {
        [Order(0)] public int PID { get; set; }
        [Order(1)] public string ParameterName { get; set; }
        [Order(2)] public string ParameterValue { get; set; }
        [Order(3)] public string ParameterValue2 { get; set; }
    }
}
