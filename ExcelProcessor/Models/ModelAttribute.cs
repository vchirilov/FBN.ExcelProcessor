using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelProcessor.Models
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ModelAttribute: Attribute
    {
        public string Table { get; set; }
    }
}
