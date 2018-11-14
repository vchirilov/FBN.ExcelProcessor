using System;
using System.Runtime.CompilerServices;

namespace ExcelProcessor
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
    public sealed class OrderAttribute : Attribute
    {
        public OrderAttribute(int key = 0)
        {
            Key = key;
        }

        public int Key { get; }
    }

    [AttributeUsage(AttributeTargets.Class)]
    public class ModelAttribute : Attribute
    {
        public string Table { get; set; }
    }
}
