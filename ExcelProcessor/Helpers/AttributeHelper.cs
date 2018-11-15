﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelProcessor.Helpers
{
    public static class AttributeHelper
    {
        public static IOrderedEnumerable<PropertyInfo> GetSortedProperties<T>()
        {
            return typeof(T)
              .GetProperties()
              .OrderBy(p => ((OrderAttribute)p.GetCustomAttribute(typeof(OrderAttribute), false)).Key);
        }

        public static IOrderedEnumerable<PropertyInfo> GetSortedProperties(object obj)
        {
            return obj.GetType()
              .GetProperties()
              .OrderBy(p => ((OrderAttribute)p.GetCustomAttribute(typeof(OrderAttribute), false)).Key);
        }
    }
}