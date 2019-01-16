using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace ExcelProcessor.Helpers
{
    public class ApplicationError : Exception
    {
        private ApplicationError(string origin, string method, string text)
        {
            CallerClass = origin;
            CallerMethod = method;
            Text = text;
        }        

        public static ApplicationError Create(string text)
        {
            StackTrace stackTrace = new StackTrace();
            var methodBase = stackTrace.GetFrame(1).GetMethod();
            var methodName = methodBase.Name;
            var className = methodBase.ReflectedType.Name;

            return new ApplicationError(className, methodName, text);
        }

        public string CallerClass { get; set; }
        public string CallerMethod { get; set; }
        public string Text { get; set; }

        public override string ToString()
        {
            return $"{Text}".Truncate(250);            
        }

        public string GetStackTrace()
        {
            return $"Class:{CallerClass}, Method:{CallerMethod}, StackTrace:{StackTrace}".Truncate(250);
        }
    }
}
