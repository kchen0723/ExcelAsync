using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;

using ExcelDna.Integration;
using ExcelDna.AsyncSample;

namespace ExcelAsyncWpf
{
    public class TestFormulas
    {
        // The better example is to use HttpClient in System.Net.Http, 
        // which allows cancellation using a CancellationToken...
        [ExcelFunction(Description = "Test downloading async")]
        public static object TestWebDownloadString(string url)
        {
            if (string.IsNullOrEmpty(url) == false)
            {
                object result = ExcelTaskUtil.RunAsTask("asyncDownloadString", url,
                    () => new WebClient().DownloadString(url));
                if (result.GetType() == typeof(string))
                {
                    return result;
                }
                else
                {
                    return "processing";
                }
            }
            else
            {
                return string.Empty;
            }
        }

        [ExcelFunction(Description = "TestGreeting you")]
        public static string TestGreeting(string name)
        {
            ExcelOperator.ReadWriteRange.ReadFromRange();
            return "Hello: " + name + " at " + DateTime.Now.ToString();
        }
    }
}
