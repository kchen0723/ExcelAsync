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
        public static object TestWebDownloadString(string url)
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

        [ExcelFunction(Description = "TestGreeting you")]
        public static string TestGreeting(string name)
        {
            return "Hello: " + name + " at " + DateTime.Now.ToString();
        }
    }
}
