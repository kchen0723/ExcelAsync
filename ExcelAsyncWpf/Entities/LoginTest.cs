using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using ExcelAsyncWpf.Interfaces;

namespace ExcelAsyncWpf.Entities
{
    [Guid("B40DD491-CAA7-45C1-8A39-40859B8E5000")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    [ProgId("ExcelAsyncWpf.LoginTest")]
    public partial class LoginTest : ComLateBindingEvent, ILoginTest
    {
        public string GetAccessToken(string userName, string password, string clientId)
        {
            this.TriggerComEvent(new object[] { DateTime.Now.ToString(), true, 78, "testingString" });
            return userName + " " + password + " " + clientId + DateTime.Now.ToString();
        }
    }
}
