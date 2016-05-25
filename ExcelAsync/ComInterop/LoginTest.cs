using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace ExcelAsync.ComInterop
{
    [Guid("B40DD491-CAA7-45C1-8A39-40859B8E5000")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    [ProgId("ExcelAsyncWpf.LoginTest")]
    public partial class LoginTest : ILoginTest
    {
        private ComLateBindingEvent m_ComLateBindingEvent = new ComLateBindingEvent();
        public void AttachEvent(object comConsumer, string comEventName)
        {
            this.m_ComLateBindingEvent.AttachEvent(comConsumer, comEventName);
        }

        public string GetAccessToken(string userName, string password, string clientId)
        {
            this.m_ComLateBindingEvent.TriggerComEvent(new object[] { DateTime.Now.ToString(), true, 78, "testingString" });
            return userName + " " + password + " " + clientId + DateTime.Now.ToString();
        }
    }
}
