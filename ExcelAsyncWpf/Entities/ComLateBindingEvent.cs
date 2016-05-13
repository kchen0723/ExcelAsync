using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using ExcelAsyncWpf.Interfaces;

namespace ExcelAsyncWpf.Entities
{
    [Guid("D68FCC41-8078-4154-B1FD-A97AF5110737")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    [ProgId("ExcelAsyncWpf.ComLateBindingEvent")]
    public partial class ComLateBindingEvent : IComLateBindingEvent
    {
        private object m_ComConsumerObject;

        private string m_ComEventName;

        public void AttachEvent(object comConsumer, string comEventName)
        {
            this.m_ComConsumerObject = comConsumer;
            this.m_ComEventName = comEventName;
        }

        public void TriggerComEvent(object[] eventArgs)
        {
            if (this.m_ComConsumerObject != null && string.IsNullOrEmpty(this.m_ComEventName) == false)
            {
                Type comClassType = this.m_ComConsumerObject.GetType();
                comClassType.InvokeMember(this.m_ComEventName, System.Reflection.BindingFlags.InvokeMethod, null, this.m_ComConsumerObject, eventArgs);
            }
        }
    }
}
