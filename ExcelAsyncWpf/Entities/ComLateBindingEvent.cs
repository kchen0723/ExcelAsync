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
        public object ComConsumerObject { get; set; }

        public string ComEventName { get; set; }

        public void AttachEvent(object comConsumer, string comEventName)
        {
            this.ComConsumerObject = comConsumer;
            this.ComEventName = comEventName;
        }

        protected void TriggerComEvent(object eventArgs)
        {
            if (this.ComConsumerObject != null && string.IsNullOrEmpty(this.ComEventName) == false)
            {
                Type comClassType = this.ComConsumerObject.GetType();
                comClassType.InvokeMember(this.ComEventName, System.Reflection.BindingFlags.InvokeMethod, null, this.ComConsumerObject, new object[] { eventArgs });
            }
        }
    }
}
