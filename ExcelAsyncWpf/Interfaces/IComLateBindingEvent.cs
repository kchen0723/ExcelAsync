using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace ExcelAsyncWpf.Interfaces
{
    [Guid("FEC822FE-D0AC-413B-8191-451238558256")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [ComVisible(true)]
    interface IComLateBindingEvent
    {
        object ComConsumerObject { get; set; }
        string ComEventName { get; set; }
        void AttachEvent(object comConsumer, string comEventName);
    }
}
