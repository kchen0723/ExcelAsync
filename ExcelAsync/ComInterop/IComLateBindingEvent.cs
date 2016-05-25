using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace ExcelAsync.ComInterop
{
    [Guid("FEC822FE-D0AC-413B-8191-451238558256")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [ComVisible(true)]
    interface IComLateBindingEvent
    {
        void AttachEvent(object comConsumer, string comEventName);
    }
}
