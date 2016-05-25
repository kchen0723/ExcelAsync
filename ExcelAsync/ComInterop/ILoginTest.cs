using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace ExcelAsync.ComInterop
{
    [Guid("9A65E021-CE21-401D-9641-94EF4AD8FB6C")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [ComVisible(true)]
    interface ILoginTest : IComLateBindingEvent
    {
        string GetAccessToken(string userName, string password, string clientId);
    }
}
