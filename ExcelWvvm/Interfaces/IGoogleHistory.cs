using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelWvvm.Interfaces
{
    interface IGoogleHistory
    {
        string SecurityId { get; set; }
        DateTime StartDate { get; set; }
        DateTime EndDate { get; set; }
        string InstanceId { get; set; }
        string RnageName { get; set; }
        string SheetId { get; set; }
    }
}
