using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelWvvm.Interfaces;

namespace ExcelWvvm.Entities
{
    public class GoogleHistory : IGoogleHistory
    {
        public string SecurityId { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
    }
}
