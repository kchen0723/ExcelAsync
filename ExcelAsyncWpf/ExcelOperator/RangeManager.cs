using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace ExcelAsyncWpf.ExcelOperator
{
    public static class RangeManager
    {
        public static Name GetName(Workbook wk, string rangeName)
        {
            try
            {
                return wk.Names.Item(rangeName);
            }
            catch  //ignore error here
            {
                return null;
            }
        }

        public static Range GetNameRange(Name name)
        {
            try
            {
                return name.RefersToRange;
            }
            catch //ignore error here
            {
                return null;
            }
        }

        public static ListObject GetListObject(Worksheet sheet, string rangeName)
        {
            try
            {
                return sheet.ListObjects[rangeName];
            }
            catch  //ignore error here
            {
                return null;
            }
        }

        public static Range GetListObjectRange(ListObject listObject)
        {
            try
            {
                return listObject.Range;
            }
            catch //ignore error here
            {
                return null;
            }
        }

        public static Range GetRange(Worksheet ws, string rangeName)
        {
            Range result = null;
            Workbook wk = ws.Parent;
            Name namedName = GetName(wk, rangeName);
            if (namedName != null)
            {
                result = GetNameRange(namedName);
            }

            if (result == null)
            {
                ListObject lo = GetListObject(ws, rangeName);
                if (lo != null)
                {
                    result = GetListObjectRange(lo);
                }
            }
            return result;
        }
    }
}
