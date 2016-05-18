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
            catch
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
            catch
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
            catch
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
            catch
            {
                return null;
            }
        }
    }
}
