using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace ExcelAsync.ExcelOperator
{
    internal static class RangeManager
    {
        internal static Name GetName(Workbook wk, string rangeName)
        {
            ExcelUIThreadProtecter.CheckIsExcelUIMainThread();
            Name result = null;
            if (wk != null && string.IsNullOrEmpty(rangeName) == false)
            {
                try
                {
                    result = wk.Names.Item(rangeName);
                }
                catch  //ignore error here
                {
                }
            }
            return result;
        }

        internal static Range GetRangeOfName(Name name)
        {
            ExcelUIThreadProtecter.CheckIsExcelUIMainThread();
            Range result = null;
            if (name != null)
            {
                try
                {
                    result = name.RefersToRange;
                }
                catch //ignore error here
                {
                }
            }
            return result;
        }

        internal static ListObject GetListObject(Worksheet sheet, string rangeName)
        {
            ExcelUIThreadProtecter.CheckIsExcelUIMainThread();
            ListObject result = null;
            if (sheet != null && string.IsNullOrEmpty(rangeName) == false)
            {
                try
                {
                    result = sheet.ListObjects[rangeName];
                }
                catch  //ignore error here
                {
                }
            }
            return result;
        }

        internal static Range GetListObjectRange(ListObject listObject)
        {
            ExcelUIThreadProtecter.CheckIsExcelUIMainThread();
            Range result = null;
            if (listObject != null)
            {
                try
                {
                    result = listObject.Range;
                }
                catch //ignore error here
                {
                }
            }
            return result;
        }

        internal static Range GetRange(Worksheet ws, string rangeName)
        {
            ExcelUIThreadProtecter.CheckIsExcelUIMainThread();
            Range result = null;
            if (ws != null && string.IsNullOrEmpty(rangeName) == false)
            {
                Workbook wk = ws.Parent;
                Name namedName = GetName(wk, rangeName);
                if (namedName != null)
                {
                    result = GetRangeOfName(namedName);
                }

                if (result == null)
                {
                    ListObject lo = GetListObject(ws, rangeName);
                    if (lo != null)
                    {
                        result = GetListObjectRange(lo);
                    }
                }
            }
            return result;
        }

        internal static bool IsRangeOverlap(Range sourceRange, Range targetRange)
        {
            ExcelUIThreadProtecter.CheckIsExcelUIMainThread();
            bool result = false;
            if (sourceRange != null && targetRange != null)
            {
                Worksheet sourceSheet = sourceRange.Worksheet;
                Worksheet targetSheet = targetRange.Worksheet;
                Workbook sourceBook = sourceSheet.Parent;
                Workbook targetBook = targetSheet.Parent;
                if (string.Compare(sourceBook.Name, targetBook.Name, false) == 0)
                {
                    if (string.Compare(sourceSheet.Name, targetSheet.Name, false) == 0)
                    {
                        if (ExcelApp.Application.Intersect(sourceRange, targetRange) != null)
                        {
                            result = true;
                        }
                    }
                }
            }
            return result;
        }

        internal static Name GetName(Range targetRange)
        {
            ExcelUIThreadProtecter.CheckIsExcelUIMainThread();
            Name result = null;
            if (targetRange != null)
            {
                Worksheet ws = targetRange.Worksheet;
                Workbook wb = ws.Parent;
                Names allNames = wb.Names;
                Name item = null;
                Range itemRange = null;
                for (int i = 1; i <= allNames.Count; i++)      //VB start from 1
                {
                    item = allNames.Item(i);
                    itemRange = GetRangeOfName(item);
                    if (IsRangeOverlap(itemRange, targetRange) == true)
                    {
                        result = item;
                        break;
                    }
                }
            }
            return result;
        }

        internal static ListObject GetListObject(Range targetRange)
        {
            ExcelUIThreadProtecter.CheckIsExcelUIMainThread();
            ListObject result = null;
            if (targetRange != null)
            {
                Worksheet ws = targetRange.Worksheet;
                ListObjects allListObject = ws.ListObjects;
                ListObject item = null;
                Range itemRange = null;
                for (int i = 1; i <= allListObject.Count; i++)     //VB start from 1
                {
                    item = allListObject[i];
                    itemRange = GetListObjectRange(item);
                    if (IsRangeOverlap(itemRange, targetRange) == true)
                    {
                        result = item;
                        break;
                    }
                }
            }
            return result;
        }

        internal static string GetRangeName(Range targetRange)
        {
            ExcelUIThreadProtecter.CheckIsExcelUIMainThread();
            string result = null;
            if (targetRange != null)
            {
                Name namedRange = GetName(targetRange);
                if (namedRange != null)
                {
                    result = namedRange.Name;
                }

                if (string.IsNullOrEmpty(result) == true)
                {
                    ListObject lo = GetListObject(targetRange);
                    if (lo != null)
                    {
                        result = lo.Name;
                    }
                }
            }
            return result;
        }
    }
}
