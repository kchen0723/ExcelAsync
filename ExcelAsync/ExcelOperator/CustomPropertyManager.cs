using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace ExcelAsync.ExcelOperator
{
    internal class CustomPropertyManager
    {
        private static dynamic GetWorkbookCustomProperty(Workbook wk, string name)
        {
            dynamic cps = wk.CustomDocumentProperties;
            dynamic item = null;
            dynamic result = null;
            for (int i = 1; i <= cps.Count; i++) //VB start from 1
            {
                item = cps.Item(i);
                if (string.Compare(item.Name, name, true) == 0)
                {
                    result = item;
                    break;
                }
            }
            return result;
        }

        internal static string GetWorkbookPropertyString(Workbook wk, string name)
        {
            ExcelUIThreadProtecter.CheckIsExcelUIMainThread();
            string result = string.Empty;
            dynamic cp = GetWorkbookCustomProperty(wk, name);
            if (cp != null)
            {
                result = cp.Value;
            }
            return result;
        }

        internal static void SetWorkbookProperty(Workbook wk, string propertyName, string propertyValue)
        {
            ExcelUIThreadProtecter.CheckIsExcelUIMainThread();
            dynamic cp = GetWorkbookCustomProperty(wk, propertyName);
            if (cp == null)
            {
                dynamic cps = wk.CustomDocumentProperties;
                cps.Add(propertyName, false, MsoDocProperties.msoPropertyTypeString, propertyValue);
            }
            else
            {
                cp.Value = propertyValue;
            }
        }

        private static dynamic GetWorksheetCusotmProperty(Worksheet ws, string name)
        {
            CustomProperties cps = ws.CustomProperties;
            dynamic result = null;
            foreach (dynamic item in cps)
            {
                if (string.Compare(item.Name, name, true) == 0)
                {
                    result = item;
                    break;
                }
            }
            return result;
        }

        internal static string GetWorksheetPropertyString(Worksheet ws, string name)
        {
            ExcelUIThreadProtecter.CheckIsExcelUIMainThread();
            string result = string.Empty;
            dynamic cp = GetWorksheetCusotmProperty(ws, name);
            if (cp != null)
            {
                result = cp.Value;
            }
            return result;
        }

        internal static void SetWorksheetProperty(Worksheet ws, string propertyName, string propertyValue)
        {
            ExcelUIThreadProtecter.CheckIsExcelUIMainThread();
            dynamic cp = GetWorksheetCusotmProperty(ws, propertyName);
            if (cp == null)
            {
                CustomProperties cps = ws.CustomProperties;
                cps.Add(propertyName, propertyValue);
            }
            else
            {
                cp.Value = propertyValue;
            }
        }
    }
}
