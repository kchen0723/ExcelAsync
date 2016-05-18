using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace ExcelAsyncWpf.ExcelOperator
{
    public class CustomPropertyManager
    {
        public static dynamic GetWorkBookCustomProperty(Workbook wk, string Name)
        {
            dynamic cps = wk.CustomDocumentProperties;
            dynamic item = null;
            dynamic result = null;
            for (int i = 1; i <= cps.Count; i++) //VB start from 1
            {
                item = cps.Item(i);
                if (string.Compare(item.Name, Name, true) == 0)
                {
                    result = item;
                    break;
                }
            }
            return result;
        }

        public static string GetWorkBookPropertyString(Workbook wk, string Name)
        {
            string result = string.Empty;
            dynamic cp = GetWorkBookCustomProperty(wk, Name);
            if (cp != null)
            {
                result = cp.Value;
            }
            return result;
        }

        public static void SetWorkBookProperty(Workbook wk, string propertyName, string propertyValue)
        {
            dynamic cp = GetWorkBookCustomProperty(wk, propertyName);
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
    }
}
