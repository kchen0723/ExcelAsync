using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelWvvm.Entities;

using Microsoft.Office.Interop.Excel;

namespace ExcelAsync.ExcelManager
{
    public class EntityManager
    {
        static object[,] m_result = null;
        static GoogleHistory m_history = null;
        public static GoogleHistory WriteToRange(object[,] result, GoogleHistory history)
        {
            m_result = result;
            m_history = history;
            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(writeRangeToExcel);
            return m_history;
        }

        private static void writeRangeToExcel()
        {
            Range result = null;
            if (string.IsNullOrEmpty(m_history.RangeName) == false)
            {
                Worksheet ws = ExcelApp.Application.ActiveSheet;
                result = ExcelOperator.RangeManager.GetRange(ws, m_history.RangeName);
                ExcelOperator.RangeManager.DeleteName(ws, m_history.RangeName);
            }
            if (result == null)
            {
                result = ExcelApp.Application.ActiveCell;
            }
            result = ExcelOperator.ReadWriteRange.WriteToRange(m_result, result);
            m_history.RangeName = "kissingerTest1" + DateTime.Now.ToString("yyyyMMddHHmmss");
            result.Name = m_history.RangeName;
        }

        public static GoogleHistory GetHistoryByRange(Range targetRange)
        {
            GoogleHistory result = null;
            string rangeName = ExcelOperator.RangeManager.GetRangeName(targetRange);
            if (string.IsNullOrEmpty(rangeName) == false)
            {
                result = ExcelWvvm.Entities.GoogleHistories.GetByRangeName(rangeName);
            }
            return result;
        }
    }
}