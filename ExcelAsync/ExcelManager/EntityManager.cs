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
            Range result = ExcelOperator.ReadWriteRange.WriteToRange(m_result);
            m_history.RnageName = "kissingerTest1";
            result.Name = m_history.RnageName;
        }
    }
}
