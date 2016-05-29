using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Excel;

namespace ExcelAsync.ExcelManager
{
    public class EntityManager
    {
        static object[,] m_result = null;
        public static bool WriteToRange(object[,] result)
        {
            m_result = result;
            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(writeRangeToExcel);
            return true;
        }

        private static void writeRangeToExcel()
        {
            Range result = ExcelOperator.ReadWriteRange.WriteToRange(m_result);
            result.Name = "kissingerTest1";
        }
    }
}
