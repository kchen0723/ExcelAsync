using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
            ExcelOperator.ReadWriteRange.WriteToRange(m_result);
        }
    }
}
