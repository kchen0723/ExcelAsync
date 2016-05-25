using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAsyncWpf.EntityLogic
{
    //Methods in this class should be run by QueueToRunUIThreadHandler. We may read/write excel multiple times in this class according to logic
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
