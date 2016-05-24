using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAsyncWvvm
{
    //Methods in this class should be run by QueueToRunUIThreadHandler. We may read/write excel multiple times in this class according to logic
    public class EntityManager
    {
        public static void WriteToRange(object[,] result)
        {
            if (ExcelHandler.WriteToRangeHandler != null)
            {
                ExcelHandler.WriteToRangeHandler(result);
            }
        }
    }
}
