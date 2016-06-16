using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelWvvm.Entities;

using Microsoft.Office.Interop.Excel;

namespace ExcelAsync.NonMainThreadLogic
{
    public class EntityManager
    {
        public static void WriteToRange(object[,] result, GoogleHistory history)
        {
            Tuple<GoogleHistory, object[,]> parameters = new Tuple<GoogleHistory, object[,]>(history, result);
            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(writeRangeToExcel, parameters);
        }

        private static void writeRangeToExcel(object parameters)
        {
            Tuple<GoogleHistory, object[,]> para = parameters as Tuple<GoogleHistory, object[,]>;
            Range result = null;
            if (string.IsNullOrEmpty(para.Item1.RangeName) == false)
            {
                Worksheet ws = ExcelApp.Application.ActiveSheet;
                result = MainThreadLogic.RangeManager.GetRange(ws, para.Item1.RangeName);
                MainThreadLogic.RangeManager.DeleteName(ws, para.Item1.RangeName);
            }
            if (result == null)
            {
                result = ExcelApp.Application.ActiveCell;
            }
            result = MainThreadLogic.ReadWriteRange.WriteToRange(para.Item2, result);
            setDateFormat(result, para.Item2.GetLength(1) + 1);
            para.Item1.RangeName = "kissingerTest1" + DateTime.Now.ToString("yyyyMMddHHmmss");
            result.Name = para.Item1.RangeName;
            ExcelWvvm.Entities.GoogleHistories.GetAllHistories[para.Item1.InstanceId] = para.Item1;
        }

        private static void setDateFormat(Range targetRange, int rows)
        {
            targetRange = targetRange.Resize[rows, 1];
            targetRange.NumberFormat = "yyyy-MM-dd";
        }
    }
}
