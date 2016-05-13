using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace ExcelAsyncWpf.ExcelOperator
{
    public class ReadWriteRange
    {
        public static bool WriteToRange(string[,] response)
        {
            ExcelReference sheet2 = XlCall.Excel(XlCall.xlSheetId, "Sheet2") as ExcelReference;
            int rowsCount = response.GetLength(0);
            int columnsCount = response.GetLength(1);
            ExcelReference target = new ExcelReference(0, rowsCount - 1, 0, columnsCount - 1, sheet2.SheetId);
            return target.SetValue(response);
        }

        public static object[,] ReadFromRange()
        {
            object[,] result = null;
            ExcelReference selection = new ExcelReference(0, 4, 0, 2);
            object selectContent = selection.GetValue();
            if (selectContent is object[,])
            {
                result = selectContent as object[,];
            }
            else if (selectContent is double)
            {
                result = new object[,] { { selectContent } };
            }

            return result;
        }
    }
}
