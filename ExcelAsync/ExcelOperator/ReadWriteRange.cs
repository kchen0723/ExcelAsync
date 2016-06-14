using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

using Microsoft.Office.Interop.Excel;

namespace ExcelAsync.ExcelOperator
{
    internal class ReadWriteRange
    {
        internal static Range WriteToRange(object[,] response, Range targetRange)
        {
            int rowsCount = response.GetLength(0);
            int columnsCount = response.GetLength(1);
            ExcelReference target = new ExcelReference(targetRange.Row - 1, targetRange.Row + rowsCount - 2, targetRange.Column - 1, targetRange.Column + columnsCount - 2);
            target.SetValue(response);
            return targetRange.Resize[rowsCount, columnsCount];
        }

        internal static object[,] ReadFromRange()
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
