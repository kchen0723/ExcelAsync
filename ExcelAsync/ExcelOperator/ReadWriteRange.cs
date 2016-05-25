using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace ExcelAsync.ExcelOperator
{
    internal class ReadWriteRange
    {
        internal static bool WriteToRange(object[,] response)
        {
            //ExcelReference sheet2 = XlCall.Excel(XlCall.xlSheetId, "Sheet2") as ExcelReference;
            Microsoft.Office.Interop.Excel.Range activeCell = ExcelApp.Application.ActiveCell;
            int rowsCount = response.GetLength(0);
            int columnsCount = response.GetLength(1);
            //ExcelReference target = new ExcelReference(0, rowsCount - 1, 0, columnsCount - 1, sheet2.SheetId);
            string address = activeCell.Address;
            ExcelReference target = new ExcelReference(activeCell.Row, activeCell.Row + rowsCount - 1, activeCell.Column, activeCell.Column + columnsCount - 1);
            return target.SetValue(response);
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
