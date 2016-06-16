using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelWvvm.Entities;

using Microsoft.Office.Interop.Excel;

namespace ExcelAsync.MainThreadLogic
{
    internal class EntityOperatior
    {
        internal static GoogleHistory GetHistoryByRange(Range targetRange)
        {
            ExcelUIThreadProtecter.CheckIsExcelUIMainThread();
            GoogleHistory result = null;
            string rangeName = MainThreadLogic.RangeManager.GetRangeName(targetRange);
            if (string.IsNullOrEmpty(rangeName) == false)
            {
                result = ExcelWvvm.Entities.GoogleHistories.GetByRangeName(rangeName);
            }
            return result;
        }

        internal static void ShowRefreshingComment(GoogleHistory history, string commnet = "Refreshing...")
        {
            ExcelUIThreadProtecter.CheckIsExcelUIMainThread();
            Worksheet ws = ExcelApp.Application.ActiveSheet;
            Range targetRange = MainThreadLogic.RangeManager.GetRange(ws, history.RangeName);
            if (targetRange != null)
            {
                targetRange = targetRange.Cells[1, 1];
                Comment refreshingComment = targetRange.Comment;
                if (refreshingComment == null)
                {
                    refreshingComment = targetRange.AddComment();
                }
                refreshingComment.Shape.TextFrame.AutoSize = true;
                refreshingComment.Shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                //refreshingComment.Shape.Fill.ForeColor.RGB =RGB(220, 220, 220);
                refreshingComment.Shape.Fill.OneColorGradient(Microsoft.Office.Core.MsoGradientStyle.msoGradientDiagonalUp, 1, (float)0.4);
                refreshingComment.Visible = true;
                refreshingComment.Text(commnet);
            }
        }
    }
}
