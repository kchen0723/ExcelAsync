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

        public static GoogleHistory GetHistoryByRange(Range targetRange)
        {
            GoogleHistory result = null;
            string rangeName = MainThreadLogic.RangeManager.GetRangeName(targetRange);
            if (string.IsNullOrEmpty(rangeName) == false)
            {
                result = ExcelWvvm.Entities.GoogleHistories.GetByRangeName(rangeName);
            }
            return result;
        }

        public static void ShowRefreshingComment(GoogleHistory history, string commnet = "Refreshing...")
        {
            Tuple<GoogleHistory, string> parameters = new Tuple<GoogleHistory, string>(history, commnet);
            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(showRefreshingComment, parameters);
        }

        private static void showRefreshingComment(object parameters)
        {
            Tuple<GoogleHistory, string> para = parameters as Tuple<GoogleHistory, string>;
            Worksheet ws = ExcelApp.Application.ActiveSheet;
            Range targetRange = MainThreadLogic.RangeManager.GetRange(ws, para.Item1.RangeName);
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
                refreshingComment.Text(para.Item2);
            }
        }
    }
}
