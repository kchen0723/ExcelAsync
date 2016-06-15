﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelWvvm.Entities;

using Microsoft.Office.Interop.Excel;

namespace ExcelAsync.ExcelManager
{
    public class EntityManager
    {
        static object[,] m_result = null;
        static GoogleHistory m_history = null;
        static string m_comment = null;
        public static void WriteToRange(object[,] result, GoogleHistory history)
        {
            m_result = result;
            m_history = history;
            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(writeRangeToExcel);
        }

        private static void writeRangeToExcel()
        {
            Range result = null;
            if (string.IsNullOrEmpty(m_history.RangeName) == false)
            {
                Worksheet ws = ExcelApp.Application.ActiveSheet;
                result = ExcelOperator.RangeManager.GetRange(ws, m_history.RangeName);
                ExcelOperator.RangeManager.DeleteName(ws, m_history.RangeName);
            }
            if (result == null)
            {
                result = ExcelApp.Application.ActiveCell;
            }
            result = ExcelOperator.ReadWriteRange.WriteToRange(m_result, result);
            setDateFormat(result, m_result.GetLength(1) + 1);
            m_history.RangeName = "kissingerTest1" + DateTime.Now.ToString("yyyyMMddHHmmss");
            result.Name = m_history.RangeName;
            ExcelWvvm.Entities.GoogleHistories.GetAllHistories[m_history.InstanceId] = m_history;
        }

        private static void setDateFormat(Range targetRange, int rows)
        {
            targetRange = targetRange.Resize[rows, 1];
            targetRange.NumberFormat = "yyyy-MM-dd";
        }

        public static GoogleHistory GetHistoryByRange(Range targetRange)
        {
            GoogleHistory result = null;
            string rangeName = ExcelOperator.RangeManager.GetRangeName(targetRange);
            if (string.IsNullOrEmpty(rangeName) == false)
            {
                result = ExcelWvvm.Entities.GoogleHistories.GetByRangeName(rangeName);
            }
            return result;
        }

        public static void ShowRefreshingComment(GoogleHistory history, string commnet = "Refreshing...")
        {
            m_history = history;
            m_comment = commnet;
            ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(showRefreshingComment);
        }

        private static void showRefreshingComment()
        {
            Worksheet ws = ExcelApp.Application.ActiveSheet;
            Range targetRange = ExcelOperator.RangeManager.GetRange(ws, m_history.RangeName);
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
                refreshingComment.Text(m_comment);
            }
        }
    }
}