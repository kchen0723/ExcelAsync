using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace ExcelAsync.ContextMenu
{
    public static class MenuManager
    {
        private static CommandBarButton button = null;
        private const string CELL_MENU = "Cell";
        private const string TABLE_MENU = "List Range Popup";

        public static void SetContentMenu()
        {
            Microsoft.Office.Interop.Excel.Range activeCell = ExcelApp.Application.ActiveCell;
            bool isThereHistory = false;
            ExcelWvvm.Entities.GoogleHistory result = ExcelManager.EntityManager.GetHistoryByRange(activeCell);
            if (result != null)
            {
                isThereHistory = true;
            }
            setContentMenuStatus(isThereHistory);
        }

        private static void setContentMenuStatus(bool isThereHistory)
        {
            CommandBars bars = ExcelApp.Application.CommandBars;
            CommandBar bar = bars[MenuManager.CELL_MENU];
            if (bar != null)
            {
                CommandBarControls controls = bar.Controls;
                try
                {
                    CommandBarControl control = controls[getRootMenuName(MenuManager.CELL_MENU)];
                    if (control != null)
                    {
                        control.Visible = isThereHistory;
                    }
                }
                catch
                { }
            }
        }

        public static void AddContextMenus()
        {
            MenuManager.DeleteContextMenus();
            addContectMenu(MenuManager.CELL_MENU);
            addContectMenu(MenuManager.TABLE_MENU);
        }

        public static void addContectMenu(string menuType)
        {
            CommandBars bars = ExcelApp.Application.CommandBars;
            CommandBar bar = bars[menuType];
            if (bar != null)
            {
                CommandBarControls controls = bar.Controls;
                dynamic rootControl = controls.Add(Type: MsoControlType.msoControlPopup, Before: 1, Temporary: true);
                rootControl.Caption = getRootMenuName(menuType);
                rootControl.BeginGroup = true;
                rootControl.Tag = getRootMenuName(menuType);

                button = rootControl.Controls.Add(Type: MsoControlType.msoControlButton) as CommandBarButton;
                button.Caption = "Refresh";
                button.Tag = "Refresh";
                button.OnAction = "OnButtonClick";
                //button.Click += Button_Click;
            }
        }

        public static void DeleteContextMenus()
        {
            deleteContextMenu(MenuManager.CELL_MENU);
            deleteContextMenu(MenuManager.TABLE_MENU);
        }

        private static void deleteContextMenu(string menuType)
        {
            CommandBars bars = ExcelApp.Application.CommandBars;
            CommandBar bar = bars[menuType];
            if (bar != null)
            {
                CommandBarControls controls = bar.Controls;
                try
                {
                    CommandBarControl control = controls[getRootMenuName(menuType)];
                    if (control != null)
                    {
                        control.Delete();
                    }
                }
                catch
                { }
            }
        }

        private static string getRootMenuName(string menuType)
        {
            string result = "ExcelAsync";
            if (string.Compare(menuType, CELL_MENU, true) == 0)
            {
                result = result + " Cell";
            }
            else
            {
                result = result + " Table";
            }
            return result;
        }

        private static void Button_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Microsoft.Office.Interop.Excel.Range activeCell = ExcelApp.Application.ActiveCell;
            ExcelWvvm.Entities.GoogleHistory history = ExcelManager.EntityManager.GetHistoryByRange(activeCell);
            if (history != null)
            {
                history.OnRetrievedDataHandler = History_OnRetrievedData;
                history.ExecuteAsync();
            }
        }

        private static void History_OnRetrievedData(object arg1, object arg2)
        {
            object[,] result = arg2 as object[,];
            ExcelManager.EntityManager.WriteToRange(result, arg1 as ExcelWvvm.Entities.GoogleHistory);
        }

        [ExcelDna.Integration.ExcelFunction(IsHidden = true)]
        public static void OnButtonClick()
        {
            System.Windows.MessageBox.Show("Hello from context menu");
            bool isCancel = false;
            Button_Click(null, ref isCancel);
        }
    }
}