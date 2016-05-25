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

        public static void AddContentMenu()
        {
            CommandBar cellMenu = ExcelApp.Application.CommandBars["Cell"];
            //There is a bug in below line. One dot is good, too dots are bad.
            //CommandBarButton button = cellMenu.Controls.Add(Type: MsoControlType.msoControlButton, Before: cellMenu.Controls.Count, Temporary: true) as CommandBarButton;
            //button.Caption = "Test Button";
            //button.Tag = "Test Button";
            //button.OnAction = "OnButtonClick";
            button = cellMenu.Controls.Add(Type: MsoControlType.msoControlButton, Before: cellMenu.Controls.Count, Temporary: true) as CommandBarButton;
            button.Caption = "Test Button";
            button.Click += Button_Click;
        }

        private static void Button_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            System.Windows.MessageBox.Show("Hello from context menu");
        }

        public static void OnButtonClick()
        {
            System.Windows.MessageBox.Show("Hello from context menu");
        }
    }
}
