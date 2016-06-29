﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using ExcelWvvm.ViewModel;

namespace ExcelWvvm.View
{
    /// <summary>
    /// Interaction logic for WinGoogleHistoryContainer.xaml
    /// </summary>
    public partial class WinGoogleHistoryContainer : Window
    {
        public WinGoogleHistoryContainer(WinGoogleHistoryContainerViewModel vm)
        {
            InitializeComponent();
            this.DataContext = vm;
        }
    }
}
