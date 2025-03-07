﻿using ExcelCompare.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelCompare
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            MainWindowViewModel mainWindowViewModel = new();
            BindingOperations.EnableCollectionSynchronization(mainWindowViewModel.ExcelInfoOrigin.Sheets, mainWindowViewModel.ExcelInfoOrigin);
            BindingOperations.EnableCollectionSynchronization(mainWindowViewModel.ExcelInfoTarget.Sheets, mainWindowViewModel.ExcelInfoTarget);
            DataContext = mainWindowViewModel;
        }
    }
}
