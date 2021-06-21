using ExcelCompare.ViewModels;
using System;
using Xunit;

namespace ExcelCompare.UnitTest
{
    public class ExcelCompareTest
    {
        [Fact]
        public void Test1()
        {
            MainWindowViewModel mainWindowViewModel = new MainWindowViewModel();
            mainWindowViewModel.ExcelInfoOrigin.LoadExcel("Assets/origin.xls");
            mainWindowViewModel.ExcelInfoTarget.LoadExcel("Assets/target.xls");
        }
    }
}
