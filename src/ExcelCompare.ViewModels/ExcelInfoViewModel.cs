using ExcelCompare.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelCompare.ViewModels
{
    public class ExcelInfoViewModel : ViewModelBase
    {
        private readonly MainWindowViewModel owner;
        private readonly ObservableCollection<string> sheets = new();
        private int sheetSelectIndex = -1;
        private string errorMsg;

        public int SheetSelectIndex
        {
            get => sheetSelectIndex;
            set
            {
                if (sheetSelectIndex != value)
                {
                    sheetSelectIndex = value;
                    OnPropertyChanged();
                    owner.RefreshCompareResult();
                }
            }
        }
        public ObservableCollection<string> Sheets => sheets;
        public bool Available => Sheets.Count > 0;
        public ExcelInfo ExcelInfo { get; private set; }
        public string ErrorMsg
        {
            get => errorMsg;
            private set
            {
                errorMsg = value;
                OnPropertyChanged();
            }
        }

        public ExcelInfoViewModel(MainWindowViewModel owner)
        {
            this.owner = owner;
        }

        public void LoadExcel(string excelPath)
        {
            ExcelInfo?.Dispose();
            Sheets.Clear();
            sheetSelectIndex = -1;
            OnPropertyChanged(nameof(sheetSelectIndex));
            ExcelInfo = ExcelInfo.LoadFromPath(excelPath, out string errorMsg);
            if (ExcelInfo == null)
            {
                ErrorMsg = errorMsg;
                OnPropertyChanged(nameof(Available));
                return;
            }
            foreach (var sheet in ExcelInfo.Sheets)
            {
                sheets.Add(sheet);
            }
            OnPropertyChanged(nameof(Available));
            if (Sheets.Count > 0)
            {
                SheetSelectIndex = 0;
            }
        }
    }
}
