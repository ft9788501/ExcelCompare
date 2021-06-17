using ExcelCompare.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Forms;

namespace ExcelCompare.ViewModels
{
    public class ExcelInfoViewModel : ViewModelBase
    {
        private readonly MainWindowViewModel owner;
        private readonly ObservableCollection<string> sheets = new();
        private int sheetSelectIndex = -1;

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
        public ExcelInfoModel ExcelInfo { get; private set; }

        public ExcelInfoViewModel(MainWindowViewModel owner)
        {
            BindingOperations.EnableCollectionSynchronization(Sheets, this);
            this.owner = owner;
        }

        public async void LoadExcel()
        {
            OpenFileDialog openFileDialog = new() { Filter = "Excel(*.xlsx;*.xls)|*.xlsx;*.xls" };
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                owner.Loading = true;
                await Task.Run(()=> 
                {
                    ExcelInfo?.Dispose();
                    Sheets.Clear();
                    sheetSelectIndex = -1;
                    OnPropertyChanged(nameof(sheetSelectIndex));
                    var excelPath = openFileDialog.FileName;
                    ExcelInfo = ExcelInfoModel.LoadFromPath(excelPath);
                    if (ExcelInfo == null)
                    {
                        return;
                    }
                    foreach (var sheet in ExcelInfo.Sheets)
                    {
                        sheets.Add(sheet);
                    }
                    OnPropertyChanged(nameof(Available));
                });
                owner.Loading = false;
                if (Sheets.Count > 0)
                {
                    SheetSelectIndex = 0;
                }
            }
        }
    }
}
