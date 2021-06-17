using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelCompare.Models
{
    public class ExcelInfoModel:IDisposable
    {
        public static ExcelInfoModel LoadFromPath(string excelPath)
        {
            try
            {
                var excelExtension = Path.GetExtension(excelPath);
                using var fileStream = File.OpenRead(excelPath);
                IWorkbook workbook = excelExtension switch
                {
                    ".xlsx" => new XSSFWorkbook(fileStream),
                    ".xls" => new HSSFWorkbook(fileStream),
                    _ => null
                };
                if (workbook == null)
                {
                    return null;
                }
                else
                {
                    return new ExcelInfoModel(workbook);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        private readonly List<string> sheets = new();
        private readonly IWorkbook workbook;

        public IEnumerable<string> Sheets => sheets;

        private ExcelInfoModel(IWorkbook workbook)
        {
            this.workbook = workbook;
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                sheets.Add(workbook.GetSheetName(i));
            }
        }

        public IEnumerable<ExcelRow> GetExcelRows(int sheetIndex)
        {
            var sheet = workbook.GetSheetAt(sheetIndex);
            var excelRows = Enumerable.Range(1, sheet.LastRowNum)
                .Select(r => sheet.GetRow(r))
                .Select((sr, i) => new ExcelRow()
                {
                    Index = i + 2,
                    WaybillNumber = sr.GetCell(0).ToString(),
                    BoxNumber = sr.GetCell(1).ToString(),
                    Amount = sr.GetCell(2).ToString()
                });
            return excelRows;
        }

        public void Dispose()
        {
            workbook.Close();
        }

        public class ExcelRow
        {
            public int Index { get; set; }
            public string WaybillNumber { get; set; }
            public string BoxNumber { get; set; }
            public string Amount { get; set; }

            public override string ToString()
            {
                return $"行号:{Index},运单号:{WaybillNumber},箱号:{BoxNumber},实际总未收本位币金额:{Amount}";
            }
        }
    }
}
