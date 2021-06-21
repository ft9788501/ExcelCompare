using ExcelCompare.Models.Utils;
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
    public class ExcelInfoModel : IDisposable
    {
        public static ExcelInfoModel LoadFromPath(string excelPath, out string errorMsg)
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
                    errorMsg = null;
                    return null;
                }
                else
                {
                    errorMsg = null;
                    return new ExcelInfoModel(workbook);
                }
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message;
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

            public double Compare(ExcelRow excelRow)
            {
                if ((Amount == excelRow.Amount ? 1 : 0) + (BoxNumber == excelRow.BoxNumber ? 1 : 0) + (WaybillNumber == excelRow.WaybillNumber ? 1 : 0) >= 2)
                {
                    return 0.99D;
                }
                var similarity = SimilarityHelper.CompareStrings(Amount, excelRow.Amount) / 3 + SimilarityHelper.CompareStrings(BoxNumber, excelRow.BoxNumber) / 3 + SimilarityHelper.CompareStrings(WaybillNumber, excelRow.WaybillNumber) / 3;
                return similarity;
            }

            public string CompareResult(ExcelRow excelRow)
            {
                StringBuilder stringBuilder = new StringBuilder();
                if (WaybillNumber != excelRow.WaybillNumber ||
                    BoxNumber != excelRow.BoxNumber ||
                    Amount != excelRow.Amount)
                {
                    var similarity = Compare(excelRow);
                    stringBuilder.AppendLine($"行号:{ Index}->行号:{ excelRow.Index} (相似度:{similarity})");
                }
                if (WaybillNumber != excelRow.WaybillNumber)
                {
                    stringBuilder.AppendLine($"运单号:{ WaybillNumber}->{ excelRow.WaybillNumber}");
                }
                if (BoxNumber != excelRow.BoxNumber)
                {
                    stringBuilder.AppendLine($"箱号:{ BoxNumber}->{ excelRow.BoxNumber}");
                }
                if (Amount != excelRow.Amount)
                {
                    stringBuilder.AppendLine($"实际总未收本位币金额:{ Amount}->{ excelRow.Amount}");
                }
                return stringBuilder.ToString();
            }
        }
    }
}
