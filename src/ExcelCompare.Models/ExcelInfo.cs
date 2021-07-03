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
    public class ExcelInfo : IDisposable
    {
        public static ExcelInfo LoadFromPath(string excelPath, out string errorMsg)
        {
            try
            {
                var excelExtension = Path.GetExtension(excelPath);
                using var fileStream = File.OpenRead(excelPath);
                IWorkbook workbook = excelExtension switch
                {
                    ".xlsx" => new XSSFWorkbook(fileStream) { IsHidden = false },
                    ".xls" => new HSSFWorkbook(fileStream) { IsHidden = false },
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
                    return new ExcelInfo(excelPath, workbook);
                }
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message;
                return null;
            }
        }

        private readonly List<string> sheets = new();
        private readonly string excelPath;
        private readonly IWorkbook workbook;

        public IEnumerable<string> Sheets => sheets;

        private ExcelInfo(string excelPath, IWorkbook workbook)
        {
            this.excelPath = excelPath;
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
                .Select((sr, i) => new ExcelRow(excelPath, i + 2, sr));
            return excelRows;
        }

        public void Save()
        {
            //var newPath = Path.Combine(Path.GetDirectoryName(excelPath), $"a{Path.GetExtension(excelPath)}");
            //using var fileStream = new FileStream(newPath, FileMode.OpenOrCreate);
            //workbook.Write(fileStream);
        }

        public void Dispose()
        {
            workbook.Close();
        }
    }
}
