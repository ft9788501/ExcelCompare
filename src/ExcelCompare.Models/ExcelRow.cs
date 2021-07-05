using ExcelCompare.Models.Utils;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCompare.Models
{
    public class ExcelRow
    {
        private IRow row;

        public int Index { get; }
        public string WaybillNumber { get; }
        public string BoxNumber { get; }
        public string Amount { get; }

        public bool Visible
        {
            get => !row.ZeroHeight;
            set
            {
                row.ZeroHeight = !value;
            }
        }

        public ExcelRow(string excelPath, int index, IRow row)
        {
            this.row = row;
            Index = index;
            WaybillNumber = row.GetCell(0).ToString().Trim();
            BoxNumber = row.GetCell(1).ToString().Trim();

            if (double.TryParse(row.GetCell(2)?.ToString().Trim() ?? "0", out double amount))
            {
                Amount = amount.ToString();
            }
            else
            {
                throw new Exception($"{Path.GetFileName(excelPath)}的第{Index}行的金额({row.GetCell(2).ToString().Trim()})非法！");
            }
        }

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
