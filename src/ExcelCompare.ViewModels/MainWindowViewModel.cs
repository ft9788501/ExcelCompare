using ExcelCompare.ViewModels.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ExcelCompare.Models.ExcelInfoModel;

namespace ExcelCompare.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    {
        private string compareResult = string.Empty;
        public bool loading = false;

        public ExcelInfoViewModel ExcelInfoOrigin { get; }
        public ExcelInfoViewModel ExcelInfoTarget { get; }

        public string CompareResult
        {
            get => compareResult;
            set
            {
                compareResult = value;
                OnPropertyChanged();
            }
        }
        public bool Loading
        {
            get => loading;
            set
            {
                loading = value;
                OnPropertyChanged();
            }
        }

        public MainWindowViewModel()
        {
            ExcelInfoOrigin = new ExcelInfoViewModel(this);
            ExcelInfoTarget = new ExcelInfoViewModel(this);
        }

        public void RefreshCompareResult()
        {
            //Loading = true;
            if (ExcelInfoOrigin.ExcelInfo == null || ExcelInfoTarget.ExcelInfo == null)
            {
                CompareResult = string.Empty;
                return;
            }
            var originRows = ExcelInfoOrigin.ExcelInfo.GetExcelRows(ExcelInfoOrigin.SheetSelectIndex);
            var targetRows = ExcelInfoTarget.ExcelInfo.GetExcelRows(ExcelInfoTarget.SheetSelectIndex);

            var lacks = originRows.Where(originRow => !targetRows.Any(targetRow =>
            originRow.WaybillNumber == targetRow.WaybillNumber &&
            originRow.BoxNumber == targetRow.BoxNumber &&
            originRow.Amount == targetRow.Amount));

            var extras = targetRows.Where(targetRow => !originRows.Any(originRow =>
            originRow.WaybillNumber == targetRow.WaybillNumber &&
            originRow.BoxNumber == targetRow.BoxNumber &&
            originRow.Amount == targetRow.Amount));

            List<ValueTuple<ExcelRow, ExcelRow>> errorExcelRows = new();
            foreach (var lack in lacks)
            {
                var similarity = extras.Where(l => !errorExcelRows.Any(e => e.Item2.Index == l.Index))
                    .Select(e => new
                    {
                        ExcelRow = e,
                        Similarity = SimilarityHelper.CompareStrings(lack.Amount, e.Amount) / 3 + SimilarityHelper.CompareStrings(lack.BoxNumber, e.BoxNumber) / 3 + SimilarityHelper.CompareStrings(lack.WaybillNumber, e.WaybillNumber) / 3
                    }).OrderByDescending(x => x.Similarity)
                    .FirstOrDefault(x => x.Similarity > 0.8);
                if (similarity != null)
                {
                    errorExcelRows.Add(ValueTuple.Create(lack, similarity.ExcelRow));
                }
            }

            var duplicates = targetRows.Where(targetRow => targetRows.Count(originRow =>
            originRow.WaybillNumber == targetRow.WaybillNumber &&
            originRow.BoxNumber == targetRow.BoxNumber &&
            originRow.Amount == targetRow.Amount) > 1);

            StringBuilder compareResult = new();
            if (lacks.Where(l => !errorExcelRows.Any(e => e.Item1.Index == l.Index)) is var lacksFiltered && lacksFiltered.Any())
            {
                compareResult.AppendLine($"缺少:\r\n{string.Join("\r\n", lacksFiltered)}");
                compareResult.AppendLine($"---------------------------------------------------------------------------------------------------------");
            }
            if (extras.Where(l => !errorExcelRows.Any(e => e.Item2.Index == l.Index)) is var extrasFiltered && extrasFiltered.Any())
            {
                compareResult.AppendLine($"多余:\r\n{string.Join("\r\n", extrasFiltered)}");
                compareResult.AppendLine($"---------------------------------------------------------------------------------------------------------");
            }
            if (errorExcelRows.Any())
            {
                compareResult.AppendLine($"错误:\r\n{string.Join("\r\n", errorExcelRows.Select(r => r.Item1.CompareResult(r.Item2)))}");
                compareResult.AppendLine($"---------------------------------------------------------------------------------------------------------");
            }
            if (duplicates.Any())
            {
                compareResult.AppendLine($"重复:\r\n{string.Join("\r\n", duplicates)}");
            }
            CompareResult = compareResult.ToString();
            //Loading = false;
        }
    }
}
