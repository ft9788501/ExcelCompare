﻿using ExcelCompare.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCompare.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    {
        private string compareResult = string.Empty;

        public ExcelInfoViewModel ExcelInfoOrigin { get; }
        public ExcelInfoViewModel ExcelInfoTarget { get; }
        public LoadingViewModel Loading { get; }

        public string CompareResult
        {
            get => compareResult;
            set
            {
                compareResult = value;
                OnPropertyChanged();
            }
        }

        public MainWindowViewModel()
        {
            ExcelInfoOrigin = new ExcelInfoViewModel(this);
            ExcelInfoTarget = new ExcelInfoViewModel(this);
            Loading = new LoadingViewModel();
        }

        public void RefreshCompareResult()
        {
            if (ExcelInfoOrigin.ExcelInfo == null || ExcelInfoTarget.ExcelInfo == null)
            {
                CompareResult = string.Empty;
                return;
            }
            CompareResult = "数据分析中...";
            Loading.RunTaskWithErrorMsgReturn((e) =>
            {
                var originRows = ExcelInfoOrigin.ExcelInfo.GetExcelRows(ExcelInfoOrigin.SheetSelectIndex).ToArray();
                var targetRows = ExcelInfoTarget.ExcelInfo.GetExcelRows(ExcelInfoTarget.SheetSelectIndex).ToArray();
                var lacks = originRows.Where(originRow => !targetRows.Any(targetRow =>
                originRow.WaybillNumber == targetRow.WaybillNumber &&
                originRow.BoxNumber == targetRow.BoxNumber &&
                originRow.Amount == targetRow.Amount)).ToArray();

                var extras = targetRows.Where(targetRow => !originRows.Any(originRow =>
                originRow.WaybillNumber == targetRow.WaybillNumber &&
                originRow.BoxNumber == targetRow.BoxNumber &&
                originRow.Amount == targetRow.Amount)).ToArray();

                List<ValueTuple<ExcelRow, ExcelRow>> errorExcelRows = new();
                for (int i = 0; i < lacks.Length; i++)
                {
                    if (e.CancellationTokenSource.IsCancellationRequested)
                    {
                        CompareResult = "已取消";
                        return;
                    }
                    var lack = lacks[i];
                    var similarity = extras.Where(l => !errorExcelRows.Any(e => e.Item2.Index == l.Index))
                        .Select(e => new
                        {
                            ExcelRow = e,
                            Similarity = lack.Compare(e)
                        })
                        .Where(x => x.Similarity > 0.8)
                        .OrderByDescending(x => x.Similarity)
                        .FirstOrDefault();
                    if (similarity != null)
                    {
                        errorExcelRows.Add(ValueTuple.Create(lack, similarity.ExcelRow));
                    }
                    e.Progress = (double)i / lacks.Length;
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
                if (string.IsNullOrEmpty(CompareResult))
                {
                    CompareResult = "完全一致";
                }
            }).ContinueWith((errorMsg) =>
            {
                if (errorMsg.Result != null)
                {
                    CompareResult = errorMsg.Result;
                }
            });
        }
    }
}
