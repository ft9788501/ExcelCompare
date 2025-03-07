﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelCompare.ViewModels
{
    public class LoadingViewModel : ViewModelBase
    {
        private bool visible = false;
        private double progress = 0;
        private CancellationTokenSource cancellationTokenSource;

        public bool Visible
        {
            get => visible;
            private set
            {
                visible = value;
                OnPropertyChanged();
            }
        }

        public double Progress
        {
            get => progress;
            private set
            {
                progress = value;
                OnPropertyChanged();
            }
        }

        public Task<string> RunTaskWithErrorMsgReturn(Action<LoadingArgs> task)
        {
            Progress = 0;
            Visible = true;
            cancellationTokenSource = new CancellationTokenSource();
            LoadingArgs loadingArgs = new LoadingArgs(this, cancellationTokenSource);
            return Task.Run(() =>
            {
                try
                {
                    task.Invoke(loadingArgs);
                    return null;
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
                finally
                {
                    Visible = false;
                }
            });
        }

        public void Cancel()
        {
            cancellationTokenSource.Cancel();
        }

        public class LoadingArgs
        {
            private LoadingViewModel loadingViewModel;
            private CancellationTokenSource cancellationTokenSource;

            public double Progress
            {
                get => loadingViewModel.Progress;
                set
                {
                    loadingViewModel.Progress = value;
                }
            }

            public CancellationTokenSource CancellationTokenSource
            {
                get => cancellationTokenSource;
            }

            public LoadingArgs(LoadingViewModel loadingViewModel, CancellationTokenSource cancellationTokenSource)
            {
                this.loadingViewModel = loadingViewModel;
                this.cancellationTokenSource = cancellationTokenSource;
            }
        }
    }
}
