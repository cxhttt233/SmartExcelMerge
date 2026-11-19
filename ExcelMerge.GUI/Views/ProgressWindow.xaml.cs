using System;
using System.ComponentModel;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows;

namespace ExcelMerge.GUI.Views
{
    /// <summary>
    /// A cancellable progress dialog. Closing the dialog never applies an incomplete result.
    /// The worker receives a cancellation token and must check it at safe boundaries.
    /// </summary>
    public partial class ProgressWindow : Window
    {
        public ProgressWindow()
        {
            InitializeComponent();
        }

        public static bool DoWorkWithModal(Action<IProgress<string>, CancellationToken> work)
        {
            if (work == null) throw new ArgumentNullException("work");

            var window = new ProgressWindow();
            var cancellation = new CancellationTokenSource();
            var token = cancellation.Token;
            var worker = new BackgroundWorker();
            Exception workerError = null;
            bool abandoned = false;
            bool completed = false;

            window.Closing += (sender, args) =>
            {
                if (completed || !worker.IsBusy) return;
                // Do not block the user's UI while a large legacy operation is unwinding.
                // Never publish its eventual result after this dialog has been dismissed.
                abandoned = true;
                cancellation.Cancel();
            };

            window.Loaded += (sender, args) =>
            {
                var progress = new Progress<string>(message =>
                {
                    if (!abandoned)
                        window.Message.Content = message;
                });

                worker.DoWork += (source, workerArgs) =>
                {
                    try
                    {
                        token.ThrowIfCancellationRequested();
                        work(progress, token);
                        token.ThrowIfCancellationRequested();
                    }
                    catch (OperationCanceledException)
                    {
                        // User cancellation, not an application error.
                    }
                    catch (Exception exception)
                    {
                        workerError = exception;
                    }
                };
                worker.RunWorkerCompleted += (source, workerArgs) =>
                {
                    completed = true;
                    if (!abandoned && window.IsVisible)
                        window.Close();
                };
                worker.RunWorkerAsync();
            };

            window.ShowDialog();
            if (abandoned || token.IsCancellationRequested)
            {
                cancellation.Dispose();
                return false;
            }

            cancellation.Dispose();
            if (workerError != null)
                ExceptionDispatchInfo.Capture(workerError).Throw();
            return completed;
        }
    }
}
