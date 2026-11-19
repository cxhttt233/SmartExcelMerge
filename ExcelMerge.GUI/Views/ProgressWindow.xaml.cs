using System;
using System.ComponentModel;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows;
using ExcelMerge.GUI.Diagnostics;

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

        private void CancelProgress_Click(object sender, RoutedEventArgs e)
        {
            RuntimeLog.Write("CANCEL_BUTTON", "User clicked cancel comparison");
            Close();
        }

        public static bool DoWorkWithModal(Action<IProgress<string>, CancellationToken> work)
        {
            if (work == null) throw new ArgumentNullException("work");

            RuntimeLog.Write("PROGRESS_BEGIN", "Background progress dialog created");
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
                RuntimeLog.Write("CANCEL_REQUEST", "Progress window closed, requesting background cancellation");
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
                    token.ThrowIfCancellationRequested();
                    work(progress, token);
                    token.ThrowIfCancellationRequested();
                };
                worker.RunWorkerCompleted += (source, workerArgs) =>
                {
                    // BackgroundWorker marshals exceptions to Error on the UI thread.
                    workerError = workerArgs.Error;
                    if (workerError != null)
                        RuntimeLog.Exception("PROGRESS_WORKER", workerError);
                    else if (workerArgs.Cancelled)
                        RuntimeLog.Write("PROGRESS_WORKER_CANCELLED", "Cancellation observed by worker");
                    else
                        RuntimeLog.Write("PROGRESS_WORKER_DONE", "Background worker completed");
                    completed = true;
                    if (!abandoned && window.IsVisible)
                        window.Close();
                };
                worker.RunWorkerAsync();
            };

            window.ShowDialog();
            if (abandoned || token.IsCancellationRequested)
            {
                RuntimeLog.Write("PROGRESS_CANCEL_RETURN", "No result applied; UI returned while background may still be unwinding");
                cancellation.Dispose();
                return false;
            }

            cancellation.Dispose();
            if (workerError != null)
                ExceptionDispatchInfo.Capture(workerError).Throw();
            RuntimeLog.Write("PROGRESS_END", "completed=" + completed);
            return completed;
        }
    }
}
