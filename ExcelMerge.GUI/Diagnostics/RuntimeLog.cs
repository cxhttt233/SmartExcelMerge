using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Threading;

namespace ExcelMerge.GUI.Diagnostics
{
    /// <summary>
    /// Lightweight on-disk troubleshooting logs; not an uploader and not an analytics SDK.
    /// The timer runs independently of WPF's UI thread, so a UI freeze can still be recorded.
    /// </summary>
    public static class RuntimeLog
    {
        private static readonly object Sync = new object();
        private static string _directory;
        private static string _file;
        private static Timer _watchdog;
        private static Dispatcher _uiDispatcher;
        private static DateTime _lastUiAckUtc = DateTime.UtcNow;
        private static DateTime _lastWatchdogLogUtc = DateTime.MinValue;
        private static DateTime _operationStartUtc;
        private static DateTime _stageStartUtc;
        private static string _operation = "";
        private static string _stage = "";
        private static int _probePending;
        private static int _active;
        private static int _started;

        public static string DirectoryPath
        {
            get
            {
                EnsureFile();
                return _directory;
            }
        }

        public static string CurrentLogFile
        {
            get
            {
                EnsureFile();
                return _file;
            }
        }

        private static void EnsureFile()
        {
            if (_file != null)
                return;

            lock (Sync)
            {
                if (_file != null)
                    return;

                try
                {
                    _directory = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "SmartMerge", "logs");
                    Directory.CreateDirectory(_directory);
                }
                catch
                {
                    _directory = Path.Combine(Path.GetTempPath(), "SmartMerge-logs");
                    try { Directory.CreateDirectory(_directory); } catch { }
                }

                _file = Path.Combine(_directory, "smartmerge-" + DateTime.Now.ToString("yyyyMMdd") + ".log");
            }
        }

        public static void Write(string category, string detail)
        {
            try
            {
                EnsureFile();
                var line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff zzz") +
                    " [pid:" + Process.GetCurrentProcess().Id + "] [" + category + "] " +
                    (detail ?? "") + Environment.NewLine;
                lock (Sync)
                {
                    // Open-write-close makes each record durable, even if the user later force-terminates the process.
                    File.AppendAllText(_file, line, Encoding.UTF8);
                }
            }
            catch
            {
                // Diagnostics should never crash the user's spreadsheet application.
            }
        }

        public static void Start(Dispatcher dispatcher)
        {
            if (Interlocked.Exchange(ref _started, 1) != 0)
                return;
            _uiDispatcher = dispatcher;
            _lastUiAckUtc = DateTime.UtcNow;
            try
            {
                var p = Process.GetCurrentProcess();
                Write("START", "exe=" + p.MainModule.FileName +
                    " version=" + p.MainModule.FileVersionInfo.FileVersion +
                    " os=" + Environment.OSVersion.VersionString +
                    " 64bit=" + Environment.Is64BitProcess +
                    " log=" + CurrentLogFile);
            }
            catch (Exception ex)
            {
                Write("START", "Process metadata unavailable: " + ex.Message);
            }

            // Timer callbacks are on the thread pool, not the dispatcher.
            _watchdog = new Timer(WatchdogTick, null, 5000, 5000);
        }

        public static void BeginOperation(string label, string detail)
        {
            lock (Sync)
            {
                _operation = label ?? "compare";
                _stage = "begin";
                _operationStartUtc = DateTime.UtcNow;
                _stageStartUtc = _operationStartUtc;
                _active = 1;
            }
            Write("OP_BEGIN", _operation + " " + detail);
        }

        public static void Stage(string stage)
        {
            string current;
            DateTime started;
            lock (Sync)
            {
                current = _stage;
                started = _stageStartUtc;
                _stage = stage ?? "";
                _stageStartUtc = DateTime.UtcNow;
            }
            if (!string.Equals(current, stage, StringComparison.Ordinal))
                Write("STAGE", (stage ?? "") +
                    " | previous=" + current +
                    " previous_ms=" + (long)(DateTime.UtcNow - started).TotalMilliseconds);
        }

        public static void EndOperation(string outcome)
        {
            string name, stage;
            DateTime started;
            lock (Sync)
            {
                name = _operation;
                stage = _stage;
                started = _operationStartUtc;
                _active = 0;
                _operation = "";
                _stage = "";
            }
            Write("OP_END", "operation=" + name + " outcome=" + outcome +
                " last_stage=" + stage + " elapsed_ms=" +
                (long)(DateTime.UtcNow - started).TotalMilliseconds);
        }

        public static void Exception(string context, Exception error)
        {
            Write("EXCEPTION", context + Environment.NewLine + (error == null ? "(null)" : error.ToString()));
        }

        private static void WatchdogTick(object state)
        {
            try
            {
                var now = DateTime.UtcNow;
                if (_uiDispatcher != null && !_uiDispatcher.HasShutdownStarted &&
                    Interlocked.CompareExchange(ref _probePending, 1, 0) == 0)
                {
                    try
                    {
                        _uiDispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() =>
                        {
                            _lastUiAckUtc = DateTime.UtcNow;
                            Interlocked.Exchange(ref _probePending, 0);
                        }));
                    }
                    catch
                    {
                        Interlocked.Exchange(ref _probePending, 0);
                    }
                }

                var lag = (now - _lastUiAckUtc).TotalSeconds;
                var running = Interlocked.CompareExchange(ref _active, 0, 0) == 1;
                double elapsed, stageElapsed;
                string operation, stage;
                lock (Sync)
                {
                    elapsed = (now - _operationStartUtc).TotalSeconds;
                    stageElapsed = (now - _stageStartUtc).TotalSeconds;
                    operation = _operation;
                    stage = _stage;
                }

                // Even an operation with a responsive UI can have a stalled worker.
                if (lag < 12 && !(running && stageElapsed >= 12))
                    return;
                if ((now - _lastWatchdogLogUtc).TotalSeconds < 10)
                    return;
                _lastWatchdogLogUtc = now;
                using (var p = Process.GetCurrentProcess())
                {
                    Write(lag >= 12 ? "UI_HANG_SUSPECT" : "LONG_STEP",
                        "ui_lag_s=" + lag.ToString("F1") +
                        " operation=" + operation + " stage=" + stage +
                        " operation_s=" + elapsed.ToString("F1") +
                        " stage_s=" + stageElapsed.ToString("F1") +
                        " cpu_s=" + p.TotalProcessorTime.TotalSeconds.ToString("F1") +
                        " working_mb=" + (p.WorkingSet64 / 1048576) +
                        " managed_mb=" + (GC.GetTotalMemory(false) / 1048576) +
                        " threads=" + p.Threads.Count);
                }
            }
            catch (Exception ex)
            {
                Write("WATCHDOG_ERROR", ex.Message);
            }
        }
    }
}
