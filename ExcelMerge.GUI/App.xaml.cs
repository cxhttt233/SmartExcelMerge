using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using ExcelMerge.GUI.Commands;
using ExcelMerge.GUI.Settings;
using ExcelMerge.GUI.Diagnostics;

namespace ExcelMerge.GUI
{
    public partial class App : Application
    {
        public ApplicationSetting Setting { get; private set; }
        public CommandLineOption CommandLineOption { get; private set; }

        public event Action OnSettingUpdated;

        [STAThread()]
        public static void Main()
        {
            App app = new App();
            app.InitializeComponent();
            RuntimeLog.Start(app.Dispatcher);
            RuntimeLog.Write("BOOT", "Loading settings");
            app.Setting = ApplicationSetting.Load();
            app.Setting.EnsureCulture();
            app.Setting.Culture = "zh-CN";
            app.UpdateResourceCulture();

            if (app.Setting.Ensure())
                app.Setting.Save();

            app.Run();
        }

        public static App Instance
        {
            get { return (App)Current; }
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            AppDomain.CurrentDomain.ProcessExit += (sender, args) =>
                RuntimeLog.Write("PROCESS_EXIT", "OS process exit event");
            DispatcherUnhandledException += (sender, args) =>
            {
                RuntimeLog.Exception("WPF DispatcherUnhandledException", args.Exception);
                // Do not pretend a failed UI operation succeeded.
            };
            System.Threading.Tasks.TaskScheduler.UnobservedTaskException += (sender, args) =>
                RuntimeLog.Exception("UnobservedTaskException", args.Exception);
            ComparisonTelemetry.StageChanged += RuntimeLog.Stage;
            RuntimeLog.Write("APP_STARTUP", "Main window startup");

            base.OnStartup(e);

            var args = Environment.GetCommandLineArgs().Skip(1).ToList();
            if (!args.Any())
                args.Add(CommandType.Diff.ToString());

            CommandLineOption = new CommandLineOption();

            var command = CreateCommand(args.ToArray());
            command.ValidateOption();
            command.Execute();
        }

        private void StoreOption()
        {
            EMEnvironmentValue.Set("SRC", CommandLineOption.SrcPath);
            EMEnvironmentValue.Set("DST", CommandLineOption.DstPath);
        }

        private ICommand CreateCommand(string[] args)
        {
            if (CommandLine.Parser.Default.ParseArguments(args, CommandLineOption))
            {
                StoreOption();
                CommandLineOption.ConvertToFullPath();
                return CommandFactory.Create(CommandLineOption);
            }

            throw new Exceptions.ExcelMergeException(true, $"参数无效。\n参数：\n{string.Join(" ", args)}");
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var exception = e.ExceptionObject as Exception;
            RuntimeLog.Exception("FATAL UnhandledException terminating=" + e.IsTerminating,
                exception ?? new Exception(Convert.ToString(e.ExceptionObject)));
            try
            {
                MessageBox.Show(
                    "SmartMerge 遇到未处理异常，诊断日志已保存。\\n\\n" +
                    (exception == null ? "未知错误" : exception.Message) +
                    "\\n\\n日志文件：\\n" + RuntimeLog.CurrentLogFile,
                    "SmartMerge · 程序错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch { }
            Environment.Exit(-1);
        }

        public void ExecuteExternalCommand()
        {
            var command = Setting.ExternalCommands.FirstOrDefault(c => c.Name == CommandLineOption.ExternalCommand);
            if (command == null)
                return;

            command.Execute(CommandLineOption.WaitExternalCommand);
        }

        public void UpdateSetting(ApplicationSetting setting)
        {
            Setting = setting.DeepClone();

            if (OnSettingUpdated == null)
                OnSettingUpdated += () => { };

            OnSettingUpdated();
        }

        public void UpdateRecentFiles(string srcPath, string dstPath)
        {
            var updated = Setting.RecentFileSets.ToList();
            var key = srcPath + "|" + dstPath;
            var index = updated.IndexOf(key);
            if (index >= 0)
            {
                updated.RemoveAt(index);
                updated.Insert(0, key);
            }
            else
            {
                updated.Insert(0, srcPath + "|" + dstPath);
            }

            while (updated.Count > 20)
            {
                updated.RemoveAt(updated.Count - 1);
            }

            Setting.RecentFileSets = new System.Collections.ObjectModel.ObservableCollection<string>(updated);
            Setting.Save();
        }

        public void UpdateResourceCulture()
        {
            if (string.IsNullOrEmpty(Setting.Culture))
                return;

            if (GUI.Properties.Resources.Culture != null)
            {
                if (GUI.Properties.Resources.Culture.Name == Setting.Culture)
                    return;

                MessageBox.Show(GUI.Properties.Resources.Message_Reboot);
            }

            GUI.Properties.Resources.Culture = new System.Globalization.CultureInfo(Setting.Culture);
        }

        public IEnumerable<string> GetRecentFiles()
        {
            return Setting.RecentFileSets.SelectMany(f => f.Split('|'));
        }

        public IEnumerable<string> GetRecentSrcFiles()
        {
            return Setting.RecentFileSets.Select(f => f.Split('|').ElementAtOrDefault(0));
        }

        public IEnumerable<string> GetRecentDstFiles()
        {
            return Setting.RecentFileSets.Select(f => f.Split('|').ElementAtOrDefault(1));
        }

        public IEnumerable<Tuple<string, string>> GetRecentFileSets()
        {
            return Setting.RecentFileSets.Select(f =>
            {
                var files = f.Split('|');
                return Tuple.Create(files.ElementAtOrDefault(0), files.ElementAtOrDefault(1));
            });
        }

        public bool KeepFileHistory
        {
            get { return CommandLineOption.KeepFileHistory; }
        }
    }
}
