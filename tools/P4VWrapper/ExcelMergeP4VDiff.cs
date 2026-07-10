using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelMergeP4VDiff
{
    internal static class Program
    {
        private const int MaxCacheAgeDays = 3;
        private const int MaxCacheDirectoryCount = 100;
        private const int KeepCacheDirectoryCount = 50;
        private const long MaxLogBytes = 1024 * 1024;

        private sealed class StablePaths
        {
            public StablePaths(string left, string right)
            {
                Left = left;
                Right = right;
            }

            public string Left { get; private set; }
            public string Right { get; private set; }
        }

        [STAThread]
        private static int Main(string[] args)
        {
            string logPath = GetLogPath();

            try
            {
                CleanupOldCache(logPath);

                List<string> normalizedArgs = NormalizeArguments(args);
                List<string> paths = ExtractPaths(normalizedArgs);

                LogArguments(logPath, args, normalizedArgs, paths);

                if (paths.Count < 2)
                {
                    Fail("P4V did not pass two file paths to ExcelMerge.", logPath);
                    return 2;
                }

                StablePaths stablePaths = CopyToStableCache(paths[0], paths[1], logPath);
                CleanupOldCache(logPath);

                string excelMerge = FindExcelMerge();
                if (excelMerge == null)
                {
                    Fail("ExcelMerge.GUI.exe was not found.", logPath);
                    return 3;
                }

                ProcessStartInfo startInfo = new ProcessStartInfo(excelMerge);
                startInfo.UseShellExecute = false;
                startInfo.Arguments = "diff -s " + Quote(stablePaths.Left) + " -d " + Quote(stablePaths.Right) + " -k" + CreateEditArguments(paths[1]);
                LogLaunch(logPath, excelMerge, startInfo.Arguments);
                Process.Start(startInfo);
                return 0;
            }
            catch (Exception ex)
            {
                File.AppendAllText(logPath, DateTime.Now.ToString("O") + " ERROR " + ex + Environment.NewLine, Encoding.UTF8);
                Fail(ex.Message, logPath);
                return 1;
            }
        }

        private static List<string> NormalizeArguments(string[] rawArgs)
        {
            List<string> normalized = new List<string>();
            foreach (string raw in rawArgs)
            {
                string value = CleanToken(raw);
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                if (string.Equals(value, "--open", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(value, "-open", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(value, "/open", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                normalized.Add(value);
            }

            return normalized;
        }

        private static List<string> ExtractPaths(List<string> args)
        {
            List<string> paths = ExtractOptionPaths(args);
            if (paths.Count >= 2)
            {
                return paths;
            }

            List<string> positional = new List<string>();
            foreach (string arg in args)
            {
                if (string.Equals(arg, "diff", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (arg.StartsWith("-", StringComparison.Ordinal))
                {
                    continue;
                }

                positional.Add(arg);
            }

            paths = ExtractExistingPaths(positional);
            if (paths.Count >= 2)
            {
                return paths;
            }

            List<string> fallback = new List<string>();
            for (int i = 0; i < positional.Count && i < 2; i++)
            {
                fallback.Add(positional[i]);
            }

            return fallback;
        }

        private static List<string> ExtractOptionPaths(List<string> args)
        {
            List<string> result = new List<string>();
            for (int i = 0; i < args.Count - 1; i++)
            {
                if (string.Equals(args[i], "-s", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(args[i], "--src-path", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(args[i], "-d", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(args[i], "--dst-path", StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(args[i + 1]);
                    i++;
                }
            }

            return result;
        }

        private static List<string> ExtractExistingPaths(List<string> args)
        {
            List<string> result = new List<string>();
            for (int i = 0; i < args.Count && result.Count < 2; i++)
            {
                string current = args[i];
                if (File.Exists(current))
                {
                    result.Add(current);
                    continue;
                }

                string combined = current;
                for (int j = i + 1; j < args.Count; j++)
                {
                    combined += " " + args[j];
                    if (!File.Exists(combined))
                    {
                        continue;
                    }

                    result.Add(combined);
                    i = j;
                    break;
                }
            }

            return result;
        }

        private static string CleanToken(string value)
        {
            value = (value ?? string.Empty).Trim();
            while (value.Length >= 2 &&
                   ((value[0] == '"' && value[value.Length - 1] == '"') ||
                    (value[0] == '\'' && value[value.Length - 1] == '\'')))
            {
                value = value.Substring(1, value.Length - 2).Trim();
            }

            return value;
        }

        private static string FindExcelMerge()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string[] candidates = new string[]
            {
                Path.GetFullPath(Path.Combine(baseDirectory, "..", "app", "ExcelMerge.GUI.exe")),
                Path.Combine(baseDirectory, "app", "ExcelMerge.GUI.exe"),
                Path.Combine(baseDirectory, "ExcelMerge.GUI.exe")
            };

            foreach (string candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }

            return null;
        }

        private static string CreateEditArguments(string originalRightPath)
        {
            if (!IsEditableOriginalPath(originalRightPath))
            {
                return string.Empty;
            }

            return " --dst-edit-path " + Quote(originalRightPath) + " --editable-side dst";
        }

        private static bool IsEditableOriginalPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return false;
            }

            string extension = Path.GetExtension(path);
            if (!string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string fullPath = Path.GetFullPath(path);
            string tempPath = Path.GetFullPath(Path.GetTempPath());
            return !fullPath.StartsWith(tempPath, StringComparison.OrdinalIgnoreCase);
        }

        private static StablePaths CopyToStableCache(string left, string right, string logPath)
        {
            if (!File.Exists(left))
            {
                throw new FileNotFoundException("Left diff file does not exist.", left);
            }

            if (!File.Exists(right))
            {
                throw new FileNotFoundException("Right diff file does not exist.", right);
            }

            string cacheRoot = GetCacheRoot();
            string cacheDir = Path.Combine(cacheRoot, DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + "_" + Process.GetCurrentProcess().Id.ToString());
            Directory.CreateDirectory(cacheDir);

            string leftPath = Path.Combine(cacheDir, "left" + GetSafeExtension(left));
            string rightPath = Path.Combine(cacheDir, "right" + GetSafeExtension(right));

            CopyShared(left, leftPath);
            CopyShared(right, rightPath);

            LogCache(logPath, left, right, leftPath, rightPath);
            return new StablePaths(leftPath, rightPath);
        }

        private static string GetSafeExtension(string path)
        {
            string extension = Path.GetExtension(path);
            if (string.IsNullOrWhiteSpace(extension) || extension.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                return ".xlsx";
            }

            return extension;
        }

        private static void CopyShared(string source, string destination)
        {
            using (FileStream input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (FileStream output = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.Read))
            {
                input.CopyTo(output);
            }
        }

        private static string GetLocalStateRoot()
        {
            string root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (string.IsNullOrWhiteSpace(root))
            {
                root = Path.GetTempPath();
            }

            return Path.Combine(root, "ExcelSmartDiff", "p4v-diff");
        }

        private static string GetCacheRoot()
        {
            string cacheRoot = Path.Combine(GetLocalStateRoot(), "cache");
            Directory.CreateDirectory(cacheRoot);
            return cacheRoot;
        }

        private static string GetLogPath()
        {
            string stateRoot = GetLocalStateRoot();
            Directory.CreateDirectory(stateRoot);
            string logPath = Path.Combine(stateRoot, "ExcelMergeP4VDiff.log");

            try
            {
                if (File.Exists(logPath) && new FileInfo(logPath).Length > MaxLogBytes)
                {
                    File.Delete(logPath);
                }
            }
            catch
            {
                // Logging must not block diff.
            }

            return logPath;
        }

        private static void CleanupOldCache(string logPath)
        {
            string cacheRoot = GetCacheRoot();
            DirectoryInfo[] directories = new DirectoryInfo(cacheRoot).GetDirectories();

            foreach (DirectoryInfo directory in directories)
            {
                if (directory.LastWriteTimeUtc < DateTime.UtcNow.AddDays(-MaxCacheAgeDays))
                {
                    DeleteCacheDirectory(directory, logPath);
                }
            }

            directories = new DirectoryInfo(cacheRoot)
                .GetDirectories()
                .OrderByDescending(d => d.LastWriteTimeUtc)
                .ToArray();

            if (directories.Length <= MaxCacheDirectoryCount)
            {
                return;
            }

            foreach (DirectoryInfo directory in directories.Skip(KeepCacheDirectoryCount))
            {
                DeleteCacheDirectory(directory, logPath);
            }
        }

        private static void DeleteCacheDirectory(DirectoryInfo directory, string logPath)
        {
            try
            {
                directory.Delete(true);
            }
            catch (Exception ex)
            {
                File.AppendAllText(logPath, DateTime.Now.ToString("O") + " CACHE CLEANUP " + directory.FullName + " " + ex.Message + Environment.NewLine, Encoding.UTF8);
            }
        }

        private static void LogArguments(string logPath, string[] rawArgs, List<string> normalizedArgs, List<string> paths)
        {
            StringBuilder message = new StringBuilder();
            message.AppendLine("[" + DateTime.Now.ToString("O") + "]");
            message.AppendLine("Local state root:");
            message.AppendLine(GetLocalStateRoot());
            message.AppendLine("Raw Environment.CommandLine:");
            message.AppendLine(Environment.CommandLine);
            message.AppendLine("Raw args:");
            for (int i = 0; i < rawArgs.Length; i++)
            {
                message.AppendLine("  [" + i + "] " + rawArgs[i]);
            }

            message.AppendLine("Normalized args:");
            for (int i = 0; i < normalizedArgs.Count; i++)
            {
                message.AppendLine("  [" + i + "] " + normalizedArgs[i]);
            }

            message.AppendLine("Selected paths:");
            for (int i = 0; i < paths.Count; i++)
            {
                message.AppendLine("  [" + i + "] " + paths[i]);
            }

            message.AppendLine();
            File.AppendAllText(logPath, message.ToString(), Encoding.UTF8);
        }

        private static void LogCache(string logPath, string left, string right, string leftPath, string rightPath)
        {
            StringBuilder message = new StringBuilder();
            message.AppendLine("Stable cache:");
            message.AppendLine("  left source exists: " + File.Exists(left));
            message.AppendLine("  right source exists: " + File.Exists(right));
            message.AppendLine("  left cache: " + leftPath);
            message.AppendLine("  left cache size: " + new FileInfo(leftPath).Length.ToString());
            message.AppendLine("  right cache: " + rightPath);
            message.AppendLine("  right cache size: " + new FileInfo(rightPath).Length.ToString());
            message.AppendLine();
            File.AppendAllText(logPath, message.ToString(), Encoding.UTF8);
        }

        private static void LogLaunch(string logPath, string excelMerge, string arguments)
        {
            StringBuilder message = new StringBuilder();
            message.AppendLine("Launch:");
            message.AppendLine("  exe: " + excelMerge);
            message.AppendLine("  args: " + arguments);
            message.AppendLine();
            File.AppendAllText(logPath, message.ToString(), Encoding.UTF8);
        }

        private static string Quote(string value)
        {
            return "\"" + value.Replace("\"", "\\\"") + "\"";
        }

        private static void Fail(string message, string logPath)
        {
            MessageBox(IntPtr.Zero, message + "\n\nLog:\n" + logPath, "ExcelMerge P4V Diff", 0x10);
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBox(IntPtr hWnd, string text, string caption, uint type);
    }
}
