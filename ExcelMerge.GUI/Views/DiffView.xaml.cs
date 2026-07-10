using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Practices.Unity;
using FastWpfGrid;
using NetDiff;
using SKCore.Collection;
using ExcelMerge.GUI.ViewModels;
using ExcelMerge.GUI.Settings;
using ExcelMerge.GUI.Models;
using ExcelMerge.GUI.Styles;

namespace ExcelMerge.GUI.Views
{
    public partial class DiffView : UserControl
    {
        private ExcelSheetDiffConfig diffConfig = new ExcelSheetDiffConfig();
        private IUnityContainer container;
        private const string srcKey = "src";
        private const string dstKey = "dst";

        private FastGridControl copyTargetGrid;
        private bool isLoaded;
        private bool isExecutingDiff;
        private bool isSynchronizingSheetSelection;
        private bool isWindowClosingHandlerAttached;
        private string copiedRowsText = string.Empty;

        public DiffView()
        {
            InitializeComponent();
            InitializeContainer();
            InitializeEventListeners();

            App.Instance.OnSettingUpdated += OnApplicationSettingUpdated;

            SearchTextCombobox.ItemsSource = App.Instance.Setting.SearchHistory.ToList();

            // In order to enable Ctrl + F immediately after startup.
            ToolExpander.IsExpanded = true;
        }

        private DiffViewModel GetViewModel()
        {
            return DataContext as DiffViewModel;
        }

        private void InitializeContainer()
        {
            container = new UnityContainer();
            container
                .RegisterInstance(srcKey, SrcDataGrid)
                .RegisterInstance(dstKey, DstDataGrid)
                .RegisterInstance(srcKey, SrcLocationGrid)
                .RegisterInstance(dstKey, DstLocationGrid)
                .RegisterInstance(srcKey, SrcViewRectangle)
                .RegisterInstance(dstKey, DstViewRectangle)
                .RegisterInstance(srcKey, SrcValueTextBox)
                .RegisterInstance(dstKey, DstValueTextBox);
        }

        private void InitializeEventListeners()
        {
            var srcEventHandler = new DiffViewEventHandler(srcKey);
            var dstEventHandler = new DiffViewEventHandler(dstKey);

            DataGridEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            DataGridEventDispatcher.Instance.Listeners.Add(dstEventHandler);
            LocationGridEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            LocationGridEventDispatcher.Instance.Listeners.Add(dstEventHandler);
            ViewportEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            ViewportEventDispatcher.Instance.Listeners.Add(dstEventHandler);
            ValueTextBoxEventDispatcher.Instance.Listeners.Add(srcEventHandler);
            ValueTextBoxEventDispatcher.Instance.Listeners.Add(dstEventHandler);
        }

        private void OnApplicationSettingUpdated()
        {
            var e = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchApplicationSettingUpdateEvent(e);
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            AttachWindowClosingHandler();

            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchParentLoadEvent(args);

            ExecuteDiff(isStartup: true);
            isLoaded = true;

            // In order to enable Ctrl + F immediately after startup.
            ToolExpander.IsExpanded = false;
        }

        private void AttachWindowClosingHandler()
        {
            if (isWindowClosingHandlerAttached)
                return;

            var window = Window.GetWindow(this);
            if (window == null)
                return;

            window.Closing += DiffWindow_Closing;
            isWindowClosingHandlerAttached = true;
        }

        private ExcelSheetDiffConfig CreateDiffConfig(FileSetting srcFileSetting, FileSetting dstFileSetting, bool isStartup)
        {
            var config = new ExcelSheetDiffConfig();
            config.UseSmartTableDiff = App.Instance.Setting.UseSmartTableDiff;

            config.SrcSheetIndex = SrcSheetCombobox.SelectedIndex;
            config.DstSheetIndex = DstSheetCombobox.SelectedIndex;

            if (srcFileSetting != null)
            {
                if (isStartup)
                    config.SrcSheetIndex = GetSheetIndex(srcFileSetting, SrcSheetCombobox.Items);

                config.SrcHeaderIndex = srcFileSetting.ColumnHeaderIndex;
                config.SrcRowHeaderIndex = srcFileSetting.RowHeaderIndex;
                config.SrcRowHeaderName = srcFileSetting.RowHeaderName ?? string.Empty;
            }

            if (dstFileSetting != null)
            {
                if (isStartup)
                    config.DstSheetIndex = GetSheetIndex(dstFileSetting, DstSheetCombobox.Items);

                config.DstHeaderIndex = dstFileSetting.ColumnHeaderIndex;
                config.DstRowHeaderIndex = dstFileSetting.RowHeaderIndex;
                config.DstRowHeaderName = dstFileSetting.RowHeaderName ?? string.Empty;
            }

            return config;
        }

        private int GetSheetIndex(FileSetting fileSetting, ItemCollection sheetNames)
        {
            if (fileSetting == null)
                return -1;

            var index = fileSetting.SheetIndex;
            if (!string.IsNullOrEmpty(fileSetting.SheetName))
                index = sheetNames.IndexOf(fileSetting.SheetName);

            if (index < 0 || index >= sheetNames.Count)
            {
                MessageBox.Show(Properties.Resources.Msg_OutofSheetRange);
                index = 0;
            }

            return index;
        }

        private void LocationGrid_MouseDown(object sender, MouseEventArgs e)
        {
            var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
            LocationGridEventDispatcher.Instance.DispatchMouseDownEvent(args, e);
        }

        private void LocationGrid_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
                LocationGridEventDispatcher.Instance.DispatchMouseDownEvent(args, e);
            }
        }

        private void LocationGrid_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
            LocationGridEventDispatcher.Instance.DispatchMouseWheelEvent(args, e);
        }

        private void DataGrid_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(sender as FastGridControl, container);
            DataGridEventDispatcher.Instance.DispatchSizeChangeEvent(args, e);
        }

        private void LocationGrid_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
            LocationGridEventDispatcher.Instance.DispatchSizeChangeEvent(args, e);
        }

        private void DataGrid_SelectedCellsChanged(object sender, FastWpfGrid.SelectionChangedEventArgs e)
        {
            var grid = copyTargetGrid = sender as FastGridControl;
            if (grid == null)
                return;

            copyTargetGrid = grid;

            var args = new DiffViewEventArgs<FastGridControl>(sender as FastGridControl, container);
            DataGridEventDispatcher.Instance.DispatchSelectedCellChangeEvent(args);

            if (!SrcDataGrid.CurrentCell.Row.HasValue || !DstDataGrid.CurrentCell.Row.HasValue)
                return;

            if (!SrcDataGrid.CurrentCell.Column.HasValue || !DstDataGrid.CurrentCell.Column.HasValue)
                return;

            if (SrcDataGrid.Model == null || DstDataGrid.Model == null)
                return;

            var srcValue =
                (SrcDataGrid.Model as DiffGridModel).GetCellText(SrcDataGrid.CurrentCell.Row.Value, SrcDataGrid.CurrentCell.Column.Value, true);
            var dstValue =
                (DstDataGrid.Model as DiffGridModel).GetCellText(DstDataGrid.CurrentCell.Row.Value, DstDataGrid.CurrentCell.Column.Value, true);

            UpdateValueDiff(srcValue, dstValue);

            if (App.Instance.Setting.AlwaysExpandCellDiff)
            {
                var a = new DiffViewEventArgs<RichTextBox>(null, container, TargetType.First);
                ValueTextBoxEventDispatcher.Instance.DispatchGotFocusEvent(a);
            }
        }

        private void ValueTextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<RichTextBox>(sender as RichTextBox, container, TargetType.First);
            ValueTextBoxEventDispatcher.Instance.DispatchGotFocusEvent(args);
        }

        private void ValueTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<RichTextBox>(sender as RichTextBox, container, TargetType.First);
            ValueTextBoxEventDispatcher.Instance.DispatchLostFocusEvent(args);
        }

        private string GetRichTextString(RichTextBox textBox)
        {
            var textRange = new TextRange(textBox.Document.ContentStart, textBox.Document.ContentEnd);

            return textRange.Text;
        }

        private IEnumerable<DiffResult<string>> DiffCellValue(IEnumerable<string> src, IEnumerable<string> dst)
        {
            var r = DiffUtil.Diff(src, dst);
            r = DiffUtil.Order(r, DiffOrderType.LazyDeleteFirst);
            return DiffUtil.OptimizeCaseDeletedFirst(r);
        }

        private string ConvertWhiteSpaces(string str)
        {
            return new string(str.Select(c =>
            {
                if (Encoding.UTF8.GetByteCount(c.ToString()) == 1)
                    return ' ';
                else
                    return '　';

            }).ToArray());
        }

        private string ConvertWhiteSpaces(char c)
        {
            if (Encoding.UTF8.GetByteCount(c.ToString()) == 1)
                return " ";
            else
                return "　";
        }

        private void DiffModifiedLine(IEnumerable<DiffResult<char>> results, List<Tuple<string, Color?>> ranges, bool isSrc)
        {
            var splited = results.SplitByRegularity((items, current) => items.Last().Status.Equals(current.Status)).ToList();

            foreach (var sr in splited)
            {
                var status = sr.First().Status;
                if (status == DiffStatus.Equal)
                {
                    ranges.Add(Tuple.Create<string, Color?>(new string(sr.Select(r => r.Obj1).ToArray()), null));
                }
                else if (status == DiffStatus.Modified)
                {
                    var str = new string(sr.Select(r => isSrc ? r.Obj1 : r.Obj2).ToArray());
                    ranges.Add(Tuple.Create<string, Color?>(str, EMColor.LightOrange));
                }
                else if (status == DiffStatus.Deleted)
                {
                    var str = new string(sr.Select(r => r.Obj1).ToArray());
                    ranges.Add(Tuple.Create<string, Color?>(str, EMColor.LightGray));
                }
                else if (status == DiffStatus.Inserted)
                {
                    var str = new string(sr.Select(r => r.Obj2).ToArray());
                    ranges.Add(Tuple.Create<string, Color?>(str, EMColor.Orange));
                }
            }

            ranges.Add(Tuple.Create<string, Color?>("\n", null));
        }

        private void DiffEqualLine(DiffResult<string> lineDiffResult, List<Tuple<string, Color?>> ranges)
        {
            ranges.Add(Tuple.Create<string, Color?>(lineDiffResult.Obj1, null));
            ranges.Add(Tuple.Create<string, Color?>("\n", null));
        }

        private void DiffDeletedLine(DiffResult<string> lineDiffResult, List<Tuple<string, Color?>> ranges, bool isSrc)
        {
            var str = isSrc ? lineDiffResult.Obj1 : ConvertWhiteSpaces(lineDiffResult.Obj1.ToString());
            ranges.Add(Tuple.Create<string, Color?>(str, isSrc ? EMColor.LightGray : EMColor.LightGray));
            ranges.Add(Tuple.Create<string, Color?>("\n", null));
        }

        private void DiffInsertedLine(DiffResult<string> lineDiffResult, List<Tuple<string, Color?>> ranges, bool isSrc)
        {
            var str = isSrc ? ConvertWhiteSpaces(lineDiffResult.Obj2) : lineDiffResult.Obj2;
            ranges.Add(Tuple.Create<string, Color?>(str, isSrc ? EMColor.LightGray : EMColor.Orange));
            ranges.Add(Tuple.Create<string, Color?>("\n", null));
        }

        private void UpdateValueDiff(string srcValue, string dstValue)
        {
            SrcValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Clear();
            DstValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Clear();

            var srcLines = srcValue.Split('\n').Select(s => s.TrimEnd());
            var dstLines = dstValue.Split('\n').Select(s => s.TrimEnd());

            var lineDiffResults = DiffCellValue(srcLines, dstLines).ToList();

            var srcRange = new List<Tuple<string, Color?>>();
            var dstRange = new List<Tuple<string, Color?>>();
            foreach (var lineDiffResult in lineDiffResults)
            {
                if (lineDiffResult.Status == DiffStatus.Equal)
                {
                    DiffEqualLine(lineDiffResult, srcRange);
                    DiffEqualLine(lineDiffResult, dstRange);
                }
                else if (lineDiffResult.Status == DiffStatus.Modified)
                {
                    var charDiffResults = DiffUtil.Diff(lineDiffResult.Obj1, lineDiffResult.Obj2);
                    charDiffResults = DiffUtil.Order(charDiffResults, DiffOrderType.LazyDeleteFirst);
                    charDiffResults = DiffUtil.OptimizeCaseDeletedFirst(charDiffResults);

                    DiffModifiedLine(charDiffResults.Where(r => r.Status != DiffStatus.Inserted), srcRange, true);
                    DiffModifiedLine(charDiffResults.Where(r => r.Status != DiffStatus.Deleted), dstRange, false);
                }
                else if (lineDiffResult.Status == DiffStatus.Deleted)
                {
                    DiffDeletedLine(lineDiffResult, srcRange, true);
                    DiffDeletedLine(lineDiffResult, dstRange, false);
                }
                else if (lineDiffResult.Status == DiffStatus.Inserted)
                {
                    DiffInsertedLine(lineDiffResult, srcRange, true);
                    DiffInsertedLine(lineDiffResult, dstRange, false);
                }
            }

            foreach (var r in srcRange)
            {
                var bc = r.Item2.HasValue ? new SolidColorBrush(r.Item2.Value) : new SolidColorBrush();
                SrcValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Add(new Run(r.Item1) { Background = bc });
            }

            foreach (var r in dstRange)
            {
                var bc = r.Item2.HasValue ? new SolidColorBrush(r.Item2.Value) : new SolidColorBrush();
                DstValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Add(new Run(r.Item1) { Background = bc });
            }
        }

        private void DiffButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteDiff();
        }

        private void SheetCombobox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (!isLoaded || isExecutingDiff || isSynchronizingSheetSelection)
                return;

            if (SrcSheetCombobox.SelectedIndex < 0 || DstSheetCombobox.SelectedIndex < 0)
                return;

            SynchronizeOtherSheetCombobox(sender as ComboBox);

            ExecuteDiff();
        }

        private void SynchronizeOtherSheetCombobox(ComboBox changedCombobox)
        {
            if (changedCombobox == null || changedCombobox.SelectedItem == null)
                return;

            var targetCombobox = changedCombobox == SrcSheetCombobox
                ? DstSheetCombobox
                : changedCombobox == DstSheetCombobox ? SrcSheetCombobox : null;

            if (targetCombobox == null)
                return;

            var sheetName = changedCombobox.SelectedItem.ToString();
            var targetIndex = targetCombobox.Items.IndexOf(sheetName);
            if (targetIndex < 0 || targetIndex == targetCombobox.SelectedIndex)
                return;

            isSynchronizingSheetSelection = true;
            try
            {
                targetCombobox.SelectedIndex = targetIndex;
            }
            finally
            {
                isSynchronizingSheetSelection = false;
            }
        }

        private ExcelSheetReadConfig CreateReadConfig()
        {
            var setting = ((App)Application.Current).Setting;

            return new ExcelSheetReadConfig()
            {
                TrimFirstBlankRows = setting.SkipFirstBlankRows,
                TrimFirstBlankColumns = setting.SkipFirstBlankColumns,
                TrimLastBlankRows = setting.TrimLastBlankRows,
                TrimLastBlankColumns = setting.TrimLastBlankColumns,
            };
        }

        private Tuple<ExcelWorkbook, ExcelWorkbook> ReadWorkbooks()
        {
            ExcelWorkbook swb = null;
            ExcelWorkbook dwb = null;
            var srcPath = SrcPathTextBox.Text;
            var dstPath = DstPathTextBox.Text;
            ProgressWindow.DoWorkWithModal(progress =>
            {
                progress.Report(Properties.Resources.Msg_ReadingFiles);

                var config = CreateReadConfig();
                swb = ExcelWorkbook.Create(srcPath, config);
                dwb = ExcelWorkbook.Create(dstPath, config);
            });

            return Tuple.Create(swb, dwb);
        }

        private Tuple<FileSetting, FileSetting> FindFileSettings(bool isStartup)
        {
            FileSetting srcSetting = null;
            FileSetting dstSetting = null;
            var srcPath = SrcPathTextBox.Text;
            var dstPath = DstPathTextBox.Text;
            if (!IgnoreFileSettingCheckbox.IsChecked.Value)
            {
                srcSetting =
                    FindFilseSetting(Path.GetFileName(srcPath), SrcSheetCombobox.SelectedIndex, SrcSheetCombobox.SelectedItem.ToString(), isStartup);

                dstSetting =
                    FindFilseSetting(Path.GetFileName(dstPath), DstSheetCombobox.SelectedIndex, DstSheetCombobox.SelectedItem.ToString(), isStartup);

                diffConfig = CreateDiffConfig(srcSetting, dstSetting, isStartup);
            }
            else
            {
                diffConfig = new ExcelSheetDiffConfig();
                diffConfig.UseSmartTableDiff = App.Instance.Setting.UseSmartTableDiff;

                diffConfig.SrcSheetIndex = Math.Max(SrcSheetCombobox.SelectedIndex, 0);
                diffConfig.DstSheetIndex = Math.Max(DstSheetCombobox.SelectedIndex, 0);
            }

            return Tuple.Create(srcSetting, dstSetting);
        }

        private ExcelSheetDiff ExecuteDiff(ExcelSheet srcSheet, ExcelSheet dstSheet)
        {
            ExcelSheetDiff diff = null;
            ProgressWindow.DoWorkWithModal(progress =>
            {
                progress.Report(Properties.Resources.Msg_ExtractingDiff);
                diff = ExcelSheet.Diff(srcSheet, dstSheet, diffConfig);
            });

            return diff;
        }

        private void ExecuteDiff(bool isStartup = false)
        {
            if (isExecutingDiff)
                return;

            if (!File.Exists(SrcPathTextBox.Text) || !File.Exists(DstPathTextBox.Text))
                return;

            isExecutingDiff = true;
            try
            {
                var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
                DataGridEventDispatcher.Instance.DispatchPreExecuteDiffEvent(args);

                var workbooks = ReadWorkbooks();
                var srcWorkbook = workbooks.Item1;
                var dstWorkbook = workbooks.Item2;

                var fileSettings = FindFileSettings(isStartup);
                var srcFileSetting = fileSettings.Item1;
                var dstFileSetting = fileSettings.Item2;

                SrcSheetCombobox.SelectedIndex = diffConfig.SrcSheetIndex;
                DstSheetCombobox.SelectedIndex = diffConfig.DstSheetIndex;

                var srcSheet = srcWorkbook.Sheets[SrcSheetCombobox.SelectedItem.ToString()];
                var dstSheet = dstWorkbook.Sheets[DstSheetCombobox.SelectedItem.ToString()];

                if (srcSheet.Rows.Count > 10000 || dstSheet.Rows.Count > 10000)
                    MessageBox.Show(Properties.Resources.Msg_WarnSize);

                var diff = ExecuteDiff(srcSheet, dstSheet);
                SrcDataGrid.Model = new DiffGridModel(diff, DiffType.Source);
                DstDataGrid.IsReadOnly = !CanEditDst();
                DstDataGrid.Model = new DiffGridModel(diff, DiffType.Dest, CanEditDst() ? (Action<int, int, string>)ApplyDstCellEdit : null);

                args = new DiffViewEventArgs<FastGridControl>(SrcDataGrid, container);
                DataGridEventDispatcher.Instance.DispatchFileSettingUpdateEvent(args, srcFileSetting);

                args = new DiffViewEventArgs<FastGridControl>(DstDataGrid, container);
                DataGridEventDispatcher.Instance.DispatchFileSettingUpdateEvent(args, dstFileSetting);

                args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
                DataGridEventDispatcher.Instance.DispatchDisplayFormatChangeEvent(args, ShowOnlyDiffRadioButton.IsChecked.Value);
                DataGridEventDispatcher.Instance.DispatchPostExecuteDiffEvent(args);

                var summary = diff.CreateSummary();
                GetViewModel().UpdateDiffSummary(summary);

                if (!App.Instance.KeepFileHistory)
                    App.Instance.UpdateRecentFiles(SrcPathTextBox.Text, DstPathTextBox.Text);

                if (App.Instance.Setting.NotifyEqual && !summary.HasDiff)
                    MessageBox.Show(Properties.Resources.Message_NoDiff);

                if (App.Instance.Setting.FocusFirstDiff)
                    MoveNextModifiedCell();
            }
            finally
            {
                isExecutingDiff = false;
            }
        }

        private bool CanEditDst()
        {
            var viewModel = GetViewModel();
            return viewModel != null
                && viewModel.DstEditingEnabled
                && IsSamePath(DstPathTextBox.Text, viewModel.DstEditWorkingPath)
                && File.Exists(DstPathTextBox.Text)
                && ExcelWorkbookEditor.IsEditableWorkbook(DstPathTextBox.Text);
        }

        private static bool IsSamePath(string left, string right)
        {
            if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
                return false;

            return string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);
        }

        private string GetDstSheetName()
        {
            return DstSheetCombobox.SelectedItem == null ? string.Empty : DstSheetCombobox.SelectedItem.ToString();
        }

        private DiffGridModel GetDstGridModel()
        {
            return DstDataGrid.Model as DiffGridModel;
        }

        private void ApplyDstCellEdit(int rowIndex, int columnIndex, string value)
        {
            if (!CanEditDst())
                return;

            ExcelWorkbookEditor.SetCellText(DstPathTextBox.Text, GetDstSheetName(), rowIndex, columnIndex, value);
            MarkDstEditedAndRefresh();
        }

        private void MarkDstEditedAndRefresh()
        {
            var viewModel = GetViewModel();
            if (viewModel != null)
                viewModel.MarkEdited();

            Dispatcher.BeginInvoke((Action)(() => ExecuteDiff()));
        }

        private bool TryGetCurrentEditableDstCell(out int rowIndex, out int columnIndex)
        {
            rowIndex = -1;
            columnIndex = -1;

            var model = GetDstGridModel();
            if (model == null)
                return false;

            var current = DstDataGrid.CurrentModelCell;
            if (!current.IsCell)
                return false;

            return model.TryGetEditableCellAddress(current.Row.Value, current.Column.Value, out rowIndex, out columnIndex);
        }

        private List<int> GetSelectedEditableDstRows()
        {
            var model = GetDstGridModel();
            if (model == null)
                return new List<int>();

            var rows = DstDataGrid.GetSelectedModelCells()
                .Where(c => c.Row.HasValue)
                .Select(c => c.Row.Value)
                .Distinct()
                .ToList();

            if (!rows.Any() && DstDataGrid.CurrentModelCell.Row.HasValue)
                rows.Add(DstDataGrid.CurrentModelCell.Row.Value);

            var originalRows = new List<int>();
            foreach (var row in rows)
            {
                int originalRow;
                if (model.TryGetEditableRowAddress(row, out originalRow))
                    originalRows.Add(originalRow);
            }

            return originalRows.Distinct().OrderBy(i => i).ToList();
        }

        private IEnumerable<IEnumerable<string>> ParseClipboardTable(string text)
        {
            text = (text ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n");
            var lines = text.Split('\n').ToList();
            if (lines.Any() && lines.Last() == string.Empty)
                lines.RemoveAt(lines.Count - 1);

            return lines.Select(line => line.Split('\t').AsEnumerable()).ToList();
        }

        private string GetClipboardTextOrCachedRows()
        {
            try
            {
                var text = Clipboard.GetText();
                if (!string.IsNullOrEmpty(text))
                    return text;
            }
            catch
            {
            }

            return copiedRowsText;
        }

        private bool PasteIntoDst()
        {
            if (!CanEditDst())
                return false;

            var model = GetDstGridModel();
            if (model == null)
                return false;

            var startCell = DstDataGrid.CurrentModelCell;
            if (!startCell.IsCell)
                return false;

            var text = Clipboard.GetText();
            if (string.IsNullOrEmpty(text))
                return false;

            var values = ParseClipboardTable(text).ToList();
            if (!values.Any())
                return false;

            var edits = new List<Tuple<int, int, string>>();
            var rowOffset = 0;
            foreach (var rowValues in values)
            {
                var columnOffset = 0;
                foreach (var value in rowValues)
                {
                    int rowIndex;
                    int columnIndex;
                    if (model.TryGetEditableCellAddress(
                        startCell.Row.Value + rowOffset,
                        startCell.Column.Value + columnOffset,
                        out rowIndex,
                        out columnIndex))
                    {
                        edits.Add(Tuple.Create(rowIndex, columnIndex, value));
                    }

                    columnOffset++;
                }

                rowOffset++;
            }

            if (!edits.Any())
                return false;

            ExcelWorkbookEditor.SetCellTexts(DstPathTextBox.Text, GetDstSheetName(), edits);
            MarkDstEditedAndRefresh();
            return true;
        }

        private bool ClearSelectedDstCells()
        {
            if (!CanEditDst())
                return false;

            var model = GetDstGridModel();
            if (model == null)
                return false;

            var cells = DstDataGrid.GetSelectedModelCells().Where(c => c.IsCell).ToList();
            if (!cells.Any() && DstDataGrid.CurrentModelCell.IsCell)
                cells.Add(DstDataGrid.CurrentModelCell);

            var changed = false;
            foreach (var cell in cells)
            {
                int rowIndex;
                int columnIndex;
                if (!model.TryGetEditableCellAddress(cell.Row.Value, cell.Column.Value, out rowIndex, out columnIndex))
                    continue;

                ExcelWorkbookEditor.SetCellText(DstPathTextBox.Text, GetDstSheetName(), rowIndex, columnIndex, string.Empty);
                changed = true;
            }

            if (changed)
                MarkDstEditedAndRefresh();

            return changed;
        }

        private bool SaveDstEdits()
        {
            var viewModel = GetViewModel();
            if (viewModel == null || !viewModel.HasUnsavedEdits)
                return true;

            try
            {
                var backupPath = ExcelWorkbookEditor.SaveWorkingCopyToOriginal(DstPathTextBox.Text, viewModel.DstEditPath);
                viewModel.MarkSaved(backupPath);
                MessageBox.Show("Saved destination edits.\nBackup:\n" + backupPath);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to save destination edits.\n\n" + ex.Message);
                return false;
            }
        }

        private void DiscardDstEdits()
        {
            var viewModel = GetViewModel();
            if (viewModel == null || !viewModel.HasUnsavedEdits)
                return;

            try
            {
                ExcelWorkbookEditor.RestoreOriginalToWorkingCopy(viewModel.DstEditPath, DstPathTextBox.Text);
                viewModel.MarkDiscarded();
                ExecuteDiff();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to discard destination edits.\n\n" + ex.Message);
            }
        }

        private void DiffWindow_Closing(object sender, CancelEventArgs e)
        {
            var viewModel = GetViewModel();
            if (viewModel == null || !viewModel.HasUnsavedEdits)
                return;

            var result = MessageBox.Show("Save destination edits before closing?", "ExcelMerge", MessageBoxButton.YesNoCancel);
            if (result == MessageBoxResult.Cancel)
            {
                e.Cancel = true;
                return;
            }

            if (result == MessageBoxResult.Yes && !SaveDstEdits())
                e.Cancel = true;
        }

        private FileSetting FindFilseSetting(string fileName, int sheetIndex, string sheetName, bool isStartup)
        {
            var results = new List<FileSetting>();
            foreach (var setting in App.Instance.Setting.FileSettings)
            {
                if (setting.UseRegex)
                {
                    var regex = new System.Text.RegularExpressions.Regex(setting.Name);

                    if (regex.IsMatch(fileName))
                        results.Add(setting);
                }
                else
                {
                    if (setting.ExactMatch)
                    {
                        if (setting.Name == fileName)
                            results.Add(setting);
                    }
                    else
                    {
                        if (fileName.Contains(setting.Name))
                            results.Add(setting);
                    }
                }
            }

            if (isStartup)
                return results.FirstOrDefault(r => r.IsStartupSheet) ?? results.FirstOrDefault() ?? null;

            return results.FirstOrDefault(r => r.SheetName == sheetName) ?? results.FirstOrDefault(r => r.SheetIndex == sheetIndex) ?? null;
        }

        private void SetRowHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchRowHeaderChagneEvent(args);
                }
            }
        }

        private void ResetRowHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchRowHeaderResetEvent(args);
                }
            }
        }

        private void SetColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchColumnHeaderChangeEvent(args);
                }
            }
        }

        private void ResetColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchColumnHeaderResetEvent(args);
                }
            }
        }

        private void SwapButton_Click(object sender, RoutedEventArgs e)
        {
            Swap();
        }

        private void Swap()
        {
            var srcTmp = SrcSheetCombobox.SelectedIndex;
            var dstTmp = DstSheetCombobox.SelectedIndex;

            var tmp = SrcPathTextBox.Text;
            SrcPathTextBox.Text = DstPathTextBox.Text;
            DstPathTextBox.Text = tmp;

            diffConfig.SrcSheetIndex = dstTmp;
            diffConfig.DstSheetIndex = srcTmp;

            ExecuteDiff();
        }

        private void DiffByHeaderSrc_Click(object sender, RoutedEventArgs e)
        {
            var headerIndex = SrcDataGrid.CurrentCell.Row.HasValue ? SrcDataGrid.CurrentCell.Row.Value : -1;

            diffConfig.SrcHeaderIndex = headerIndex;

            ExecuteDiff();
        }

        private void DiffByHeaderDst_Click(object sender, RoutedEventArgs e)
        {
            var headerIndex = DstDataGrid.CurrentCell.Row.HasValue ? DstDataGrid.CurrentCell.Row.Value : -1;

            diffConfig.DstHeaderIndex = headerIndex;

            ExecuteDiff();
        }

        private void ShowAllRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchDisplayFormatChangeEvent(args, false);
        }

        private void ShowOnlyDiffRadioButton_Checked(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchDisplayFormatChangeEvent(args, true);
        }

        private bool ValidateDataGrids()
        {
            return SrcDataGrid.Model != null && DstDataGrid.Model != null;
        }

        private void ValuteTextBox_ScrollChanged(object sender, RoutedEventArgs e)
        {
            var args = new DiffViewEventArgs<RichTextBox>(sender as RichTextBox, container);
            ValueTextBoxEventDispatcher.Instance.DispatchScrolledEvent(args, (ScrollChangedEventArgs)e);
        }

        private void NextModifiedCellButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextModifiedCell();
        }

        private void MoveNextModifiedCell()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextModifiedCell(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void PrevModifiedCellButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevModifiedCell();
        }

        private void MovePrevModifiedCell()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousModifiedCell(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void NextModifiedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextModifiedRow();
        }

        private void MoveNextModifiedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextModifiedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void PrevModifiedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevModifiedRow();
        }

        private void MovePrevModifiedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousModifiedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void NextAddedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextAddedRow();
        }

        private void MoveNextAddedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextAddedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void PrevAddedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevAddedRow();
        }

        private void MovePrevAddedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousAddedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void NextRemovedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextRemovedRow();
        }

        private void MoveNextRemovedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextRemovedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void PrevRemovedRowButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevRemovedRow();
        }

        private void MovePrevRemovedRow()
        {
            if (!ValidateDataGrids())
                return;

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousRemovedRow(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void PrevMatchCellButton_Click(object sender, RoutedEventArgs e)
        {
            MovePrevMatchCell();
        }

        private void MovePrevMatchCell()
        {
            if (!ValidateDataGrids())
                return;

            var text = SearchTextCombobox.Text;
            if (string.IsNullOrEmpty(text))
                return;

            var history = App.Instance.Setting.SearchHistory.ToList();
            if (history.Contains(text))
                history.Remove(text);

            history.Insert(0, text);
            history = history.Take(10).ToList();

            App.Instance.Setting.SearchHistory = new ObservableCollection<string>(history);
            App.Instance.Setting.Save();

            SearchTextCombobox.ItemsSource = App.Instance.Setting.SearchHistory.ToList();

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetPreviousMatchCell(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell, text,
                ExactMatchCheckBox.IsChecked.Value, CaseSensitiveCheckBox.IsChecked.Value, RegexCheckBox.IsChecked.Value, ShowOnlyDiffRadioButton.IsChecked.Value);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private void NextMatchCellButton_Click(object sender, RoutedEventArgs e)
        {
            MoveNextMatchCell();
        }

        private void MoveNextMatchCell()
        {
            if (!ValidateDataGrids())
                return;

            var text = SearchTextCombobox.Text;
            if (string.IsNullOrEmpty(text))
                return;

            var history = App.Instance.Setting.SearchHistory.ToList();
            if (history.Contains(text))
                history.Remove(text);

            history.Insert(0, text);
            history = history.Take(10).ToList();

            App.Instance.Setting.SearchHistory = new ObservableCollection<string>(history);
            App.Instance.Setting.Save();

            SearchTextCombobox.ItemsSource = App.Instance.Setting.SearchHistory.ToList();

            var nextCell = (SrcDataGrid.Model as DiffGridModel).GetNextMatchCell(
                SrcDataGrid.CurrentCell.IsEmpty ? FastGridCellAddress.Zero : SrcDataGrid.CurrentCell, text,
                ExactMatchCheckBox.IsChecked.Value, CaseSensitiveCheckBox.IsChecked.Value, RegexCheckBox.IsChecked.Value, ShowOnlyDiffRadioButton.IsChecked.Value);
            if (nextCell.IsEmpty)
                return;

            SrcDataGrid.CurrentCell = nextCell;
        }

        private FastGridControl GetCopySourceGrid(object sender)
        {
            var menuItem = sender as MenuItem;
            var contextMenu = menuItem == null ? null : menuItem.Parent as ContextMenu;
            var grid = contextMenu == null ? null : contextMenu.PlacementTarget as FastGridControl;
            return grid ?? copyTargetGrid;
        }

        private void SetCopiedRowsText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return;

            copiedRowsText = text;
            try
            {
                Clipboard.SetText(text, TextDataFormat.UnicodeText);
            }
            catch
            {
                // The in-window cache still supports paste when another process locks the clipboard.
            }
        }

        private void CopyToClipboardSelectedCells(string separator, FastGridControl grid)
        {
            if (grid == null)
                return;

            var model = grid.Model as DiffGridModel;
            if (model == null)
                return;

            var selectedCells = grid.GetSelectedModelCells()
                .OrderBy(cell => cell.Row.Value)
                .ThenBy(cell => cell.Column.Value)
                .ToList();

            string text;
            if (selectedCells.Any())
            {
                text = string.Join(Environment.NewLine,
                    selectedCells
                        .GroupBy(cell => cell.Row.Value)
                        .OrderBy(group => group.Key)
                        .Select(group => string.Join(separator,
                            group.Select(cell => model.GetCellText(cell.Row.Value, cell.Column.Value)))));
            }
            else
            {
                var currentCell = grid.CurrentModelCell;
                if (!currentCell.Row.HasValue)
                    return;

                text = string.Join(separator,
                    Enumerable.Range(0, model.ColumnCount)
                        .Select(column => model.GetCellText(currentCell.Row.Value, column)));
            }

            SetCopiedRowsText(text);
        }

        private void UserControl_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Right:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextModifiedCell();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.Left:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevModifiedCell();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.Down:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextModifiedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.Up:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevModifiedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.L:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextRemovedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.O:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevRemovedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.K:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MoveNextAddedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.I:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            MovePrevAddedRow();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.F8:
                    {
                        MovePrevMatchCell();
                        e.Handled = true;
                    }
                    break;
                case Key.F9:
                    {
                        MoveNextMatchCell();
                        e.Handled = true;
                    }
                    break;
                case Key.F:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            ToolExpander.IsExpanded = true;
                            SearchTextCombobox.Focus();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.C:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            CopyToClipboardSelectedCells(
                                Keyboard.IsKeyDown(Key.RightShift) || Keyboard.IsKeyDown(Key.LeftShift) ? "," : "\t",
                                copyTargetGrid);
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.V:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl)
                            && !(Keyboard.FocusedElement is TextBox)
                            && !(Keyboard.FocusedElement is RichTextBox)
                            && PasteIntoDst())
                        {
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.S:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            SaveDstEdits();
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.Delete:
                    {
                        if (!(Keyboard.FocusedElement is TextBox)
                            && !(Keyboard.FocusedElement is RichTextBox)
                            && ClearSelectedDstCells())
                        {
                            e.Handled = true;
                        }
                    }
                    break;
                case Key.B:
                    {
                        if (Keyboard.IsKeyDown(Key.LeftCtrl))
                        {
                            ShowLog();
                            e.Handled = true;
                        }
                    }
                    break;
            }
        }

        private void ShowLog()
        {
            var log = BuildCellBaseLog();

            (App.Current.MainWindow as MainWindow).WriteToConsole(log);
        }

        private void BuildCellBaseLog_Click(object sender, RoutedEventArgs e)
        {
            ShowLog();
        }

        private string BuildCellBaseLog()
        {
            var srcModel = SrcDataGrid.Model as DiffGridModel;
            if (srcModel == null)
                return string.Empty;

            var dstModel = DstDataGrid.Model as DiffGridModel;
            if (dstModel == null)
                return string.Empty;

            var builder = new StringBuilder();

            var selectedCells = SrcDataGrid.SelectedCells;

            var modifiedLogFormat = App.Instance.Setting.LogFormat;
            var addedLogFormat = App.Instance.Setting.AddedRowLogFormat;
            var removedLogFormat = App.Instance.Setting.RemovedRowLogFormat;

            foreach (var row in SrcDataGrid.SelectedCells.GroupBy(c => c.Row))
            {
                var rowHeaderText = srcModel.GetRowHeaderText(row.Key.Value);
                if (string.IsNullOrEmpty(rowHeaderText))
                    rowHeaderText = dstModel.GetRowHeaderText(row.Key.Value);

                if (dstModel.IsAddedRow(row.Key.Value, true))
                {
                    var log = addedLogFormat
                        .Replace("${ROW}", RemoveMultiLine(rowHeaderText));

                    builder.AppendLine(log);

                    continue;
                }

                if (dstModel.IsRemovedRow(row.Key.Value, true))
                {
                    var log = removedLogFormat
                        .Replace("${ROW}", RemoveMultiLine(rowHeaderText));

                    builder.AppendLine(log);

                    continue;
                }

                foreach (var cell in row)
                {
                    if (cell.Row.Value == srcModel.ColumnHeaderIndex)
                        continue;

                    var srcText = srcModel.GetCellText(cell, true);
                    var dstText = dstModel.GetCellText(cell, true);
                    if (srcText == dstText)
                        continue;

                    var colHeaderText = srcModel.GetColumnHeaderText(cell.Column.Value);

                    if (string.IsNullOrEmpty(colHeaderText))
                        colHeaderText = dstModel.GetColumnHeaderText(cell.Column.Value);

                    if (string.IsNullOrEmpty(srcText))
                        srcText = Properties.Resources.Word_Blank;

                    if (string.IsNullOrEmpty(dstText))
                        dstText = Properties.Resources.Word_Blank;

                    if (string.IsNullOrEmpty(rowHeaderText))
                        rowHeaderText = Properties.Resources.Word_Blank;

                    if (string.IsNullOrEmpty(colHeaderText))
                        colHeaderText = Properties.Resources.Word_Blank;

                    var log = modifiedLogFormat
                        .Replace("${ROW}", RemoveMultiLine(rowHeaderText))
                        .Replace("${COL}", RemoveMultiLine(colHeaderText))
                        .Replace("${LEFT}", RemoveMultiLine(srcText))
                        .Replace("${RIGHT}", RemoveMultiLine(dstText));

                    builder.AppendLine(log);
                }
            }

            return builder.ToString();
        }

        private string RemoveMultiLine(string log)
        {
            return log.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
        }

        private void CopyAsTsv_Click(object sender, RoutedEventArgs e)
        {
            CopyToClipboardSelectedCells("\t", GetCopySourceGrid(sender));
        }

        private void CopyAsCsv_Click(object sender, RoutedEventArgs e)
        {
            CopyToClipboardSelectedCells(",", GetCopySourceGrid(sender));
        }

        private void SaveDstEditButton_Click(object sender, RoutedEventArgs e)
        {
            SaveDstEdits();
        }

        private void DiscardDstEditButton_Click(object sender, RoutedEventArgs e)
        {
            DiscardDstEdits();
        }

        private List<int> GetSelectedVisualRows(FastGridControl grid)
        {
            var rows = grid.GetSelectedModelCells()
                .Where(c => c.Row.HasValue)
                .Select(c => c.Row.Value)
                .Distinct()
                .OrderBy(i => i)
                .ToList();

            if (!rows.Any() && grid.CurrentModelCell.Row.HasValue)
                rows.Add(grid.CurrentModelCell.Row.Value);

            return rows;
        }

        private void CopyRowsAsTsv(FastGridControl grid)
        {
            var model = grid.Model as DiffGridModel;
            if (model == null)
                return;

            var rows = GetSelectedVisualRows(grid);

            if (!rows.Any())
                return;

            var text = string.Join(Environment.NewLine,
                rows.Select(row => string.Join("\t",
                    Enumerable.Range(0, model.ColumnCount).Select(column => model.GetCellText(row, column)))));

            SetCopiedRowsText(text);
        }

        private void CopySrcRowsAsTsv_Click(object sender, RoutedEventArgs e)
        {
            CopyRowsAsTsv(SrcDataGrid);
        }

        private void CopyDstRowsAsTsv_Click(object sender, RoutedEventArgs e)
        {
            CopyRowsAsTsv(DstDataGrid);
        }

        private void InsertDstRowAbove_Click(object sender, RoutedEventArgs e)
        {
            InsertDstRows(above: true);
        }

        private void InsertDstRowBelow_Click(object sender, RoutedEventArgs e)
        {
            InsertDstRows(above: false);
        }

        private void InsertDstRows(bool above)
        {
            if (!CanEditDst())
                return;

            var rows = GetSelectedEditableDstRows();
            if (!rows.Any())
                return;

            var count = Math.Max(1, rows.Count);
            var rowIndex = above ? rows.Min() : rows.Max() + 1;
            ExcelWorkbookEditor.InsertBlankRows(DstPathTextBox.Text, GetDstSheetName(), rowIndex, count);
            MarkDstEditedAndRefresh();
        }

        private void InsertCopiedRowsBelow_Click(object sender, RoutedEventArgs e)
        {
            if (!CanEditDst())
                return;

            var model = GetDstGridModel();
            if (model == null)
                return;

            var text = GetClipboardTextOrCachedRows();
            if (string.IsNullOrEmpty(text))
            {
                MessageBox.Show("请先在左侧表格中复制一行。");
                return;
            }

            var values = ParseClipboardTable(text).Select(row => row.ToList()).ToList();
            if (!values.Any())
            {
                MessageBox.Show("复制的内容不是可粘贴的表格行。");
                return;
            }

            var rows = GetSelectedEditableDstRows();
            if (!rows.Any())
            {
                MessageBox.Show("请先在右侧表格中选择一个可编辑的数据行。");
                return;
            }

            var visualRows = GetSelectedVisualRows(DstDataGrid);
            if (!visualRows.Any())
            {
                MessageBox.Show("请先在右侧表格中选择插入位置。");
                return;
            }

            var anchorVisualRow = visualRows.Max();
            var insertIndex = rows.Max() + 1;

            var edits = new List<Tuple<int, int, string>>();
            for (var rowOffset = 0; rowOffset < values.Count; rowOffset++)
            {
                for (var columnOffset = 0; columnOffset < values[rowOffset].Count; columnOffset++)
                {
                    int originalRow;
                    int originalColumn;
                    if (model.TryGetEditableCellAddress(anchorVisualRow, columnOffset, out originalRow, out originalColumn))
                        edits.Add(Tuple.Create(insertIndex + rowOffset, originalColumn, values[rowOffset][columnOffset]));
                }
            }

            if (!edits.Any())
            {
                MessageBox.Show("复制的行没有可粘贴到右侧表格的单元格。");
                return;
            }

            ExcelWorkbookEditor.InsertBlankRows(DstPathTextBox.Text, GetDstSheetName(), insertIndex, values.Count);
            ExcelWorkbookEditor.SetCellTexts(DstPathTextBox.Text, GetDstSheetName(), edits);
            MarkDstEditedAndRefresh();
        }

        private void DeleteDstRow_Click(object sender, RoutedEventArgs e)
        {
            if (!CanEditDst())
                return;

            var rows = GetSelectedEditableDstRows();
            if (!rows.Any())
                return;

            ExcelWorkbookEditor.DeleteRows(DstPathTextBox.Text, GetDstSheetName(), rows);
            MarkDstEditedAndRefresh();
        }
    }
}
