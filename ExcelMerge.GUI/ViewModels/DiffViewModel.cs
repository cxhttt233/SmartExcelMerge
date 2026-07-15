using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.IO;
using System.Windows;
using Prism.Mvvm;
using FastWpfGrid;
using ExcelMerge.GUI.Settings;
using ExcelMerge.GUI.Behaviors;
using ExcelMerge.GUI.Commands;

namespace ExcelMerge.GUI.ViewModels
{
    public class DiffViewModel : BindableBase
    {
        private bool showLocationGridLine;
        public bool ShowLocationGridLine
        {
            get { return showLocationGridLine; }
            set { SetProperty(ref showLocationGridLine, value); }
        }

        private string srcPath;
        public string SrcPath
        {
            get { return srcPath; }
            set
            {
                SetProperty(ref srcPath, value);
                Settings.EMEnvironmentValue.Set("SRC", value);
                UpdateExecutableFlag();
            }
        }

        private string dstPath;
        public string DstPath
        {
            get { return dstPath; }
            set
            {
                SetProperty(ref dstPath, value);
                Settings.EMEnvironmentValue.Set("DST", value);
                UpdateExecutableFlag();
            }
        }

        private List<SheetSelectionItem> srcSheetNames;
        public List<SheetSelectionItem> SrcSheetNames
        {
            get { return srcSheetNames; }
            private set { SetProperty(ref srcSheetNames, value); }
        }

        private List<SheetSelectionItem> dstSheetNames;
        public List<SheetSelectionItem> DstSheetNames
        {
            get { return dstSheetNames; }
            private set { SetProperty(ref dstSheetNames, value); }
        }

        private int selectedSrcSheetIndex;
        public int SelectedSrcSheetIndex
        {
            get { return selectedSrcSheetIndex; }
            set { SetProperty(ref selectedSrcSheetIndex, value); }
        }

        private int selectedDstSheetIndex;
        public int SelectedDstSheetIndex
        {
            get { return selectedDstSheetIndex; }
            set { SetProperty(ref selectedDstSheetIndex, value); }
        }

        private bool executable;
        public bool Executable
        {
            get { return executable; }
            private set { SetProperty(ref executable, value); }
        }

        private string dstEditPath;
        public string DstEditPath
        {
            get { return dstEditPath; }
            private set { SetProperty(ref dstEditPath, value); }
        }

        private string dstEditWorkingPath;
        public string DstEditWorkingPath
        {
            get { return dstEditWorkingPath; }
            private set { SetProperty(ref dstEditWorkingPath, value); }
        }

        private bool dstEditingEnabled;
        public bool DstEditingEnabled
        {
            get { return dstEditingEnabled; }
            private set
            {
                if (SetProperty(ref dstEditingEnabled, value))
                    RaisePropertyChanged(nameof(IsDstReadOnly));
            }
        }

        public bool IsDstReadOnly
        {
            get { return !DstEditingEnabled; }
        }

        private bool hasUnsavedEdits;
        public bool HasUnsavedEdits
        {
            get { return hasUnsavedEdits; }
            private set
            {
                if (SetProperty(ref hasUnsavedEdits, value))
                    RaisePropertyChanged(nameof(EditStatusText));
            }
        }

        private string lastEditBackupPath;
        public string LastEditBackupPath
        {
            get { return lastEditBackupPath; }
            private set { SetProperty(ref lastEditBackupPath, value); }
        }

        public string EditStatusText
        {
            get
            {
                if (!DstEditingEnabled)
                    return "Read only";

                return HasUnsavedEdits ? "Dst edited" : "Dst editable";
            }
        }

        private int modifiedCellCount;
        public int ModifiedCellCount
        {
            get { return modifiedCellCount; }
            private set { SetProperty(ref modifiedCellCount, value); }
        }

        private int modifiedRowCount;
        public int ModifiedRowCount
        {
            get { return modifiedRowCount; }
            private set { SetProperty(ref modifiedRowCount, value); }
        }

        private int addedRowCount;
        public int AddedRowCount
        {
            get { return addedRowCount; }
            private set { SetProperty(ref addedRowCount, value); }
        }

        private int removedRowCount;
        public int RemovedRowCount
        {
            get { return removedRowCount; }
            private set { SetProperty(ref removedRowCount, value); }
        }

        private int addedColumnCount;
        public int AddedColumnCount
        {
            get { return addedColumnCount; }
            private set { SetProperty(ref addedColumnCount, value); }
        }

        private int removedColumnCount;
        public int RemovedColumnCount
        {
            get { return removedColumnCount; }
            private set { SetProperty(ref removedColumnCount, value); }
        }

        private DragAcceptDescription description;
        public DragAcceptDescription Description
        {
            get { return description; }
            private set { SetProperty(ref description, value); }
        }

        public DiffViewModel()
        {
            Description = new DragAcceptDescription();
            Description.DragDrop += DragDrop;
            Description.DragDrop += DragOver;

            SrcPath = string.Empty;
            DstPath = string.Empty;
        }

        public DiffViewModel(string src, string dst, MainWindowViewModel mwv) : this()
        {
            SrcPath = src;
            DstPath = dst;

            mwv.PropertyChanged += Mwv_PropertyChanged;
        }

        public DiffViewModel(CommandLineOption option, MainWindowViewModel mwv)
            : this(option.SrcPath, option.DstPath, mwv)
        {
            DstEditPath = option.DstEditPath;
            DstEditWorkingPath = option.DstPath;
            DstEditingEnabled = option.IsDstEditable && File.Exists(DstEditPath) && ExcelWorkbookEditor.IsEditableWorkbook(DstPath);
        }

        public void UpdateDiffSummary(ExcelSheetDiffSummary summary)
        {
            ModifiedCellCount = summary.ModifiedCellCount;
            ModifiedRowCount = summary.ModifiedRowCount;
            AddedRowCount = summary.AddedRowCount;
            RemovedRowCount = summary.RemovedRowCount;
            AddedColumnCount = summary.AddedColumnCount;
            RemovedColumnCount = summary.RemovedColumnCount;
        }

        public void UpdateSheetDiffStates(IEnumerable<string> differingSheetNames)
        {
            var names = new HashSet<string>(differingSheetNames ?? Enumerable.Empty<string>(), StringComparer.Ordinal);
            UpdateSheetDiffStates(SrcSheetNames, names);
            UpdateSheetDiffStates(DstSheetNames, names);
        }

        public void UpdateSheetDiffState(string sheetName, bool hasDiff)
        {
            if (string.IsNullOrEmpty(sheetName))
                return;

            UpdateSheetDiffState(SrcSheetNames, sheetName, hasDiff);
            UpdateSheetDiffState(DstSheetNames, sheetName, hasDiff);
        }

        public void MarkEdited()
        {
            HasUnsavedEdits = true;
        }

        public void MarkSaved(string backupPath)
        {
            LastEditBackupPath = backupPath;
            HasUnsavedEdits = false;
        }

        public void MarkDiscarded()
        {
            HasUnsavedEdits = false;
        }

        private void Mwv_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SrcPath))
            {
                var vm = sender as MainWindowViewModel;
                if (vm != null)
                {
                    var prop = typeof(MainWindowViewModel).GetProperties().FirstOrDefault(p => p.Name == e.PropertyName);
                    if (prop != null)
                    {
                        SrcPath = prop.GetValue(vm) as string;
                    }
                }
            }
            else if (e.PropertyName == nameof(DstPath))
            {
                var vm = sender as MainWindowViewModel;
                if (vm != null)
                {
                    var prop = typeof(MainWindowViewModel).GetProperties().FirstOrDefault(p => p.Name == e.PropertyName);
                    if (prop != null)
                    {
                        DstPath = prop.GetValue(vm) as string;
                    }
                }
            }
        }

        private void DragDrop(DragEventArgs e)
        {
            var paths = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (paths == null || !paths.Any())
                return;

            var target = e.Source as FrameworkElement;
            if (target == null)
                return;

            OnDragDrop(paths, target);
        }

        protected virtual void OnDragDrop(string[] filePath, FrameworkElement target)
        {
            if (filePath.Length > 1)
            {
                SrcPath = filePath[1];
                DstPath = filePath[0];

                return;
            }

            var tag = Convert.ToInt32(target.Tag);
            if (tag == 0)
                SrcPath = filePath[0];
            else if (tag == 1)
                DstPath = filePath[0];
        }

        private void DragOver(DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, true))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;

            e.Handled = true;
        }

        private void UpdateExecutableFlag()
        {
            var existsSrc = File.Exists(SrcPath);
            var existsDst = File.Exists(DstPath);

            if (existsSrc)
            {
                SrcSheetNames = ExcelWorkbook.GetSheetNames(SrcPath)
                    .Select(name => new SheetSelectionItem(name))
                    .ToList();
                SelectedSrcSheetIndex = 0;
            }
            else
            {
                SrcSheetNames = new List<SheetSelectionItem>();
                SelectedSrcSheetIndex = -1;
            }

            if (existsDst)
            {
                DstSheetNames = ExcelWorkbook.GetSheetNames(DstPath)
                    .Select(name => new SheetSelectionItem(name))
                    .ToList();
                SelectedDstSheetIndex = 0;
            }
            else
            {
                DstSheetNames = new List<SheetSelectionItem>();
                SelectedDstSheetIndex = -1;
            }

            Executable = existsSrc && existsDst;
        }

        private static void UpdateSheetDiffStates(IEnumerable<SheetSelectionItem> items, ISet<string> differingSheetNames)
        {
            if (items == null)
                return;

            foreach (var item in items)
                item.HasDiff = differingSheetNames.Contains(item.Name);
        }

        private static void UpdateSheetDiffState(IEnumerable<SheetSelectionItem> items, string sheetName, bool hasDiff)
        {
            if (items == null)
                return;

            foreach (var item in items.Where(i => string.Equals(i.Name, sheetName, StringComparison.Ordinal)))
                item.HasDiff = hasDiff;
        }
    }

    public sealed class SheetSelectionItem : BindableBase
    {
        private bool hasDiff;

        public SheetSelectionItem(string name)
        {
            Name = name ?? string.Empty;
        }

        public string Name { get; private set; }

        public bool HasDiff
        {
            get { return hasDiff; }
            set { SetProperty(ref hasDiff, value); }
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
