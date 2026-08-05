from __future__ import annotations

import html
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def write_text(path: Path, text: str, *, bom: bool = True) -> None:
    path.write_text(text, encoding="utf-8-sig" if bom else "utf-8")


def replace_required(path: Path, replacements: dict[str, str]) -> None:
    text = read_text(path)
    for old, new in replacements.items():
        if old not in text:
            raise RuntimeError(f"Source text not found in {path}: {old}")
        text = text.replace(old, new)
    write_text(path, text)


def replace_resx_values(path: Path, translations: dict[str, str]) -> None:
    text = read_text(path)
    for key, value in translations.items():
        pattern = re.compile(
            rf'(<data\s+name="{re.escape(key)}"[^>]*>\s*<value>)(.*?)(</value>)',
            re.DOTALL,
        )
        escaped = html.escape(value, quote=False)
        text, count = pattern.subn(
            lambda match: match.group(1) + escaped + match.group(3),
            text,
            count=1,
        )
        if count != 1:
            raise RuntimeError(f"Chinese resource key not found or duplicated: {key}")

    declaration = re.match(r"<\?xml[^>]+\?>", text)
    if declaration:
        text = text[: declaration.start()] + '<?xml version="1.0" encoding="utf-8"?>' + text[declaration.end() :]
    write_text(path, text, bom=False)


def add_chinese_resource_to_project(path: Path) -> None:
    text = read_text(path)
    if 'EmbeddedResource Include="Properties\\Resources.zh-CN.resx"' in text:
        return

    pattern = re.compile(
        r'(\s*<EmbeddedResource Include="Properties\\Resources\.ja-JP\.resx">\s*'
        r'<SubType>Designer</SubType>\s*</EmbeddedResource>)'
    )
    newline = "\r\n" if "\r\n" in text else "\n"
    addition = (
        r'\1'
        + newline
        + '    <EmbeddedResource Include="Properties\\Resources.zh-CN.resx">'
        + newline
        + '      <SubType>Designer</SubType>'
        + newline
        + '    </EmbeddedResource>'
    )
    text, count = pattern.subn(addition, text, count=1)
    if count != 1:
        raise RuntimeError("Could not locate Japanese resource entry in ExcelMerge.GUI.csproj")
    write_text(path, text)


def set_default_culture(path: Path) -> None:
    text = read_text(path)
    if 'app.Setting.Culture = "zh-CN";' not in text:
        marker = "app.Setting.EnsureCulture();"
        if marker not in text:
            raise RuntimeError("EnsureCulture call not found")
        newline = "\r\n" if "\r\n" in text else "\n"
        text = text.replace(
            marker,
            marker + newline + '            app.Setting.Culture = "zh-CN";',
            1,
        )

    replacements = {
        'throw new Exceptions.ExcelMergeException(true, $"Invalid argument.\\nargument:\\n{string.Join(" ", args)}");':
            'throw new Exceptions.ExcelMergeException(true, $"参数无效。\\n参数：\\n{string.Join(" ", args)}");',
        'var message = $"Execute external command ? \\n\\n------------------------------------\\n {exception.Message}\\n{exception.StackTrace}";':
            'var message = $"是否执行外部命令？ \\n\\n------------------------------------\\n {exception.Message}\\n{exception.StackTrace}";',
        'MessageBox.Show(message, "An error occurred.", MessageBoxButton.YesNo);':
            'MessageBox.Show(message, "发生错误", MessageBoxButton.YesNo);',
    }
    for old, new in replacements.items():
        if old not in text:
            raise RuntimeError(f"Source text not found in {path}: {old}")
        text = text.replace(old, new)
    write_text(path, text)


def main() -> None:
    replace_resx_values(
        ROOT / "ExcelMerge.GUI" / "Properties" / "Resources.zh-CN.resx",
        {
            "Button_ExtractDiff": "重新比较",
            "Label_SkipFirstBlankRows": "忽略开头的空白行",
            "Label_SkipFirstBlankColumns": "忽略开头的空白列",
            "Label_TrimLastBlankRows": "忽略末尾的空白行",
            "Label_TrimLastBlankColumns": "忽略末尾的空白列",
            "MenuItem_OpenSrcFile": "选择修改前文件",
            "MenuItem_OpenDstFile": "选择修改后文件",
            "MenuItem_OpenAsSrcFile": "设为修改前文件",
            "MenuItem_OpenAsDstFile": "设为修改后文件",
            "MenuItem_RecentFiles": "最近文件",
            "MenuItem_RecentFileSets": "最近对比记录",
            "ToolTip_SrcFilePath": "修改前文件路径",
            "ToolTip_DstFilePath": "修改后文件路径",
            "ToolTip_SrcSheet": "修改前工作表",
            "ToolTip_DstSheet": "修改后工作表",
            "Word_Done": "确定",
            "Word_Name": "名称",
            "Word_Executable": "可执行程序",
            "Word_Sheet": "工作表",
            "Word_SrcFile": "修改前文件",
            "Word_DstFile": "修改后文件",
            "ContextMenu_DiffAsHeader": "以此行作为列标题重新比较",
            "ContextMenu_SetRowHeader": "← 设置行标识列",
            "ContextMenu_ResetRowHeader": "← 重置行标识列",
            "ContextMenu_SetColumnHeader": "↑ 设置列标题行",
            "ContextMenu_ResetColumnHeader": "↑ 重置列标题行",
            "ContextMenu_BuildCellBaseLog": "输出变更日志",
            "ContextMenu_BuildColumnBaseLog": "输出按列汇总的变更日志",
            "ContextMenu_BuildRowBaseLog": "输出按行汇总的变更日志",
            "ContextMenu_CopyAsCsv": "复制为 CSV",
            "ContextMenu_CopyAsTsv": "复制为 TSV",
            "GroupBox_DisplayFormat": "显示范围",
            "RadioButton_ShowAll": "显示全部",
            "RadioButton_ShowOnlyDiff": "仅显示差异",
            "ToolTip_Swap": "交换修改前与修改后文件",
            "Label_AddedColor": "新增内容颜色",
            "Label_RemovedColor": "删除内容颜色",
            "Label_ModifiedColor": "修改内容颜色",
            "Label_ColorModifiedRow": "标记修改行",
            "Label_ModifiedRowColor": "修改行背景色",
            "Label_RowColor": "交替行颜色",
            "Label_AddedRows": "新增行（{0}）",
            "Label_RemovedRows": "删除行（{0}）",
            "Label_ModifiedRows": "修改行（{0}）",
            "Label_ModifiedCells": "修改单元格（{0}）",
            "Word_CaseSensitive": "区分大小写",
            "Word_ExactMatch": "完全匹配",
            "Word_RowHeaderIndex": "行标识列序号",
            "Word_ColumnHeaderIndex": "列标题行序号",
            "Word_RowHeaderName": "行标识字段名",
            "Word_StartupSheet": "设为启动时默认工作表",
            "Msg_ExtractingDiff": "正在计算差异……",
            "Msg_ReadingFiles": "正在读取文件……",
            "Msg_WarnSize": "数据量较大，将优先显示差异附近内容。",
            "Msg_Undisplayable": "无法显示",
            "Msg_OutofSheetRange": "文件设置中的工作表序号超出范围。\n将改用第一个工作表。\n不需要应用文件设置时，请启用‘忽略文件设置’。",
            "MenuItem_View": "视图",
            "GroupBox_ReadSetting": "读取设置",
            "GroupBox_Behavior": "操作设置",
            "Label_NotifyEqual": "没有差异时弹出提示",
            "Message_NoDiff": "没有差异。",
            "Label_AlwaysExpandCellDiff": "始终展开单元格差异",
            "Label_FitRowHeight": "自动适应单元格高度",
            "Label_FocusFirstDiff": "比较后自动定位到第一处差异",
        },
    )

    set_default_culture(ROOT / "ExcelMerge.GUI" / "App.xaml.cs")

    replace_required(
        ROOT / "ExcelMerge.GUI" / "Views" / "DiffView.xaml",
        {
            'ContentStringFormat="Added Columns({0})"': 'ContentStringFormat="新增列（{0}）"',
            'ContentStringFormat="Removed Columns({0})"': 'ContentStringFormat="删除列（{0}）"',
            'Header="Edit Dst (编辑右侧表格)"': 'Header="编辑修改后表格"',
            'Content="Save (保存修改)"': 'Content="保存修改"',
            'Content="Discard (放弃修改)"': 'Content="放弃修改"',
            'Header="Copy Row as TSV (复制整行)"': 'Header="复制整行（TSV）"',
            'Header="Insert Row Above (上方插入空行)"': 'Header="在上方插入空行"',
            'Header="Insert Row Below (下方插入空行)"': 'Header="在下方插入空行"',
            'Header="Insert Copied Rows Below (下方插入粘贴复制行)"': 'Header="在下方插入已复制的行"',
            'Header="Delete Row (删除行)"': 'Header="删除行"',
        },
    )

    replace_required(
        ROOT / "ExcelMerge.GUI" / "Views" / "DiffExtractionSettingWindow.xaml",
        {'Content="Smart Table Diff"': 'Content="智能表格对比（按表头和行标识对齐）"'},
    )

    replace_required(
        ROOT / "ExcelMerge.GUI" / "ViewModels" / "DiffViewModel.cs",
        {
            'return "Read only";': 'return "只读";',
            'return HasUnsavedEdits ? "Dst edited" : "Dst editable";':
                'return HasUnsavedEdits ? "修改尚未保存" : "可编辑";',
        },
    )

    replace_required(
        ROOT / "ExcelMerge.GUI" / "Views" / "DiffView.xaml.cs",
        {
            'MessageBox.Show("Saved destination edits.\\nBackup:\\n" + backupPath);':
                'MessageBox.Show("修改已保存。\\n备份文件：\\n" + backupPath);',
            'MessageBox.Show("Failed to save destination edits.\\n\\n" + ex.Message);':
                'MessageBox.Show("保存修改失败。\\n\\n" + ex.Message);',
            'MessageBox.Show("Failed to discard destination edits.\\n\\n" + ex.Message);':
                'MessageBox.Show("放弃修改失败。\\n\\n" + ex.Message);',
            'MessageBox.Show("Save destination edits before closing?", "ExcelMerge", MessageBoxButton.YesNoCancel);':
                'MessageBox.Show("关闭前是否保存对修改后表格所做的更改？", "ExcelMerge", MessageBoxButton.YesNoCancel);',
        },
    )

    replace_required(
        ROOT / "ExcelMerge.GUI" / "Settings" / "ApplicationSetting.cs",
        {'FontName = "Arial";': 'FontName = "Microsoft YaHei UI";'},
    )

    add_chinese_resource_to_project(ROOT / "ExcelMerge.GUI" / "ExcelMerge.GUI.csproj")
    print("Simplified Chinese localization applied successfully.")


if __name__ == "__main__":
    main()
