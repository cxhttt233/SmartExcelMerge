using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge
{
    internal sealed class PreparedRowKeySelection
    {
        public ExcelSheet Source { get; set; }
        public ExcelSheet Destination { get; set; }
        public int? SyntheticSourceColumn { get; set; }
        public int? SyntheticDestinationColumn { get; set; }
    }

    internal static class RowKeySelectionEngine
    {
        private const string SyntheticHeader = "__SMART_EXCEL_MERGE_COMPOSITE_KEY__";
        private const double CoverageThreshold = 0.60;
        private const double UniqueThreshold = 0.95;
        private const double OverlapThreshold = 0.50;

        private sealed class Profile
        {
            public int TotalRows;
            public int NonEmptyRows;
            public Dictionary<string, int> Counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            public HashSet<string> Unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        private sealed class ColumnPair
        {
            public string Name;
            public int SourceIndex;
            public int DestinationIndex;
        }

        internal static PreparedRowKeySelection Prepare(ExcelSheet src, ExcelSheet dst, ExcelSheetDiffConfig config)
        {
            var result = new PreparedRowKeySelection { Source = src, Destination = dst };
            var headers = DetectHeaders(src, dst, config.SrcHeaderIndex, config.DstHeaderIndex);
            config.SrcHeaderIndex = headers.Item1;
            config.DstHeaderIndex = headers.Item2;

            var pairs = GetColumnPairs(src, dst, config.SrcHeaderIndex, config.DstHeaderIndex);
            var manual = RowKeySelectionRuntime.GetManualSelection(config.SrcSheetIndex, config.DstSheetIndex).ToList();
            var analysis = new RowKeyAnalysis
            {
                SelectionMode = manual.Count > 0 ? RowKeySelectionMode.Manual : RowKeySelectionMode.Automatic
            };

            foreach (var pair in pairs)
                analysis.Candidates.Add(Analyze(src, dst, config, new[] { pair }));

            List<ColumnPair> selected = null;
            if (manual.Count > 0)
            {
                selected = new List<ColumnPair>();
                foreach (var name in manual)
                {
                    var pair = pairs.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
                    if (pair == null) { selected = null; break; }
                    selected.Add(pair);
                }
            }
            else
            {
                var best = analysis.Candidates.Where(c => c.IsValidAutomaticKey)
                    .OrderByDescending(c => c.Score).FirstOrDefault();
                if (best != null)
                    selected = pairs.Where(p => best.ColumnNames.Contains(p.Name, StringComparer.OrdinalIgnoreCase)).ToList();
            }

            if (selected == null || selected.Count == 0)
            {
                analysis.SelectionReason = manual.Count > 0
                    ? "手动选择的字段在两表中没有找到，已回退到原有智能匹配。"
                    : "没有找到满足非空率、唯一率和重合率要求的可靠主键，使用原有智能匹配。";
                config.RowKeyAnalysis = analysis;
                RowKeySelectionRuntime.SetAnalysis(config.SrcSheetIndex, config.DstSheetIndex, analysis);
                return result;
            }

            var selectedAnalysis = Analyze(src, dst, config, selected);
            if (manual.Count > 0)
                selectedAnalysis.Reason = selected.Count == 1 ? "手动单主键分析" : "手动联合主键分析";

            analysis.SelectedAnalysis = selectedAnalysis;
            analysis.SelectedColumnNames = selected.Select(p => p.Name).ToList();
            analysis.SelectedScore = selectedAnalysis.Score;
            analysis.SelectedOverlapRate = selectedAnalysis.OverlapRate;
            analysis.MatchedKeyCount = selectedAnalysis.OverlapCount;

            var usable = selectedAnalysis.SourceCoverageRate > 0
                && selectedAnalysis.DestinationCoverageRate > 0
                && selectedAnalysis.OverlapCount > 0
                && selectedAnalysis.SourceUniqueRate >= UniqueThreshold
                && selectedAnalysis.DestinationUniqueRate >= UniqueThreshold;

            if (!usable)
            {
                analysis.SelectedColumnNames = new List<string>();
                analysis.SelectionReason = string.Format(
                    "所选字段分析未通过：左侧非空率 {0:P1}、右侧非空率 {1:P1}、左侧唯一率 {2:P1}、右侧唯一率 {3:P1}、两表重合率 {4:P1}。未强制作为主键。",
                    selectedAnalysis.SourceCoverageRate,
                    selectedAnalysis.DestinationCoverageRate,
                    selectedAnalysis.SourceUniqueRate,
                    selectedAnalysis.DestinationUniqueRate,
                    selectedAnalysis.OverlapRate);
            }
            else if (selected.Count == 1)
            {
                config.SrcRowHeaderName = selected[0].Name;
                config.DstRowHeaderName = selected[0].Name;
                analysis.SelectionReason = manual.Count > 0
                    ? "使用手动单主键匹配记录；已重新计算非空率、唯一率、重合率和匹配数量。"
                    : "自动选择综合得分最高的字段；评分考虑字段名、非空率、唯一率和两表重合率。";
            }
            else
            {
                int srcSynthetic;
                int dstSynthetic;
                result.Source = CloneWithCompositeKey(src, config.SrcHeaderIndex, selected.Select(p => p.SourceIndex).ToList(), out srcSynthetic);
                result.Destination = CloneWithCompositeKey(dst, config.DstHeaderIndex, selected.Select(p => p.DestinationIndex).ToList(), out dstSynthetic);
                result.SyntheticSourceColumn = srcSynthetic;
                result.SyntheticDestinationColumn = dstSynthetic;
                config.SrcRowHeaderName = SyntheticHeader;
                config.DstRowHeaderName = SyntheticHeader;
                analysis.SelectionReason = "使用手动联合主键；已按所选字段组合后重新计算非空率、唯一率、重合率和匹配数量。";
            }

            config.RowKeyAnalysis = analysis;
            RowKeySelectionRuntime.SetAnalysis(config.SrcSheetIndex, config.DstSheetIndex, analysis);
            return result;
        }

        internal static void RemoveSyntheticColumn(ExcelSheetDiff diff, int srcIndex, int dstIndex)
        {
            if (diff == null || diff.Rows.Count == 0) return;
            int? display = null;
            foreach (var row in diff.Rows.Values)
            {
                var match = row.Cells.FirstOrDefault(c =>
                    c.Value.SrcCell != null && c.Value.DstCell != null
                    && c.Value.SrcCell.OriginalColumnIndex == srcIndex
                    && c.Value.DstCell.OriginalColumnIndex == dstIndex);
                if (match.Value != null) { display = match.Key; break; }
            }
            if (!display.HasValue) return;
            diff.Columns.Remove(display.Value);
            foreach (var row in diff.Rows.Values) row.Cells.Remove(display.Value);
        }

        private static Tuple<int, int> DetectHeaders(ExcelSheet src, ExcelSheet dst, int srcConfigured, int dstConfigured)
        {
            int bestSrc = src.Rows.ContainsKey(srcConfigured) ? srcConfigured : (src.Rows.Any() ? src.Rows.Keys.First() : 0);
            int bestDst = dst.Rows.ContainsKey(dstConfigured) ? dstConfigured : (dst.Rows.Any() ? dst.Rows.Keys.First() : 0);
            double best = HeaderScore(GetRow(src, bestSrc), GetRow(dst, bestDst));
            foreach (var sr in src.Rows.Values.Take(20))
                foreach (var dr in dst.Rows.Values.Take(20))
                {
                    var score = HeaderScore(sr, dr);
                    if (score > best) { best = score; bestSrc = sr.Index; bestDst = dr.Index; }
                }
            return Tuple.Create(bestSrc, bestDst);
        }

        private static double HeaderScore(ExcelRow src, ExcelRow dst)
        {
            if (src == null || dst == null) return 0;
            var a = src.Cells.Select(c => Normalize(c.Value)).Where(v => v.Length > 0).ToList();
            var b = dst.Cells.Select(c => Normalize(c.Value)).Where(v => v.Length > 0).ToList();
            var overlap = a.Intersect(b, StringComparer.OrdinalIgnoreCase).Count();
            var keyBonus = a.Concat(b).Any(IsPreferredHeader) ? 50 : 0;
            return keyBonus + overlap * 2 + Math.Min(a.Count, b.Count) * 0.25;
        }

        private static List<ColumnPair> GetColumnPairs(ExcelSheet src, ExcelSheet dst, int srcHeader, int dstHeader)
        {
            var sr = GetRow(src, srcHeader); var dr = GetRow(dst, dstHeader);
            if (sr == null || dr == null) return new List<ColumnPair>();
            var source = UniqueHeaders(sr); var destination = UniqueHeaders(dr);
            return source.Where(x => destination.ContainsKey(x.Key))
                .Select(x => new ColumnPair { Name = x.Key, SourceIndex = x.Value, DestinationIndex = destination[x.Key] })
                .ToList();
        }

        private static Dictionary<string, int> UniqueHeaders(ExcelRow row)
        {
            var groups = row.Cells.Select((c, i) => new { Name = Normalize(c.Value), Index = i })
                .Where(x => x.Name.Length > 0).GroupBy(x => x.Name, StringComparer.OrdinalIgnoreCase);
            return groups.Where(g => g.Count() == 1).ToDictionary(g => g.Key, g => g.First().Index, StringComparer.OrdinalIgnoreCase);
        }

        private static RowKeyCandidateAnalysis Analyze(ExcelSheet src, ExcelSheet dst, ExcelSheetDiffConfig config, IList<ColumnPair> pairs)
        {
            var sp = BuildProfile(src, config.SrcHeaderIndex, pairs.Select(p => p.SourceIndex).ToList());
            var dp = BuildProfile(dst, config.DstHeaderIndex, pairs.Select(p => p.DestinationIndex).ToList());
            var result = new RowKeyCandidateAnalysis { ColumnNames = pairs.Select(p => p.Name).ToList() };
            result.SourceCoverageRate = sp.TotalRows == 0 ? 0 : (double)sp.NonEmptyRows / sp.TotalRows;
            result.DestinationCoverageRate = dp.TotalRows == 0 ? 0 : (double)dp.NonEmptyRows / dp.TotalRows;
            result.SourceUniqueRate = sp.NonEmptyRows == 0 ? 0 : (double)sp.Unique.Count / sp.NonEmptyRows;
            result.DestinationUniqueRate = dp.NonEmptyRows == 0 ? 0 : (double)dp.Unique.Count / dp.NonEmptyRows;
            result.OverlapCount = sp.Unique.Intersect(dp.Unique, StringComparer.OrdinalIgnoreCase).Count();
            var min = Math.Min(sp.Unique.Count, dp.Unique.Count);
            result.OverlapRate = min == 0 ? 0 : (double)result.OverlapCount / min;
            result.IsPreferredHeader = pairs.Count == 1 && IsPreferredHeader(pairs[0].Name);
            result.Score = (result.IsPreferredHeader ? 4 : 0) + result.OverlapRate * 4
                + result.SourceCoverageRate + result.DestinationCoverageRate
                + result.SourceUniqueRate + result.DestinationUniqueRate
                + Math.Min(result.OverlapCount, 100) / 100.0;
            result.IsValidAutomaticKey = result.SourceCoverageRate >= CoverageThreshold
                && result.DestinationCoverageRate >= CoverageThreshold
                && result.SourceUniqueRate >= UniqueThreshold
                && result.DestinationUniqueRate >= UniqueThreshold
                && result.OverlapRate >= OverlapThreshold && min >= 2;
            if (result.SourceCoverageRate < CoverageThreshold || result.DestinationCoverageRate < CoverageThreshold) result.Reason = "非空率不足";
            else if (result.SourceUniqueRate < UniqueThreshold || result.DestinationUniqueRate < UniqueThreshold) result.Reason = "唯一率不足";
            else if (result.OverlapRate < OverlapThreshold || min < 2) result.Reason = "两表重合率不足";
            else result.Reason = result.IsPreferredHeader ? "有效候选，字段名符合ID/编码特征" : "有效候选";
            return result;
        }

        private static Profile BuildProfile(ExcelSheet sheet, int header, IList<int> columns)
        {
            var profile = new Profile();
            foreach (var row in sheet.Rows.Values.Where(r => r.Index > header && !r.IsBlank()))
            {
                profile.TotalRows++;
                var key = CompositeValue(row, columns);
                if (key.Length == 0) continue;
                profile.NonEmptyRows++;
                int count; profile.Counts.TryGetValue(key, out count); profile.Counts[key] = count + 1;
            }
            foreach (var item in profile.Counts.Where(x => x.Value == 1)) profile.Unique.Add(item.Key);
            return profile;
        }

        private static ExcelSheet CloneWithCompositeKey(ExcelSheet sheet, int header, IList<int> columns, out int syntheticIndex)
        {
            syntheticIndex = sheet.Rows.Any() ? sheet.Rows.Max(r => r.Value.Cells.Count) : 0;
            var clone = new ExcelSheet();
            foreach (var row in sheet.Rows.Values)
            {
                var cells = row.Cells.Select(c => new ExcelCell(c.Value, c.OriginalColumnIndex, c.OriginalRowIndex)).ToList();
                while (cells.Count < syntheticIndex) cells.Add(new ExcelCell(string.Empty, cells.Count, row.Index));
                var value = row.Index == header ? SyntheticHeader : (row.Index > header ? CompositeValue(row, columns) : string.Empty);
                cells.Add(new ExcelCell(value, syntheticIndex, row.Index));
                clone.Rows.Add(row.Index, new ExcelRow(row.Index, cells));
            }
            return clone;
        }

        private static string CompositeValue(ExcelRow row, IList<int> columns)
        {
            var values = new List<string>();
            foreach (var index in columns)
            {
                if (index < 0 || index >= row.Cells.Count) return string.Empty;
                var value = (row.Cells[index].Value ?? string.Empty).Trim();
                if (value.Length == 0) return string.Empty;
                values.Add(value);
            }
            return string.Join("\u001f", values);
        }

        private static ExcelRow GetRow(ExcelSheet sheet, int index) { ExcelRow row; return sheet.Rows.TryGetValue(index, out row) ? row : null; }
        private static string Normalize(string value) { return (value ?? string.Empty).Trim(); }
        private static bool IsPreferredHeader(string value)
        {
            var text = Normalize(value).ToLowerInvariant();
            return text == "id" || text == "key" || text.EndsWith("id") || text.Contains("code")
                || text.Contains("编码") || text.Contains("编号") || text.Contains("代码") || text.Contains("主键");
        }
    }
}
