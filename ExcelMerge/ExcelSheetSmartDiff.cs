using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetDiff;

namespace ExcelMerge
{
    public partial class ExcelSheet
    {
        private const int SmartSimilarityPairLimit = 40000;
        private const long SmartRowLcsPairLimit = 2000000;
        private const int SmartHeaderScanRowLimit = 20;
        private const double SmartColumnSimilarityThreshold = 0.80;
        private const double SmartRowSimilarityThreshold = 0.80;
        private const double SmartKeyOverlapThreshold = 0.50;
        private const double SmartKeyCoverageThreshold = 0.60;

        private class SmartColumn
        {
            public int DisplayIndex { get; set; }
            public int? SrcIndex { get; set; }
            public int? DstIndex { get; set; }
            public ExcelColumnStatus Status { get; set; }
        }

        private class SmartRow
        {
            public ExcelRow SrcRow { get; private set; }
            public ExcelRow DstRow { get; private set; }
            public ExcelRowStatus Status { get; private set; }

            public SmartRow(ExcelRow srcRow, ExcelRow dstRow, ExcelRowStatus status)
            {
                SrcRow = srcRow;
                DstRow = dstRow;
                Status = status;
            }
        }

        private class SmartRowKey
        {
            public int SrcColumnIndex { get; private set; }
            public int DstColumnIndex { get; private set; }
            public double Score { get; private set; }

            public SmartRowKey(int srcColumnIndex, int dstColumnIndex, double score)
            {
                SrcColumnIndex = srcColumnIndex;
                DstColumnIndex = dstColumnIndex;
                Score = score;
            }
        }

        private class SmartHeaderContext
        {
            public int SrcHeaderIndex { get; private set; }
            public int DstHeaderIndex { get; private set; }
            public double Score { get; private set; }

            public SmartHeaderContext(int srcHeaderIndex, int dstHeaderIndex, double score)
            {
                SrcHeaderIndex = srcHeaderIndex;
                DstHeaderIndex = dstHeaderIndex;
                Score = score;
            }
        }

        private class RowKeyProfile
        {
            public int TotalRows { get; set; }
            public int NonEmptyCount { get; set; }
            public Dictionary<string, ExcelRow> UniqueRows { get; private set; }
            public Dictionary<string, int> KeyCounts { get; private set; }

            public RowKeyProfile()
            {
                UniqueRows = new Dictionary<string, ExcelRow>(StringComparer.OrdinalIgnoreCase);
                KeyCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            }
        }

        private class SmartRowSignature : IEquatable<SmartRowSignature>
        {
            public ExcelRow Row { get; private set; }
            public string Signature { get; private set; }

            public SmartRowSignature(ExcelRow row, string signature)
            {
                Row = row;
                Signature = signature;
            }

            public bool Equals(SmartRowSignature other)
            {
                return other != null && Signature == other.Signature;
            }

            public override bool Equals(object obj)
            {
                return Equals(obj as SmartRowSignature);
            }

            public override int GetHashCode()
            {
                return Signature.GetHashCode();
            }
        }

        private class SimilarityPair
        {
            public int SrcIndex { get; private set; }
            public int DstIndex { get; private set; }
            public double Score { get; private set; }

            public SimilarityPair(int srcIndex, int dstIndex, double score)
            {
                SrcIndex = srcIndex;
                DstIndex = dstIndex;
                Score = score;
            }
        }

        private class SmartRowMatch
        {
            public ExcelRow SrcRow { get; private set; }
            public ExcelRow DstRow { get; private set; }

            public SmartRowMatch(ExcelRow srcRow, ExcelRow dstRow)
            {
                SrcRow = srcRow;
                DstRow = dstRow;
            }
        }

        private static ExcelSheetDiff SmartDiff(ExcelSheet src, ExcelSheet dst, ExcelSheetDiffConfig config)
        {
            var headerContext = CreateSmartHeaderContext(src, dst, config);
            config.SrcHeaderIndex = headerContext.SrcHeaderIndex;
            config.DstHeaderIndex = headerContext.DstHeaderIndex;

            bool hasColumnAnchor;
            var columns = CreateSmartColumns(src, dst, config, out hasColumnAnchor);
            var rowKey = FindSmartRowKey(src, dst, columns, config);

            if (rowKey == null && !hasColumnAnchor && ShouldFallbackRowsToLegacy(src, dst))
                return DiffLegacy(src, dst, config);

            if (rowKey == null && ShouldFallbackRowsToLegacy(src, dst))
                return DiffLegacy(src, dst, config);

            var rows = rowKey != null
                ? CreateSmartRowsByKey(src, dst, rowKey, columns, config)
                : CreateSmartRowsByLcs(src, dst, columns);

            return CreateSmartSheetDiff(rows, columns);
        }

        private static SmartHeaderContext CreateSmartHeaderContext(ExcelSheet src, ExcelSheet dst, ExcelSheetDiffConfig config)
        {
            var configuredSrcHeaderIndex = NormalizeHeaderIndex(src, config.SrcHeaderIndex);
            var configuredDstHeaderIndex = NormalizeHeaderIndex(dst, config.DstHeaderIndex);
            var configuredScore = ScoreHeaderPair(
                GetRowOrNull(src, configuredSrcHeaderIndex),
                GetRowOrNull(dst, configuredDstHeaderIndex));
            var configured = new SmartHeaderContext(configuredSrcHeaderIndex, configuredDstHeaderIndex, configuredScore);
            var best = configured;
            var srcRows = src.Rows.Values.Take(SmartHeaderScanRowLimit).ToList();
            var dstRows = dst.Rows.Values.Take(SmartHeaderScanRowLimit).ToList();

            foreach (var srcRow in srcRows)
            {
                foreach (var dstRow in dstRows)
                {
                    var score = ScoreHeaderPair(srcRow, dstRow);
                    if (score > best.Score)
                        best = new SmartHeaderContext(srcRow.Index, dstRow.Index, score);
                }
            }

            return best.Score >= 20 && best.Score > configured.Score + 5
                ? best
                : configured;
        }

        private static int NormalizeHeaderIndex(ExcelSheet sheet, int headerIndex)
        {
            if (sheet.Rows.ContainsKey(headerIndex))
                return headerIndex;

            return sheet.Rows.Any() ? sheet.Rows.Keys.First() : 0;
        }

        private static ExcelRow GetRowOrNull(ExcelSheet sheet, int rowIndex)
        {
            ExcelRow row;
            return sheet.Rows.TryGetValue(rowIndex, out row) ? row : null;
        }

        private static double ScoreHeaderPair(ExcelRow srcRow, ExcelRow dstRow)
        {
            if (srcRow == null || dstRow == null)
                return 0;

            var srcHeaders = GetHeaderValues(srcRow).ToList();
            var dstHeaders = GetHeaderValues(dstRow).ToList();
            if (!srcHeaders.Any() || !dstHeaders.Any())
                return 0;

            var srcUnique = new HashSet<string>(srcHeaders, StringComparer.OrdinalIgnoreCase);
            var dstUnique = new HashSet<string>(dstHeaders, StringComparer.OrdinalIgnoreCase);
            var overlapCount = srcUnique.Intersect(dstUnique, StringComparer.OrdinalIgnoreCase).Count();
            var srcFirst = NormalizeHeader(GetCellValue(srcRow, 0));
            var dstFirst = NormalizeHeader(GetCellValue(dstRow, 0));
            var markerBonus = IsFieldHeaderMarker(srcFirst) || IsFieldHeaderMarker(dstFirst) ? 100.0 : 0.0;
            var keyBonus = srcUnique.Any(IsPreferredKeyHeader) || dstUnique.Any(IsPreferredKeyHeader) ? 50.0 : 0.0;
            var typePenalty = IsTypeHeaderRow(srcRow) || IsTypeHeaderRow(dstRow) ? 40.0 : 0.0;
            var metadataPenalty = IsMetadataHeaderRow(srcFirst) || IsMetadataHeaderRow(dstFirst) ? 20.0 : 0.0;

            return markerBonus + keyBonus + overlapCount * 2.0 + Math.Min(srcUnique.Count, dstUnique.Count) * 0.25 - typePenalty - metadataPenalty;
        }

        private static IEnumerable<string> GetHeaderValues(ExcelRow row)
        {
            return row.Cells
                .Select(c => NormalizeHeader(c.Value))
                .Where(v => !string.IsNullOrEmpty(v));
        }

        private static ExcelSheetDiff CreateSmartSheetDiff(IEnumerable<SmartRow> rows, IList<SmartColumn> columns)
        {
            var sheetDiff = new ExcelSheetDiff();
            foreach (var column in columns)
                sheetDiff.Columns[column.DisplayIndex] = column.Status;

            foreach (var smartRow in rows)
            {
                var rowStatus = smartRow.Status == ExcelRowStatus.None
                    ? (ExcelRowStatus?)ExcelRowStatus.None
                    : smartRow.Status;
                var row = sheetDiff.CreateRow(rowStatus);
                var hasModifiedCell = false;

                foreach (var column in columns)
                {
                    var srcCell = GetSmartCell(smartRow.SrcRow, column.SrcIndex, smartRow.DstRow, column.DstIndex, column.DisplayIndex);
                    var dstCell = GetSmartCell(smartRow.DstRow, column.DstIndex, smartRow.SrcRow, column.SrcIndex, column.DisplayIndex);
                    var status = GetSmartCellStatus(smartRow, column, srcCell, dstCell);

                    if (status == ExcelCellStatus.Modified)
                        hasModifiedCell = true;

                    row.CreateCell(srcCell, dstCell, column.DisplayIndex, status);
                }

                if (smartRow.Status == ExcelRowStatus.None && hasModifiedCell)
                    row.SetStatus(ExcelRowStatus.Modified);
            }

            return sheetDiff;
        }

        private static ExcelCellStatus GetSmartCellStatus(SmartRow row, SmartColumn column, ExcelCell srcCell, ExcelCell dstCell)
        {
            if (row.Status == ExcelRowStatus.Added)
                return ExcelCellStatus.Added;

            if (row.Status == ExcelRowStatus.Removed)
                return ExcelCellStatus.Removed;

            if (column.Status == ExcelColumnStatus.Inserted)
                return ExcelCellStatus.Added;

            if (column.Status == ExcelColumnStatus.Deleted)
                return ExcelCellStatus.Removed;

            return srcCell.Value == dstCell.Value ? ExcelCellStatus.None : ExcelCellStatus.Modified;
        }

        private static ExcelCell GetSmartCell(ExcelRow row, int? columnIndex, ExcelRow fallbackRow, int? fallbackColumnIndex, int displayColumnIndex)
        {
            if (row != null && columnIndex.HasValue && columnIndex.Value >= 0 && columnIndex.Value < row.Cells.Count)
                return row.Cells[columnIndex.Value];

            var originalRowIndex = row != null
                ? row.Index
                : fallbackRow != null ? fallbackRow.Index : 0;
            var originalColumnIndex = columnIndex.HasValue
                ? columnIndex.Value
                : fallbackColumnIndex.HasValue ? fallbackColumnIndex.Value : displayColumnIndex;

            return new ExcelCell(string.Empty, originalColumnIndex, originalRowIndex);
        }

        private static List<SmartColumn> CreateSmartColumns(ExcelSheet src, ExcelSheet dst, ExcelSheetDiffConfig config, out bool hasColumnAnchor)
        {
            var srcColumns = src.CreateColumns().ToList();
            var dstColumns = dst.CreateColumns().ToList();
            var srcCount = srcColumns.Count;
            var dstCount = dstColumns.Count;
            var srcToDst = new Dictionary<int, int>();
            var dstToSrc = new Dictionary<int, int>();
            hasColumnAnchor = false;

            MatchColumnsByUniqueHeader(src, dst, config, srcCount, dstCount, srcToDst, dstToSrc);
            MatchColumnsByAnchoredPosition(srcCount, dstCount, srcToDst, dstToSrc);
            if (srcToDst.Any())
                hasColumnAnchor = true;

            var unmatchedSrc = Enumerable.Range(0, srcCount).Where(i => !srcToDst.ContainsKey(i)).ToList();
            var unmatchedDst = Enumerable.Range(0, dstCount).Where(i => !dstToSrc.ContainsKey(i)).ToList();
            if ((long)unmatchedSrc.Count * unmatchedDst.Count <= SmartSimilarityPairLimit)
            {
                var addedBySimilarity = MatchColumnsBySimilarity(srcColumns, dstColumns, config, unmatchedSrc, unmatchedDst, srcToDst, dstToSrc);
                hasColumnAnchor |= addedBySimilarity;
            }
            else if (!hasColumnAnchor)
            {
                return CreateSmartColumnsFromLegacy(srcColumns, dstColumns, config);
            }

            return CreateSmartColumnsFromMatches(srcCount, dstCount, srcToDst, dstToSrc);
        }

        private static void MatchColumnsByUniqueHeader(
            ExcelSheet src,
            ExcelSheet dst,
            ExcelSheetDiffConfig config,
            int srcCount,
            int dstCount,
            Dictionary<int, int> srcToDst,
            Dictionary<int, int> dstToSrc)
        {
            var srcHeaders = BuildUniqueHeaderMap(src, srcCount, config.SrcHeaderIndex);
            var dstHeaders = BuildUniqueHeaderMap(dst, dstCount, config.DstHeaderIndex);

            foreach (var srcHeader in srcHeaders)
            {
                int dstIndex;
                if (!dstHeaders.TryGetValue(srcHeader.Key, out dstIndex))
                    continue;

                srcToDst[srcHeader.Value] = dstIndex;
                dstToSrc[dstIndex] = srcHeader.Value;
            }
        }

        private static Dictionary<string, int> BuildUniqueHeaderMap(ExcelSheet sheet, int columnCount, int headerRowIndex)
        {
            var occurrences = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                var header = NormalizeHeader(GetCellValue(sheet.Rows.ContainsKey(headerRowIndex) ? sheet.Rows[headerRowIndex] : null, columnIndex));
                if (string.IsNullOrEmpty(header))
                    continue;

                List<int> indices;
                if (!occurrences.TryGetValue(header, out indices))
                {
                    indices = new List<int>();
                    occurrences[header] = indices;
                }

                indices.Add(columnIndex);
            }

            return occurrences
                .Where(i => i.Value.Count == 1)
                .ToDictionary(i => i.Key, i => i.Value[0], StringComparer.OrdinalIgnoreCase);
        }

        private static void MatchColumnsByAnchoredPosition(
            int srcCount,
            int dstCount,
            Dictionary<int, int> srcToDst,
            Dictionary<int, int> dstToSrc)
        {
            if (!srcToDst.Any())
                return;

            var anchors = srcToDst
                .Select(i => new { SrcIndex = i.Key, DstIndex = i.Value })
                .OrderBy(i => i.SrcIndex)
                .ToList();
            anchors.Insert(0, new { SrcIndex = -1, DstIndex = -1 });
            anchors.Add(new { SrcIndex = srcCount, DstIndex = dstCount });

            for (var i = 0; i < anchors.Count - 1; i++)
            {
                var left = anchors[i];
                var right = anchors[i + 1];
                if (left.DstIndex >= right.DstIndex)
                    continue;

                var srcStart = left.SrcIndex + 1;
                var srcEnd = right.SrcIndex - 1;
                var dstStart = left.DstIndex + 1;
                var dstEnd = right.DstIndex - 1;
                var srcGapCount = srcEnd - srcStart + 1;
                var dstGapCount = dstEnd - dstStart + 1;
                if (srcGapCount <= 0 || srcGapCount != dstGapCount)
                    continue;

                for (var offset = 0; offset < srcGapCount; offset++)
                {
                    var srcIndex = srcStart + offset;
                    var dstIndex = dstStart + offset;
                    if (srcToDst.ContainsKey(srcIndex) || dstToSrc.ContainsKey(dstIndex))
                        continue;

                    srcToDst[srcIndex] = dstIndex;
                    dstToSrc[dstIndex] = srcIndex;
                }
            }
        }

        private static bool MatchColumnsBySimilarity(
            IList<ExcelColumn> srcColumns,
            IList<ExcelColumn> dstColumns,
            ExcelSheetDiffConfig config,
            IList<int> unmatchedSrc,
            IList<int> unmatchedDst,
            Dictionary<int, int> srcToDst,
            Dictionary<int, int> dstToSrc)
        {
            var candidates = new List<SimilarityPair>();
            foreach (var srcIndex in unmatchedSrc)
            {
                foreach (var dstIndex in unmatchedDst)
                {
                    var score = CalculateColumnSimilarity(srcColumns[srcIndex], dstColumns[dstIndex], config.SrcHeaderIndex, config.DstHeaderIndex);
                    if (score >= SmartColumnSimilarityThreshold)
                        candidates.Add(new SimilarityPair(srcIndex, dstIndex, score));
                }
            }

            var matched = false;
            foreach (var candidate in candidates.OrderByDescending(c => c.Score))
            {
                if (srcToDst.ContainsKey(candidate.SrcIndex) || dstToSrc.ContainsKey(candidate.DstIndex))
                    continue;

                srcToDst[candidate.SrcIndex] = candidate.DstIndex;
                dstToSrc[candidate.DstIndex] = candidate.SrcIndex;
                matched = true;
            }

            return matched;
        }

        private static List<SmartColumn> CreateSmartColumnsFromLegacy(
            IEnumerable<ExcelColumn> srcColumns,
            IEnumerable<ExcelColumn> dstColumns,
            ExcelSheetDiffConfig config)
        {
            var legacyStatusMap = CreateColumnStatusMap(srcColumns, dstColumns, config);
            var columns = new List<SmartColumn>();
            var srcIndex = 0;
            var dstIndex = 0;
            foreach (var legacyStatus in legacyStatusMap)
            {
                var column = new SmartColumn
                {
                    DisplayIndex = columns.Count,
                    Status = legacyStatus.Value,
                };

                if (legacyStatus.Value == ExcelColumnStatus.Deleted)
                {
                    column.SrcIndex = srcIndex++;
                }
                else if (legacyStatus.Value == ExcelColumnStatus.Inserted)
                {
                    column.DstIndex = dstIndex++;
                }
                else
                {
                    column.SrcIndex = srcIndex++;
                    column.DstIndex = dstIndex++;
                }

                columns.Add(column);
            }

            return columns;
        }

        private static List<SmartColumn> CreateSmartColumnsFromMatches(
            int srcCount,
            int dstCount,
            Dictionary<int, int> srcToDst,
            Dictionary<int, int> dstToSrc)
        {
            var columns = new List<SmartColumn>();
            var emittedSrcColumns = new HashSet<int>();

            for (int dstIndex = 0; dstIndex < dstCount; dstIndex++)
            {
                int matchedSrcIndex;
                if (dstToSrc.TryGetValue(dstIndex, out matchedSrcIndex))
                {
                    AddDeletedColumnsBefore(columns, srcCount, matchedSrcIndex, emittedSrcColumns, srcToDst);
                    columns.Add(new SmartColumn
                    {
                        DisplayIndex = columns.Count,
                        SrcIndex = matchedSrcIndex,
                        DstIndex = dstIndex,
                        Status = ExcelColumnStatus.None,
                    });
                    emittedSrcColumns.Add(matchedSrcIndex);
                }
                else
                {
                    columns.Add(new SmartColumn
                    {
                        DisplayIndex = columns.Count,
                        DstIndex = dstIndex,
                        Status = ExcelColumnStatus.Inserted,
                    });
                }
            }

            AddDeletedColumnsBefore(columns, srcCount, int.MaxValue, emittedSrcColumns, srcToDst);

            return columns;
        }

        private static void AddDeletedColumnsBefore(
            IList<SmartColumn> columns,
            int srcCount,
            int beforeSrcIndex,
            HashSet<int> emittedSrcColumns,
            Dictionary<int, int> srcToDst)
        {
            for (int srcIndex = 0; srcIndex < srcCount && srcIndex < beforeSrcIndex; srcIndex++)
            {
                if (emittedSrcColumns.Contains(srcIndex) || srcToDst.ContainsKey(srcIndex))
                    continue;

                columns.Add(new SmartColumn
                {
                    DisplayIndex = columns.Count,
                    SrcIndex = srcIndex,
                    Status = ExcelColumnStatus.Deleted,
                });
                emittedSrcColumns.Add(srcIndex);
            }
        }

        private static SmartRowKey FindSmartRowKey(ExcelSheet src, ExcelSheet dst, IList<SmartColumn> columns, ExcelSheetDiffConfig config)
        {
            var configuredKey = FindConfiguredRowKey(src, dst, columns, config);
            if (configuredKey != null)
                return configuredKey;

            SmartRowKey best = null;
            foreach (var column in columns.Where(c => c.SrcIndex.HasValue && c.DstIndex.HasValue))
            {
                var score = ScoreRowKey(src, dst, column.SrcIndex.Value, column.DstIndex.Value, config, false);
                if (score < 0)
                    continue;

                if (best == null || score > best.Score)
                    best = new SmartRowKey(column.SrcIndex.Value, column.DstIndex.Value, score);
            }

            return best;
        }

        private static SmartRowKey FindConfiguredRowKey(ExcelSheet src, ExcelSheet dst, IList<SmartColumn> columns, ExcelSheetDiffConfig config)
        {
            var srcIndex = ResolveConfiguredRowHeaderIndex(src, config.SrcHeaderIndex, config.SrcRowHeaderIndex, config.SrcRowHeaderName);
            var dstIndex = ResolveConfiguredRowHeaderIndex(dst, config.DstHeaderIndex, config.DstRowHeaderIndex, config.DstRowHeaderName);

            if (srcIndex >= 0 && dstIndex < 0)
            {
                var matchedColumn = columns.FirstOrDefault(c => c.SrcIndex == srcIndex && c.DstIndex.HasValue);
                if (matchedColumn != null)
                    dstIndex = matchedColumn.DstIndex.Value;
            }

            if (dstIndex >= 0 && srcIndex < 0)
            {
                var matchedColumn = columns.FirstOrDefault(c => c.DstIndex == dstIndex && c.SrcIndex.HasValue);
                if (matchedColumn != null)
                    srcIndex = matchedColumn.SrcIndex.Value;
            }

            if (srcIndex < 0 || dstIndex < 0)
                return null;

            var score = ScoreRowKey(src, dst, srcIndex, dstIndex, config, true);
            return score >= 0 ? new SmartRowKey(srcIndex, dstIndex, score) : null;
        }

        private static int ResolveConfiguredRowHeaderIndex(ExcelSheet sheet, int columnHeaderRowIndex, int rowHeaderIndex, string rowHeaderName)
        {
            if (rowHeaderIndex >= 0)
                return rowHeaderIndex;

            if (!string.IsNullOrWhiteSpace(rowHeaderName))
                return FindColumnByHeaderName(sheet, columnHeaderRowIndex, rowHeaderName);

            return -1;
        }

        private static int FindColumnByHeaderName(ExcelSheet sheet, int columnHeaderRowIndex, string rowHeaderName)
        {
            ExcelRow headerRow;
            if (!sheet.Rows.TryGetValue(columnHeaderRowIndex, out headerRow))
                return -1;

            var normalizedTarget = NormalizeHeader(rowHeaderName);
            for (int i = 0; i < headerRow.Cells.Count; i++)
            {
                if (NormalizeHeader(headerRow.Cells[i].Value) == normalizedTarget)
                    return i;
            }

            return -1;
        }

        private static double ScoreRowKey(ExcelSheet src, ExcelSheet dst, int srcColumnIndex, int dstColumnIndex, ExcelSheetDiffConfig config, bool isConfigured)
        {
            var skippedSrcRows = CreateSkippedHeaderRows(src, config.SrcHeaderIndex);
            var skippedDstRows = CreateSkippedHeaderRows(dst, config.DstHeaderIndex);
            var srcProfile = BuildRowKeyProfile(src, srcColumnIndex, skippedSrcRows);
            var dstProfile = BuildRowKeyProfile(dst, dstColumnIndex, skippedDstRows);

            if (srcProfile.NonEmptyCount == 0 || dstProfile.NonEmptyCount == 0)
                return -1;

            var srcCoverageRate = srcProfile.TotalRows == 0 ? 0 : (double)srcProfile.NonEmptyCount / srcProfile.TotalRows;
            var dstCoverageRate = dstProfile.TotalRows == 0 ? 0 : (double)dstProfile.NonEmptyCount / dstProfile.TotalRows;
            if (!isConfigured && (srcCoverageRate < SmartKeyCoverageThreshold || dstCoverageRate < SmartKeyCoverageThreshold))
                return -1;

            var srcUniqueRate = (double)srcProfile.UniqueRows.Count / srcProfile.NonEmptyCount;
            var dstUniqueRate = (double)dstProfile.UniqueRows.Count / dstProfile.NonEmptyCount;
            if (srcUniqueRate < 0.95 || dstUniqueRate < 0.95)
                return -1;

            var overlapCount = srcProfile.UniqueRows.Keys.Intersect(dstProfile.UniqueRows.Keys, StringComparer.OrdinalIgnoreCase).Count();
            var minUniqueCount = Math.Min(srcProfile.UniqueRows.Count, dstProfile.UniqueRows.Count);
            if (minUniqueCount == 0)
                return -1;

            var overlapRate = (double)overlapCount / minUniqueCount;
            if (isConfigured)
            {
                if (overlapCount == 0)
                    return -1;
            }
            else if (overlapRate < SmartKeyOverlapThreshold || minUniqueCount < 2)
            {
                return -1;
            }

            var srcHeader = GetHeaderText(src, config.SrcHeaderIndex, srcColumnIndex);
            var dstHeader = GetHeaderText(dst, config.DstHeaderIndex, dstColumnIndex);
            var preferredHeaderBonus = IsPreferredKeyHeader(srcHeader) || IsPreferredKeyHeader(dstHeader) ? 4.0 : 0.0;
            var overlapCountBonus = Math.Min(overlapCount, 100) / 100.0;

            return preferredHeaderBonus + overlapRate * 4.0 + srcUniqueRate + dstUniqueRate + srcCoverageRate + dstCoverageRate + overlapCountBonus;
        }

        private static HashSet<int> CreateSkippedHeaderRows(ExcelSheet sheet, int headerRowIndex)
        {
            var skippedRows = new HashSet<int>();
            foreach (var row in sheet.Rows.Values)
            {
                if (row.Index <= headerRowIndex)
                    skippedRows.Add(row.Index);
            }

            var nextRowIndex = headerRowIndex + 1;
            while (true)
            {
                ExcelRow row;
                if (!sheet.Rows.TryGetValue(nextRowIndex, out row))
                    break;

                if (!row.IsBlank() && !IsDefaultMetadataRow(row))
                    break;

                skippedRows.Add(row.Index);
                nextRowIndex++;
            }

            return skippedRows;
        }

        private static RowKeyProfile BuildRowKeyProfile(ExcelSheet sheet, int columnIndex, HashSet<int> skippedRows)
        {
            var profile = new RowKeyProfile();
            var seen = new Dictionary<string, ExcelRow>(StringComparer.OrdinalIgnoreCase);
            var duplicates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var row in sheet.Rows.Values)
            {
                if (skippedRows.Contains(row.Index))
                    continue;

                if (row.IsBlank())
                    continue;

                profile.TotalRows++;
                var key = NormalizeKey(GetCellValue(row, columnIndex));
                if (string.IsNullOrEmpty(key))
                    continue;

                profile.NonEmptyCount++;
                int count;
                profile.KeyCounts.TryGetValue(key, out count);
                profile.KeyCounts[key] = count + 1;
                if (seen.ContainsKey(key))
                    duplicates.Add(key);
                else
                    seen[key] = row;
            }

            foreach (var item in seen)
            {
                if (!duplicates.Contains(item.Key))
                    profile.UniqueRows[item.Key] = item.Value;
            }

            return profile;
        }

        private static IEnumerable<SmartRow> CreateSmartRowsByKey(
            ExcelSheet src,
            ExcelSheet dst,
            SmartRowKey rowKey,
            IList<SmartColumn> columns,
            ExcelSheetDiffConfig config)
        {
            var rows = new List<SmartRow>();
            var pairedSrcRows = new HashSet<int>();
            var pairedDstRows = new HashSet<int>();

            PairPrefixRows(src, dst, config, rows, pairedSrcRows, pairedDstRows);

            var srcProfile = BuildRowKeyProfile(src, rowKey.SrcColumnIndex, pairedSrcRows);
            var dstProfile = BuildRowKeyProfile(dst, rowKey.DstColumnIndex, pairedDstRows);
            var dstRows = dst.Rows.Values.Where(r => !pairedDstRows.Contains(r.Index)).ToList();
            var matches = new List<SmartRowMatch>();

            foreach (var dstRow in dstRows)
            {
                var key = NormalizeKey(GetCellValue(dstRow, rowKey.DstColumnIndex));
                ExcelRow srcRow;
                if (!string.IsNullOrEmpty(key)
                    && GetKeyCount(srcProfile, key) == 1
                    && GetKeyCount(dstProfile, key) == 1
                    && srcProfile.UniqueRows.TryGetValue(key, out srcRow)
                    && !pairedSrcRows.Contains(srcRow.Index))
                {
                    matches.Add(new SmartRowMatch(srcRow, dstRow));
                    pairedSrcRows.Add(srcRow.Index);
                    pairedDstRows.Add(dstRow.Index);
                }
            }

            PairRowsWithoutReliableKey(
                src,
                dst,
                rowKey,
                srcProfile,
                dstProfile,
                columns,
                matches,
                pairedSrcRows,
                pairedDstRows);

            var addedDstRows = dst.Rows.Values.Where(r => !pairedDstRows.Contains(r.Index)).ToList();
            AddDataRowsInSourceOrder(src, rows, pairedSrcRows, matches, addedDstRows);
            return rows;
        }

        private static void PairRowsWithoutReliableKey(
            ExcelSheet src,
            ExcelSheet dst,
            SmartRowKey rowKey,
            RowKeyProfile srcProfile,
            RowKeyProfile dstProfile,
            IList<SmartColumn> columns,
            IList<SmartRowMatch> matches,
            HashSet<int> pairedSrcRows,
            HashSet<int> pairedDstRows)
        {
            var srcRowsByKey = src.Rows.Values
                .Where(r => !pairedSrcRows.Contains(r.Index)
                    && HasUnreliableKey(r, rowKey.SrcColumnIndex, srcProfile, dstProfile))
                .GroupBy(r => NormalizeKey(GetCellValue(r, rowKey.SrcColumnIndex)), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);
            var dstRowsByKey = dst.Rows.Values
                .Where(r => !pairedDstRows.Contains(r.Index)
                    && HasUnreliableKey(r, rowKey.DstColumnIndex, dstProfile, srcProfile))
                .GroupBy(r => NormalizeKey(GetCellValue(r, rowKey.DstColumnIndex)), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

            foreach (var key in srcRowsByKey.Keys.Union(dstRowsByKey.Keys, StringComparer.OrdinalIgnoreCase))
            {
                List<ExcelRow> srcRows;
                if (!srcRowsByKey.TryGetValue(key, out srcRows))
                    srcRows = new List<ExcelRow>();

                List<ExcelRow> dstRows;
                if (!dstRowsByKey.TryGetValue(key, out dstRows))
                    dstRows = new List<ExcelRow>();

                PairRowsWithinKeyGroup(srcRows, dstRows, columns, matches, pairedSrcRows, pairedDstRows);
            }
        }

        private static void PairRowsWithinKeyGroup(
            IList<ExcelRow> srcRows,
            IList<ExcelRow> dstRows,
            IList<SmartColumn> columns,
            IList<SmartRowMatch> matches,
            HashSet<int> pairedSrcRows,
            HashSet<int> pairedDstRows)
        {
            PairExactRows(srcRows, dstRows, columns, matches, pairedSrcRows, pairedDstRows);

            srcRows = srcRows.Where(r => !pairedSrcRows.Contains(r.Index)).ToList();
            dstRows = dstRows.Where(r => !pairedDstRows.Contains(r.Index)).ToList();
            if ((long)srcRows.Count * dstRows.Count > SmartRowLcsPairLimit)
                return;

            foreach (var row in CreateSmartRowsByLcs(srcRows, dstRows, columns))
            {
                if (row.SrcRow == null || row.DstRow == null)
                    continue;

                matches.Add(new SmartRowMatch(row.SrcRow, row.DstRow));
                pairedSrcRows.Add(row.SrcRow.Index);
                pairedDstRows.Add(row.DstRow.Index);
            }
        }

        private static bool HasUnreliableKey(
            ExcelRow row,
            int columnIndex,
            RowKeyProfile profile,
            RowKeyProfile otherProfile)
        {
            var key = NormalizeKey(GetCellValue(row, columnIndex));
            if (string.IsNullOrEmpty(key))
                return true;

            return GetKeyCount(profile, key) > 1 || GetKeyCount(otherProfile, key) > 1;
        }

        private static int GetKeyCount(RowKeyProfile profile, string key)
        {
            int count;
            return profile.KeyCounts.TryGetValue(key, out count) ? count : 0;
        }

        private static void PairExactRows(
            IEnumerable<ExcelRow> srcRows,
            IEnumerable<ExcelRow> dstRows,
            IList<SmartColumn> columns,
            IList<SmartRowMatch> matches,
            HashSet<int> pairedSrcRows,
            HashSet<int> pairedDstRows)
        {
            var dstRowsBySignature = new Dictionary<string, Queue<ExcelRow>>();
            foreach (var dstRow in dstRows)
            {
                var signature = CreateRowSignature(dstRow, columns, false);
                Queue<ExcelRow> rowsWithSameSignature;
                if (!dstRowsBySignature.TryGetValue(signature, out rowsWithSameSignature))
                {
                    rowsWithSameSignature = new Queue<ExcelRow>();
                    dstRowsBySignature[signature] = rowsWithSameSignature;
                }

                rowsWithSameSignature.Enqueue(dstRow);
            }

            foreach (var srcRow in srcRows)
            {
                var signature = CreateRowSignature(srcRow, columns, true);
                Queue<ExcelRow> rowsWithSameSignature;
                if (!dstRowsBySignature.TryGetValue(signature, out rowsWithSameSignature)
                    || rowsWithSameSignature.Count == 0)
                    continue;

                var dstRow = rowsWithSameSignature.Dequeue();
                matches.Add(new SmartRowMatch(srcRow, dstRow));
                pairedSrcRows.Add(srcRow.Index);
                pairedDstRows.Add(dstRow.Index);
            }
        }

        private static void AddDataRowsInSourceOrder(
            ExcelSheet src,
            IList<SmartRow> rows,
            HashSet<int> pairedSrcRows,
            IList<SmartRowMatch> matches,
            IList<ExcelRow> addedDstRows)
        {
            var matchesBySrcIndex = matches.ToDictionary(m => m.SrcRow.Index, m => m);
            var matchesByDstOrder = matches.OrderBy(m => m.DstRow.Index).ToList();
            var addedRowsBeforeSrc = new Dictionary<int, List<ExcelRow>>();
            var trailingAddedRows = new List<ExcelRow>();

            foreach (var addedDstRow in addedDstRows.OrderBy(r => r.Index))
            {
                var nextMatch = matchesByDstOrder.FirstOrDefault(m => m.DstRow.Index > addedDstRow.Index);
                if (nextMatch == null)
                {
                    trailingAddedRows.Add(addedDstRow);
                    continue;
                }

                List<ExcelRow> list;
                if (!addedRowsBeforeSrc.TryGetValue(nextMatch.SrcRow.Index, out list))
                {
                    list = new List<ExcelRow>();
                    addedRowsBeforeSrc[nextMatch.SrcRow.Index] = list;
                }

                list.Add(addedDstRow);
            }

            foreach (var srcRow in src.Rows.Values)
            {
                List<ExcelRow> addedRows;
                if (addedRowsBeforeSrc.TryGetValue(srcRow.Index, out addedRows))
                {
                    foreach (var addedRow in addedRows)
                        rows.Add(new SmartRow(null, addedRow, ExcelRowStatus.Added));
                }

                SmartRowMatch match;
                if (matchesBySrcIndex.TryGetValue(srcRow.Index, out match))
                    rows.Add(new SmartRow(match.SrcRow, match.DstRow, ExcelRowStatus.None));
                else if (!pairedSrcRows.Contains(srcRow.Index))
                    rows.Add(new SmartRow(srcRow, null, ExcelRowStatus.Removed));
            }

            foreach (var addedRow in trailingAddedRows)
                rows.Add(new SmartRow(null, addedRow, ExcelRowStatus.Added));
        }

        private static void PairPrefixRows(
            ExcelSheet src,
            ExcelSheet dst,
            ExcelSheetDiffConfig config,
            IList<SmartRow> rows,
            HashSet<int> pairedSrcRows,
            HashSet<int> pairedDstRows)
        {
            var srcPrefixRows = src.Rows.Values.Where(r => CreateSkippedHeaderRows(src, config.SrcHeaderIndex).Contains(r.Index)).ToList();
            var dstPrefixRows = dst.Rows.Values.Where(r => CreateSkippedHeaderRows(dst, config.DstHeaderIndex).Contains(r.Index)).ToList();
            var count = Math.Max(srcPrefixRows.Count, dstPrefixRows.Count);
            for (int i = 0; i < count; i++)
            {
                var srcRow = i < srcPrefixRows.Count ? srcPrefixRows[i] : null;
                var dstRow = i < dstPrefixRows.Count ? dstPrefixRows[i] : null;
                var status = ExcelRowStatus.None;
                if (srcRow == null)
                    status = ExcelRowStatus.Added;
                else if (dstRow == null)
                    status = ExcelRowStatus.Removed;

                rows.Add(new SmartRow(srcRow, dstRow, status));
                if (srcRow != null)
                    pairedSrcRows.Add(srcRow.Index);
                if (dstRow != null)
                    pairedDstRows.Add(dstRow.Index);
            }
        }

        private static IEnumerable<SmartRow> CreateSmartRowsByLcs(ExcelSheet src, ExcelSheet dst, IList<SmartColumn> columns)
        {
            return CreateSmartRowsByLcs(src.Rows.Values, dst.Rows.Values, columns);
        }

        private static IEnumerable<SmartRow> CreateSmartRowsByLcs(
            IEnumerable<ExcelRow> srcRows,
            IEnumerable<ExcelRow> dstRows,
            IList<SmartColumn> columns)
        {
            var srcSignatures = srcRows.Select(r => new SmartRowSignature(r, CreateRowSignature(r, columns, true))).ToList();
            var dstSignatures = dstRows.Select(r => new SmartRowSignature(r, CreateRowSignature(r, columns, false))).ToList();
            var option = new DiffOption<SmartRowSignature>();
            var results = DiffUtil.Diff(srcSignatures, dstSignatures, option);
            results = DiffUtil.Order(results, DiffOrderType.LazyDeleteFirst);
            results = DiffUtil.OptimizeCaseDeletedFirst(results);

            foreach (var result in results)
            {
                if (result.Status == DiffStatus.Equal)
                {
                    yield return new SmartRow(result.Obj1.Row, result.Obj2.Row, ExcelRowStatus.None);
                }
                else if (result.Status == DiffStatus.Modified)
                {
                    var similarity = CalculateRowSimilarity(result.Obj1.Row, result.Obj2.Row, columns);
                    if (similarity >= SmartRowSimilarityThreshold)
                    {
                        yield return new SmartRow(result.Obj1.Row, result.Obj2.Row, ExcelRowStatus.None);
                    }
                    else
                    {
                        yield return new SmartRow(result.Obj1.Row, null, ExcelRowStatus.Removed);
                        yield return new SmartRow(null, result.Obj2.Row, ExcelRowStatus.Added);
                    }
                }
                else if (result.Status == DiffStatus.Deleted)
                {
                    yield return new SmartRow(result.Obj1.Row, null, ExcelRowStatus.Removed);
                }
                else if (result.Status == DiffStatus.Inserted)
                {
                    yield return new SmartRow(null, result.Obj2.Row, ExcelRowStatus.Added);
                }
            }
        }

        private static string CreateRowSignature(ExcelRow row, IEnumerable<SmartColumn> columns, bool isSource)
        {
            var builder = new StringBuilder();
            foreach (var column in columns.Where(c => c.SrcIndex.HasValue && c.DstIndex.HasValue))
            {
                var columnIndex = isSource ? column.SrcIndex.Value : column.DstIndex.Value;
                var value = GetCellValue(row, columnIndex);
                builder.Append(value.Length);
                builder.Append(':');
                builder.Append(value);
                builder.Append('\u001f');
            }

            return builder.ToString();
        }

        private static bool ShouldFallbackRowsToLegacy(ExcelSheet src, ExcelSheet dst)
        {
            return (long)src.Rows.Count * dst.Rows.Count > SmartRowLcsPairLimit;
        }

        private static double CalculateColumnSimilarity(ExcelColumn srcColumn, ExcelColumn dstColumn, int srcHeaderIndex, int dstHeaderIndex)
        {
            var count = Math.Max(srcColumn.Cells.Count, dstColumn.Cells.Count);
            if (count == 0)
                return 0;

            var compared = 0;
            var equal = 0;
            for (int i = 0; i < count; i++)
            {
                if (i == srcHeaderIndex || i == dstHeaderIndex)
                    continue;

                var srcValue = i < srcColumn.Cells.Count ? srcColumn.Cells[i].Value : string.Empty;
                var dstValue = i < dstColumn.Cells.Count ? dstColumn.Cells[i].Value : string.Empty;
                if (string.IsNullOrEmpty(srcValue) && string.IsNullOrEmpty(dstValue))
                    continue;

                compared++;
                if (srcValue == dstValue)
                    equal++;
            }

            if (compared == 0)
                return srcColumn.IsBlank() && dstColumn.IsBlank() ? 1.0 : 0.0;

            return (double)equal / compared;
        }

        private static double CalculateRowSimilarity(ExcelRow srcRow, ExcelRow dstRow, IEnumerable<SmartColumn> columns)
        {
            var compared = 0;
            var equal = 0;
            foreach (var column in columns.Where(c => c.SrcIndex.HasValue && c.DstIndex.HasValue))
            {
                var srcValue = GetCellValue(srcRow, column.SrcIndex.Value);
                var dstValue = GetCellValue(dstRow, column.DstIndex.Value);
                if (string.IsNullOrEmpty(srcValue) && string.IsNullOrEmpty(dstValue))
                    continue;

                compared++;
                if (srcValue == dstValue)
                    equal++;
            }

            return compared == 0 ? 0 : (double)equal / compared;
        }

        private static string GetHeaderText(ExcelSheet sheet, int rowIndex, int columnIndex)
        {
            ExcelRow row;
            if (!sheet.Rows.TryGetValue(rowIndex, out row))
                return string.Empty;

            return GetCellValue(row, columnIndex);
        }

        private static string GetCellValue(ExcelRow row, int columnIndex)
        {
            if (row == null || columnIndex < 0 || columnIndex >= row.Cells.Count)
                return string.Empty;

            return row.Cells[columnIndex].Value ?? string.Empty;
        }

        private static string NormalizeHeader(string value)
        {
            return (value ?? string.Empty).Trim();
        }

        private static string NormalizeKey(string value)
        {
            return (value ?? string.Empty).Trim();
        }

        private static bool IsPreferredKeyHeader(string header)
        {
            var text = NormalizeHeader(header).ToLowerInvariant();
            return text == "id"
                || text == "key"
                || text.EndsWith("id")
                || text.EndsWith("key")
                || text.Contains("_id")
                || text.Contains(" id")
                || text.Contains("编号")
                || text.Contains("编码")
                || text.Contains("序号");
        }

        private static bool IsFieldHeaderMarker(string header)
        {
            var text = NormalizeHeader(header).ToLowerInvariant();
            return text == "表头"
                || text == "header"
                || text == "headers"
                || text == "field"
                || text == "fields"
                || text == "fieldname"
                || text == "field_name"
                || text == "字段名";
        }

        private static bool IsMetadataHeaderRow(string firstCell)
        {
            var text = NormalizeHeader(firstCell).ToLowerInvariant();
            return text.Contains("export_type")
                || text.Contains("table_event")
                || text == "描述位"
                || text == "字段说明";
        }

        private static bool IsTypeHeaderRow(ExcelRow row)
        {
            var firstCell = NormalizeHeader(GetCellValue(row, 0));
            if (firstCell.Equals("Type", StringComparison.OrdinalIgnoreCase))
                return true;

            var nonEmpty = row.Cells.Select(c => NormalizeHeader(c.Value)).Where(v => !string.IsNullOrEmpty(v)).ToList();
            if (!nonEmpty.Any())
                return false;

            var typeLikeCount = nonEmpty.Count(IsTypeLikeValue);
            return typeLikeCount >= Math.Max(2, nonEmpty.Count / 2);
        }

        private static bool IsTypeLikeValue(string value)
        {
            var text = NormalizeHeader(value).ToLowerInvariant();
            return text == "type"
                || text == "int"
                || text == "bool"
                || text == "float"
                || text == "double"
                || text == "string"
                || text == "language"
                || text.StartsWith("enum(")
                || text.StartsWith("list(");
        }

        private static bool IsDefaultMetadataRow(ExcelRow row)
        {
            return NormalizeHeader(GetCellValue(row, 0)).Equals("Default", StringComparison.OrdinalIgnoreCase);
        }
    }
}
