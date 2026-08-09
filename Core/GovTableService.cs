using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GBT9704_2012排版工具.Core;

public sealed class GovTableService
{
    private const string 中文字体 = "仿宋_GB2312";
    private const string 西文字体 = "Times New Roman";
    private const string 字号 = "28";
    private static readonly Regex 纯数字正则 = new(@"^\(?-?[\d,]+(?:\.\d+)?\)?$", RegexOptions.Compiled);

    /// <summary>
    /// 上一轮 格式化 中因含高风险结构（图片/超链接/书签/批注/域等）而跳过文本重写的数字单元格数。
    /// 这些单元格仍做了边框/对齐/字体统一，只是不整体重写，调用方应将该计数呈现给用户以便人工复核。
    /// </summary>
    public int 跳过不安全单元格数 { get; private set; }

    public void 格式化(Body body)
    {
        跳过不安全单元格数 = 0;

        foreach (var table in body.Descendants<Table>())
        {
            删除尾部空行(table);
            删除尾部空列(table);
            设置表格边框(table);
            格式化单元格(table);
        }
    }

    public static string? 格式化纯数字文本(string? rawText, bool isFirstColumn)
    {
        if (string.IsNullOrWhiteSpace(rawText))
            return null;

        var text = rawText.Trim();
        if (!纯数字正则.IsMatch(text))
            return null;

        // 防御：前导零（例如 "0123", "001"）可能是编码，不格式化
        if (text.Length > 1 && text.StartsWith('0') && !text.StartsWith("0."))
            return null;

        var hasOpenBracket = text.Contains('(');
        var hasCloseBracket = text.Contains(')');
        if (hasOpenBracket != hasCloseBracket)
            return null;

        var isBracketNegative = text.StartsWith('(') && text.EndsWith(')');
        var normalized = text.Replace(",", string.Empty)
            .Replace("(", string.Empty)
            .Replace(")", string.Empty);

        if (!decimal.TryParse(normalized, NumberStyles.Number | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var value))
            return null;

        // 防御：首列正整数（通常是序号 1, 2, 3），不格式化为小数
        if (isFirstColumn && value > 0 && value == Math.Truncate(value) && !text.Contains('.'))
            return null;

        if (isBracketNegative && value > 0)
        {
            value = -value;
        }

        return value.ToString("#,##0.00", CultureInfo.InvariantCulture);
    }

    private void 格式化单元格(Table table)
    {
        foreach (var row in table.Descendants<TableRow>())
        {
            var cells = row.Elements<TableCell>().ToList();
            for (var i = 0; i < cells.Count; i++)
            {
                var cell = cells[i];
                var isFirstColumn = i == 0;

                cell.TableCellProperties ??= new TableCellProperties();
                cell.TableCellProperties.TableCellVerticalAlignment = new TableCellVerticalAlignment
                {
                    Val = TableVerticalAlignmentValues.Center
                };

                var paragraphs = cell.Elements<Paragraph>().ToList();
                if (paragraphs.Count == 0)
                {
                    paragraphs.Add(cell.AppendChild(new Paragraph()));
                }

                var cellText = string.Join(Environment.NewLine, paragraphs.Select(GovOpenXmlHelper.段落文本)).Trim();
                var formattedNumber = 格式化纯数字文本(cellText, isFirstColumn);

                // 防护门：数字单元格整体重写前先过安全分析，含图片/超链接/书签/批注/域等高风险结构
                // （或多段落）的单元格禁止 RemoveAllChildren，否则书签/批注/域会被静默删除。
                // 不安全时按普通文本单元格处理：只统一字体，不动结构。
                var isNumericCell = formattedNumber != null && 结构安全分析器.单元格可安全重写(cell);
                if (formattedNumber != null && !isNumericCell)
                {
                    跳过不安全单元格数++;
                    System.Diagnostics.Trace.WriteLine(
                        $"[GovTableService] 数字单元格含高风险结构，已跳过文本重写：{cellText}");
                }

                if (isNumericCell)
                {
                    cell.RemoveAllChildren<Paragraph>();
                    paragraphs = [cell.AppendChild(new Paragraph())];
                }

                foreach (var paragraph in paragraphs)
                {
                    GovOpenXmlHelper.清理段落污染(paragraph);
                    GovOpenXmlHelper.设置段落行距(paragraph, 578);
                    GovOpenXmlHelper.设置段落缩进(paragraph);

                    var pPr = GovOpenXmlHelper.确保段落属性(paragraph);
                    pPr.Justification = new Justification
                    {
                        Val = isNumericCell ? JustificationValues.Right : (isFirstColumn ? JustificationValues.Left : JustificationValues.Center)
                    };

                    if (isNumericCell)
                    {
                        GovOpenXmlHelper.替换段落文本(paragraph, formattedNumber!, 中文字体, 西文字体, 字号);
                    }
                    else
                    {
                        foreach (var run in GovOpenXmlHelper.获取文本运行(paragraph))
                        {
                            GovOpenXmlHelper.设置运行格式(run, 中文字体, 西文字体, 字号);
                        }
                    }
                }
            }
        }
    }

    private static void 设置表格边框(Table table)
    {
        var props = table.GetFirstChild<TableProperties>();
        if (props == null)
        {
            props = new TableProperties();
            table.PrependChild(props);
        }

        props.TableBorders = new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 12U, Color = "000000" },
            new BottomBorder { Val = BorderValues.Single, Size = 12U, Color = "000000" },
            new LeftBorder { Val = BorderValues.Single, Size = 12U, Color = "000000" },
            new RightBorder { Val = BorderValues.Single, Size = 12U, Color = "000000" },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4U, Color = "000000" }
        );
    }

    private static void 删除尾部空行(Table table)
    {
        while (true)
        {
            var lastRow = table.Elements<TableRow>().LastOrDefault();
            if (lastRow == null)
                return;

            var allEmpty = lastRow.Elements<TableCell>()
                .All(cell => string.IsNullOrWhiteSpace(cell.InnerText));

            if (!allEmpty)
                return;

            // 仅在尾行存在合并续接单元格时跳过，避免误删跨行合并的目标行
            // 注意 <w:vMerge/> 无 val 属性时 OOXML 默认即 continue，必须一并识别
            if (lastRow.Elements<TableCell>().Any(cell =>
                cell.TableCellProperties?.VerticalMerge is { } merge &&
                (merge.Val == null || merge.Val.Value == MergedCellValues.Continue)))
                return;

            lastRow.Remove();
        }
    }

    private static void 删除尾部空列(Table table)
    {
        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count == 0)
            return;

        while (true)
        {
            var maxCellCount = rows.Max(x => x.Elements<TableCell>().Count());
            if (maxCellCount == 0)
                return;

            var lastColumnIndex = maxCellCount - 1;
            var allEmpty = rows.All(row =>
            {
                var cells = row.Elements<TableCell>().ToList();
                return cells.Count <= lastColumnIndex || string.IsNullOrWhiteSpace(cells[lastColumnIndex].InnerText);
            });

            if (!allEmpty)
                return;

            // 仅在尾列存在合并续接单元格时跳过，避免误删跨列合并的目标列
            // 与纵合并同理：<w:hMerge/> 无 val 属性时默认即 continue
            var hasMerge = rows.Any(row =>
            {
                var cells = row.Elements<TableCell>().ToList();
                return cells.Count > lastColumnIndex &&
                    cells[lastColumnIndex].TableCellProperties?.HorizontalMerge is { } merge &&
                    (merge.Val == null || merge.Val.Value == MergedCellValues.Continue);
            });
            if (hasMerge)
                return;

            foreach (var row in rows)
            {
                var cells = row.Elements<TableCell>().ToList();
                if (cells.Count > lastColumnIndex)
                {
                    cells[lastColumnIndex].Remove();
                }
            }
        }
    }
}

