using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using GBT9704_2012排版工具.Contracts;

namespace GBT9704_2012排版工具.Core;

public static class 正式公文服务
{
    public static void 格式化(WordprocessingDocument document, RequestContract request)
    {
        var main = document.MainDocumentPart!;
        var body = main.Document!.Body!;
        var f = request.要素!;
        var original = 公文要素服务.读取(main);
        var ps = body.Elements<Paragraph>().ToArray();
        var structure = new GovStructureAnalyzer().分析(main);
        var titles = structure.标题段索引;
        var firstTitle = titles.Order().FirstOrDefault(-1);
        if (firstTitle < 0)
            firstTitle = Array.FindIndex(ps, p => 公文要素服务.文本(p) == f.标题);
        if (firstTitle < 0) throw new InvalidOperationException("无法定位原稿标题，请核对标题原文后再处理。");

        var migrate = new HashSet<Paragraph>();
        foreach (var p in ps.Take(firstTitle))
        {
            request.CancellationToken.ThrowIfCancellationRequested();
            var t = 公文要素服务.文本(p);
            if (!结构安全分析器.段落可安全重写(p))
                throw new InvalidOperationException("原稿版头包含复杂对象，不能自动拆除或叠加红头，请先核对该对象。");
            if (t.Length == 0 || t == original.机关标志 || 公文要素服务.是文号候选(t)) migrate.Add(p);
            else throw new InvalidOperationException($"原稿版头存在未识别内容：{t}，请明确其用途后再处理。");
        }
        migrate.Add(ps[firstTitle]);
        foreach (var index in structure.主送机关段索引) migrate.Add(ps[index]);
        foreach (var p in ps)
        {
            request.CancellationToken.ThrowIfCancellationRequested();
            var t = 公文要素服务.文本(p);
            if (t.Length > 0 && (t == original.署名 || t == original.成文日期 || t == original.抄送 ||
                (original.印发机关.Length > 0 && t.StartsWith(original.印发机关) && t.EndsWith("印发")))) migrate.Add(p);
        }
        if (migrate.Any(p => !结构安全分析器.段落可安全重写(p)))
            throw new InvalidOperationException("待迁移的标题、落款或版记包含复杂对象，已停止自动重排。");
        // 页眉中的原有红头或文字对象不能在未定位来源时覆盖。
        if (main.HeaderParts.Any(h => h.Header?.Descendants<RunProperties>().Any(r => r.Color?.Val?.Value?.Equals("FF0000", StringComparison.OrdinalIgnoreCase) == true) == true))
            throw new InvalidOperationException("原稿页眉存在红色内容，可能是已有红头，请先核对以免重复。");
        foreach (var p in migrate) p.Remove();

        var intro = new List<Paragraph>();
        if (request.模式 == 排版模式.电子红头)
        {
            var agency = 段落(f.机关标志, "方正小标宋简体", 64, JustificationValues.Center, 800);
            // Word 与已核验小标宋 4.00 的字面上缘相差 108 缇，按字面而非段落起点定位。
            agency.ParagraphProperties!.SpacingBetweenLines!.Before = (毫米(35) - 108).ToString();
            agency.ParagraphProperties.SpacingBetweenLines.After = "1156";
            agency.ParagraphProperties.KeepNext = new KeepNext();
            foreach (var r in agency.Elements<Run>()) r.RunProperties!.Color = new Color { Val = "FF0000" };
            intro.Add(agency);
        }
        var number = 段落(f.文号, "仿宋_GB2312", 32, JustificationValues.Center);
        number.ParagraphProperties!.KeepNext = new KeepNext();
        number.ParagraphProperties.SpacingBetweenLines!.After = "1156";
        if (request.模式 == 排版模式.电子红头)
        {
            number.ParagraphProperties.ParagraphBorders = new ParagraphBorders(new BottomBorder { Val = BorderValues.Single, Color = "FF0000", Size = 8, Space = 8 });
            // Word 段落边框默认向两侧各延伸 1.5 磅，抵消后与 156 毫米版心等宽。
            number.ParagraphProperties.Indentation = new Indentation { Left = "30", Right = "30", FirstLine = "0" };
        }
        else
            number.ParagraphProperties.SpacingBetweenLines.Before = Math.Max(0, 毫米(request.套打红线距纸顶毫米 - 37 - 4) - 505).ToString();
        intro.Add(number);
        {
            var title = 段落(f.标题, "方正小标宋简体", 44, JustificationValues.Center);
            title.ParagraphProperties!.KeepNext = new KeepNext();
            title.ParagraphProperties.SpacingBetweenLines!.After = "578";
            intro.Add(title);
        }
        if (f.主送机关.Length > 0)
        {
            var recipient = 段落(f.主送机关, "仿宋_GB2312", 32, JustificationValues.Left);
            recipient.ParagraphProperties!.KeepNext = new KeepNext();
            intro.Add(recipient);
        }
        foreach (var p in intro.AsEnumerable().Reverse()) body.PrependChild(p);

        var attachment = body.Elements<Paragraph>().FirstOrDefault(p => Regex.IsMatch(公文要素服务.文本(p), @"^附件\s*\d*$"));
        var originalDate = Array.FindLastIndex(ps, p => 公文要素服务.文本(p) == original.成文日期 && original.成文日期.Length > 0);
        var following = originalDate >= 0 ? ps.Skip(originalDate + 1).FirstOrDefault(p => p.Parent == body) : null;
        var anchor = (OpenXmlElement?)following ?? attachment ?? (OpenXmlElement?)body.Elements<SectionProperties>().LastOrDefault();
        if (f.署名.Length > 0)
        {
            var signature = 段落(f.署名, "仿宋_GB2312", 32, JustificationValues.Right);
            var dateWidth = 文字宽度(f.成文日期);
            var nameWidth = 文字宽度(f.署名);
            var right = f.预留盖章 ? 64 + (dateWidth - nameWidth) / 2 : dateWidth > nameWidth ? 64 + dateWidth - nameWidth : 32;
            signature.ParagraphProperties!.Indentation = new Indentation { Right = ((int)Math.Round(right * 20)).ToString() };
            signature.ParagraphProperties!.SpacingBetweenLines!.Before = f.预留盖章 ? "1156" : "578";
            signature.ParagraphProperties.KeepNext = new KeepNext();
            插在末尾(body, signature, anchor);
        }
        var date = 段落(f.成文日期, "仿宋_GB2312", 32, JustificationValues.Right);
        GovOpenXmlHelper.设置段落缩进(date, rightChars: f.预留盖章 ? 4 : 2);
        if (!f.预留盖章 && f.署名.Length > 0 && 文字宽度(f.署名) >= 文字宽度(f.成文日期))
            date.ParagraphProperties!.Indentation = new Indentation { Right = ((int)Math.Round((文字宽度(f.署名) - 文字宽度(f.成文日期)) * 20)).ToString() };
        插在末尾(body, date, anchor);

        var remaining = body.Elements<Paragraph>().ToArray();
        for (var i = 0; i < remaining.Length; i++)
        {
            request.CancellationToken.ThrowIfCancellationRequested();
            var label = remaining[i];
            if (!Regex.IsMatch(公文要素服务.文本(label), @"^附件\s*\d*$")) continue;
            if (!结构安全分析器.段落可安全重写(label)) continue;
            label.ParagraphProperties ??= new ParagraphProperties();
            label.ParagraphProperties.PageBreakBefore = new PageBreakBefore();
            label.ParagraphProperties.KeepNext = new KeepNext();
            label.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines { Before = "0", After = "578", Line = "578", LineRule = LineSpacingRuleValues.Exact };
            GovOpenXmlHelper.设置段落缩进(label);
            foreach (var r in label.Elements<Run>()) GovOpenXmlHelper.设置运行格式(r, "黑体", "Times New Roman", "32");
            var title = remaining.Skip(i + 1).FirstOrDefault(p => 公文要素服务.文本(p).Length > 0);
            if (title is not null && 结构安全分析器.段落可安全重写(title))
            {
                GovOpenXmlHelper.设置段落缩进(title);
                title.ParagraphProperties!.Justification = new Justification { Val = JustificationValues.Center };
                title.ParagraphProperties.KeepNext = new KeepNext();
                foreach (var r in title.Elements<Run>()) GovOpenXmlHelper.设置运行格式(r, "方正小标宋简体", "Times New Roman", "44");
            }
        }
        if (f.抄送.Length > 0 || f.印发机关.Length > 0)
            插在末尾(body, 版记(f), body.Elements<SectionProperties>().LastOrDefault());
        request.检查记录.Add(new("FORMAL_LAYOUT", "正式公文结构", "通过", $"已生成{request.模式}结构；页面位置与分页仍需 Word 实测。"));
    }

    private static int 毫米(double value) => (int)Math.Round(value * 1440 / 25.4);

    private static double 文字宽度(string text)
    {
        var total = 0d;
        foreach (var c in text)
        {
            var family = new System.Windows.Media.Typeface(c < 128 ? "Times New Roman" : "仿宋_GB2312");
            if (family.TryGetGlyphTypeface(out var glyph) && glyph.CharacterToGlyphMap.TryGetValue(c, out var index))
                total += glyph.AdvanceWidths[index] * 16;
            else total += 16;
        }
        return total;
    }

    private static void 插在末尾(Body body, OpenXmlElement element, OpenXmlElement? anchor)
    {
        if (anchor is null) body.Append(element); else body.InsertBefore(element, anchor);
    }

    private static Paragraph 段落(string text, string font, int size, JustificationValues alignment, int line = 578)
    {
        var run = new Run();
        foreach (var lineText in text.Replace("\r", "").Split('\n'))
        {
            if (run.ChildElements.Count > 0) run.Append(new Break());
            run.Append(new Text(lineText) { Space = SpaceProcessingModeValues.Preserve });
        }
        var p = new Paragraph(run);
        GovOpenXmlHelper.设置段落行距(p, line);
        GovOpenXmlHelper.设置段落缩进(p);
        p.ParagraphProperties!.Justification = new Justification { Val = alignment };
        p.ParagraphProperties.SnapToGrid = new SnapToGrid { Val = false };
        p.ParagraphProperties.AutoSpaceDE = new AutoSpaceDE { Val = false };
        p.ParagraphProperties.AutoSpaceDN = new AutoSpaceDN { Val = false };
        foreach (var r in p.Elements<Run>()) GovOpenXmlHelper.设置运行格式(r, font, "Times New Roman", size.ToString());
        return p;
    }

    private static Table 版记(公文要素 f)
    {
        var props = new TableProperties();
        props.AddChild(new TablePositionProperties
        {
            VerticalAnchor = VerticalAnchorValues.Margin,
            HorizontalAnchor = HorizontalAnchorValues.Margin,
            TablePositionX = 0,
            TablePositionYAlignment = VerticalAlignmentValues.Bottom,
            TopFromText = 280,
            BottomFromText = 0,
            LeftFromText = 0,
            RightFromText = 0
        });
        props.AddChild(new TableWidth { Type = TableWidthUnitValues.Dxa, Width = "8844" });
        props.AddChild(new TableBorders(new TopBorder { Val = BorderValues.Single, Size = 8 }, new BottomBorder { Val = BorderValues.Single, Size = 8 }, new InsideHorizontalBorder { Val = BorderValues.Single, Size = 6 }));
        props.AddChild(new TableLayout { Type = TableLayoutValues.Fixed });
        props.AddChild(new TableCellMarginDefault(new TableCellLeftMargin { Width = 0, Type = TableWidthValues.Dxa }, new TableCellRightMargin { Width = 0, Type = TableWidthValues.Dxa }));
        var table = new Table(props, new TableGrid(new GridColumn { Width = "8844" }));
        if (f.抄送.Length > 0) table.Append(行(f.抄送));
        if (f.印发机关.Length > 0)
        {
            var row = 行(f.印发机关);
            var p = row.Descendants<Paragraph>().Single();
            p.ParagraphProperties!.Tabs = new Tabs(new TabStop { Val = TabStopValues.Right, Position = 8564 });
            var right = new Run(new TabChar(), new Text(f.印发日期 + "印发"));
            GovOpenXmlHelper.设置运行格式(right, "仿宋_GB2312", "Times New Roman", "28");
            p.Append(right);
            table.Append(row);
        }
        return table;
    }

    private static TableRow 行(string text)
    {
        var p = 段落(text, "仿宋_GB2312", 28, JustificationValues.Left, 400);
        p.ParagraphProperties!.Indentation = new Indentation { Left = "280", Right = "280" };
        return new TableRow(new TableCell(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "8844" }), p));
    }
}
