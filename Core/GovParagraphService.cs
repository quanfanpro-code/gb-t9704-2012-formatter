using DocumentFormat.OpenXml.Wordprocessing;

using DocumentFormat.OpenXml.Packaging;

namespace GBT9704_2012排版工具.Core;

public sealed class GovParagraphService
{
    private const int 固定行距 = 578;
    // 未安装方正小标宋_GBK 时由 GovFontCheckService 退回方正小标宋简体
    internal static readonly string 标题字体 = GovFontCheckService.选定标题字体();
    private const string 一级标题字体 = "黑体";
    private const string 二级标题字体 = "楷体_GB2312";
    private const string 正文字体 = "仿宋_GB2312";
    private const string 西文字体 = "Times New Roman";
    private const string 标题字号 = "44";
    private const string 正文字号 = "32";

    public void 格式化(Body body, GovDocumentStructure structure, CancellationToken cancellationToken = default)
    {
        var paragraphs = body.Elements<Paragraph>()
            .Where(x => !x.Ancestors<Table>().Any())
            .ToList();

        for (var i = 0; i < paragraphs.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var paragraph = paragraphs[i];
            var text = GovOpenXmlHelper.提取可见文本(paragraph).Trim();
            var isEmpty = string.IsNullOrWhiteSpace(text);

            if (structure.文号段索引.Count > 0 && i < structure.文号段索引.Min())
                continue;

            if (isEmpty)
            {
                continue;
            }

            GovOpenXmlHelper.设置段落行距(paragraph, 固定行距);
            paragraph.ParagraphProperties!.OverflowPunctuation = new OverflowPunctuation { Val = false };
            paragraph.ParagraphProperties.AutoSpaceDE = new AutoSpaceDE { Val = false };
            paragraph.ParagraphProperties.AutoSpaceDN = new AutoSpaceDN { Val = false };

            if (structure.标题段索引.Contains(i))
            {
                格式化标题区(paragraph);
                continue;
            }

            if (structure.文号段索引.Contains(i))
            {
                格式化文号(paragraph);
                continue;
            }

            if (structure.主送机关段索引.Contains(i))
            {
                格式化主送机关(paragraph);
                continue;
            }

            if (structure.日期段索引.Contains(i))
            {
                格式化日期(paragraph);
                continue;
            }

            if (structure.附件段索引.Contains(i))
            {
                格式化附件(paragraph);
                continue;
            }

            if (structure.版记段索引.Contains(i))
            {
                格式化版记(paragraph);
                continue;
            }

            if (structure.标题级别映射.TryGetValue(i, out var level))
            {
                格式化正文标题(paragraph, level);
                continue;
            }

            格式化正文(paragraph);
        }
    }

    private static void 格式化标题区(Paragraph paragraph)
    {
        GovOpenXmlHelper.设置段落缩进(paragraph);
        var pPr = GovOpenXmlHelper.确保段落属性(paragraph);
        pPr.Justification = new Justification { Val = JustificationValues.Center };
        格式化运行(paragraph, 标题字体, "Times New Roman", 标题字号);
    }

    private static void 格式化文号(Paragraph paragraph)
    {
        GovOpenXmlHelper.设置段落缩进(paragraph);
        var pPr = GovOpenXmlHelper.确保段落属性(paragraph);
        pPr.Justification = new Justification { Val = JustificationValues.Center };
        格式化运行(paragraph, 正文字体, 西文字体, 正文字号);
    }

    private static void 格式化主送机关(Paragraph paragraph)
    {
        GovOpenXmlHelper.设置段落缩进(paragraph);
        var pPr = GovOpenXmlHelper.确保段落属性(paragraph);
        pPr.Justification = new Justification { Val = JustificationValues.Left };
        格式化运行(paragraph, 正文字体, 西文字体, 正文字号);
    }

    private static void 格式化日期(Paragraph paragraph)
    {
        GovOpenXmlHelper.设置段落缩进(paragraph, rightChars: 4);
        var pPr = GovOpenXmlHelper.确保段落属性(paragraph);
        pPr.Justification = new Justification { Val = JustificationValues.Right };
        格式化运行(paragraph, 正文字体, 西文字体, 正文字号);
    }

    private static void 格式化附件(Paragraph paragraph)
    {
        var first = GovOpenXmlHelper.提取可见文本(paragraph).TrimStart().StartsWith("附件");
        GovOpenXmlHelper.设置段落缩进(paragraph, leftChars: 6);
        var pPr = GovOpenXmlHelper.确保段落属性(paragraph);
        pPr.Indentation!.FirstLine = null;
        pPr.Indentation.FirstLineChars = null;
        pPr.Indentation.HangingChars = first ? 400 : 100;
        pPr.OutlineLevel = null;
        pPr.ParagraphStyleId = null;
        pPr.KeepNext = new KeepNext();
        if (first) pPr.SpacingBetweenLines!.Before = "578";
        pPr.Justification = new Justification { Val = JustificationValues.Left };
        格式化运行(paragraph, 正文字体, 西文字体, 正文字号);
    }

    private static void 格式化版记(Paragraph paragraph)
    {
        GovOpenXmlHelper.设置段落缩进(paragraph);
        var pPr = GovOpenXmlHelper.确保段落属性(paragraph);
        pPr.Justification = new Justification { Val = JustificationValues.Left };
        格式化运行(paragraph, 正文字体, 西文字体, "28");
    }

    private static void 格式化正文标题(Paragraph paragraph, int level)
    {
        // GB/T 9704-2012: 层次标题（一、（一）等）与正文一样左空二字
        GovOpenXmlHelper.设置段落缩进(paragraph, firstLineChars: 2);
        var pPr = GovOpenXmlHelper.确保段落属性(paragraph);
        pPr.Justification = new Justification { Val = JustificationValues.Left };

        var font = level switch
        {
            1 => 一级标题字体,
            2 => 二级标题字体,
            _ => 正文字体
        };

        // GB/T 9704-2012: 第一层黑体，第二层楷体，第三四层仿宋。不加粗，避免黑体在Word中触发伪粗体(Fake Bold)发虚
        var bold = false;
        pPr.ParagraphStyleId = new ParagraphStyleId { Val = $"GBT9704Heading{level}" };
        pPr.OutlineLevel = new OutlineLevel { Val = level - 1 };
        pPr.KeepNext = new KeepNext();
        格式化运行(paragraph, font, 西文字体, 正文字号, bold);
    }

    public void 补充层次样式(MainDocumentPart main)
    {
        var part = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        part.Styles ??= new Styles();
        for (var level = 1; level <= 4; level++)
        {
            var id = $"GBT9704Heading{level}";
            if (part.Styles.Elements<Style>().Any(s => s.StyleId == id)) continue;
            part.Styles.Append(new Style(new StyleName { Val = $"公文层次标题{level}" },
                new StyleParagraphProperties(new KeepNext(), new OutlineLevel { Val = level - 1 }))
            { Type = StyleValues.Paragraph, StyleId = id, CustomStyle = true });
        }
    }

    private static void 格式化正文(Paragraph paragraph)
    {
        GovOpenXmlHelper.设置段落缩进(paragraph, firstLineChars: 2);
        var pPr = GovOpenXmlHelper.确保段落属性(paragraph);
        pPr.Justification = new Justification { Val = JustificationValues.Both };
        格式化运行(paragraph, 正文字体, 西文字体, 正文字号);
        // 三号字在 156 毫米版心中排 28 字，收紧字间距而不缩放字形；Word 样本验证首行 26 字、后续 28 字。
        foreach (var run in GovOpenXmlHelper.获取文本运行(paragraph))
            run.RunProperties!.Spacing = new Spacing { Val = -5 };
    }

    private static void 格式化运行(Paragraph paragraph, string 中文字体, string latinFont, string fontSize, bool bold = false)
    {
        foreach (var run in GovOpenXmlHelper.获取文本运行(paragraph))
            GovOpenXmlHelper.设置运行格式(run, 中文字体, latinFont, fontSize, bold);
    }
}
