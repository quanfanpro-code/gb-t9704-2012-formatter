using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GBT9704_2012排版工具.Core;

public static class GovOpenXmlHelper
{
    private static readonly Regex 空白压缩 = new(@"\s+", RegexOptions.Compiled);
    private static readonly Regex 零宽字符正则 = new(@"[\u200B-\u200D\uFEFF]", RegexOptions.Compiled);

    public static string 规范文本(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        return 零宽字符正则.Replace(text, "")
            .Replace("\u00A0", " ")
            .Replace("\u3000", " ")
            .Replace("\t", " ")
            .Replace("\r", " ")
            .Replace("\n", " ")
            .Replace("．", ".")
            .Replace("－", "-")
            .Replace("—", "-")
            .Replace("　", " ")
            .Pipe(x => 空白压缩.Replace(x, ""))
            .Replace("（", "(")
            .Replace("）", ")")
            .Replace("：", ":")
            .Trim();
    }

    public static string 段落文本(Paragraph paragraph)
    {
        return string.Concat(paragraph.Descendants<Text>().Select(x => x.Text));
    }

    public static string 提取可见文本(OpenXmlElement? element)
    {
        if (element == null)
            return string.Empty;

        var builder = new StringBuilder();
        追加可见文本(element, builder);
        return builder.ToString();
    }

    public static string 提取归一化可见文本(OpenXmlElement? element)
    {
        return 规范文本(提取可见文本(element));
    }

    /// <summary>
    /// 仅取段落的直接子 Run。取舍说明：超链接/SimpleField 内的 Run 不参与字体统一——
    /// 这类结构会被 结构安全分析器 判为高风险并跳过整体重写，若此处再深入子孙节点改字体，
    /// 等于绕过防护门去动我们声明不碰的内容（Drawing 文本框、域结果等），风险大于收益。
    /// 代价是超链接文字字体不统一，属可接受的视觉残留，留给人工处理。
    /// </summary>
    public static IEnumerable<Run> 获取文本运行(Paragraph paragraph)
    {
        return paragraph.Elements<Run>().Where(x => !string.IsNullOrWhiteSpace(x.InnerText));
    }

    public static ParagraphProperties 确保段落属性(Paragraph paragraph)
    {
        paragraph.ParagraphProperties ??= new ParagraphProperties();
        return paragraph.ParagraphProperties;
    }

    /// <summary>
    /// 删除段落内除 pPr 外的一切子元素（含 Drawing/Hyperlink/书签/批注/域等）。
    /// 调用方必须先过 结构安全分析器.段落可安全重写，确认不含高风险结构后才能调用。
    /// </summary>
    public static void 清理段落内容(Paragraph paragraph)
    {
        foreach (var child in paragraph.ChildElements.Where(x => x is not ParagraphProperties).ToList())
        {
            child.Remove();
        }
    }

    public static void 清理段落污染(Paragraph paragraph)
    {
        var pPr = 确保段落属性(paragraph);
        pPr.GetFirstChild<Tabs>()?.Remove();
        pPr.GetFirstChild<NumberingProperties>()?.Remove();
        pPr.GetFirstChild<ParagraphStyleId>()?.Remove();
        pPr.GetFirstChild<ContextualSpacing>()?.Remove();
        pPr.GetFirstChild<MirrorIndents>()?.Remove();
        pPr.GetFirstChild<SuppressAutoHyphens>()?.Remove();
        pPr.GetFirstChild<SuppressLineNumbers>()?.Remove();
        pPr.GetFirstChild<WordWrap>()?.Remove();
        pPr.GetFirstChild<TextDirection>()?.Remove();
        pPr.GetFirstChild<TextAlignment>()?.Remove();
        if (pPr.KeepNext != null) pPr.KeepNext.Remove();
        if (pPr.KeepLines != null) pPr.KeepLines.Remove();
        if (pPr.PageBreakBefore != null) pPr.PageBreakBefore.Remove();
    }

    public static void 设置中文西文字体(OpenXmlElement runPropertiesElement, string eastAsiaFont, string latinFont)
    {
        if (runPropertiesElement is not OpenXmlCompositeElement composite)
            return;

        var fonts = composite.GetFirstChild<RunFonts>();
        if (fonts == null)
        {
            fonts = new RunFonts();
            composite.PrependChild(fonts);
        }

        fonts.Ascii = latinFont;
        fonts.HighAnsi = latinFont;
        fonts.ComplexScript = latinFont;
        fonts.EastAsia = eastAsiaFont;
    }

    public static void 设置运行格式(Run run, string eastAsiaFont, string latinFont, string fontSize, bool bold = false)
    {
        run.RunProperties ??= new RunProperties();

        var rPr = run.RunProperties;
        rPr.FontSize = new FontSize { Val = fontSize };
        rPr.FontSizeComplexScript = new FontSizeComplexScript { Val = fontSize };
        rPr.Bold = new Bold { Val = bold };
        rPr.BoldComplexScript = new BoldComplexScript { Val = bold };
        rPr.Italic = new Italic { Val = false };
        rPr.ItalicComplexScript = new ItalicComplexScript { Val = false };
        rPr.GetFirstChild<Underline>()?.Remove();
        rPr.GetFirstChild<RunStyle>()?.Remove();
        rPr.GetFirstChild<FitText>()?.Remove();
        rPr.GetFirstChild<CharacterScale>()?.Remove();
        rPr.GetFirstChild<Spacing>()?.Remove();
        rPr.GetFirstChild<Position>()?.Remove();
        设置中文西文字体(rPr, eastAsiaFont, latinFont);
    }

    public static void 设置段落行距(Paragraph paragraph, int lineTwips)
    {
        var pPr = 确保段落属性(paragraph);
        pPr.SpacingBetweenLines = new SpacingBetweenLines
        {
            Line = lineTwips.ToString(),
            LineRule = LineSpacingRuleValues.Exact,
            Before = "0",
            After = "0",
            BeforeLines = 0,
            AfterLines = 0,
            BeforeAutoSpacing = false,
            AfterAutoSpacing = false
        };
    }

    public static void 设置段落缩进(Paragraph paragraph, int? firstLineChars = null, int? leftChars = null, int? rightChars = null)
    {
        var pPr = 确保段落属性(paragraph);
        var ind = new Indentation
        {
            Left = "0",
            Right = "0",
            FirstLine = "0"
        };

        if (firstLineChars.HasValue)
        {
            ind.FirstLine = null;
            ind.FirstLineChars = firstLineChars.Value * 100;
        }
        else
        {
            ind.FirstLineChars = 0;
        }

        if (leftChars.HasValue)
        {
            ind.Left = null;
            ind.LeftChars = leftChars.Value * 100;
        }
        else
        {
            ind.LeftChars = 0;
        }

        if (rightChars.HasValue)
        {
            ind.Right = null;
            ind.RightChars = rightChars.Value * 100;
        }
        else
        {
            ind.RightChars = 0;
        }

        ind.Hanging = null;
        ind.HangingChars = null;
        pPr.Indentation = ind;
    }

    /// <summary>
    /// 整体重写段落文本（内部调用 清理段落内容，会删除 Drawing/Hyperlink/书签/批注/域等）。
    /// 调用方必须先过 结构安全分析器.段落可安全重写，确认不含高风险结构后才能调用。
    /// </summary>
    public static void 替换段落文本(Paragraph paragraph, string text, string eastAsiaFont, string latinFont, string fontSize, bool bold = false)
    {
        清理段落内容(paragraph);
        var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        设置运行格式(run, eastAsiaFont, latinFont, fontSize, bold);
        paragraph.AppendChild(run);
    }

    public static int 解析样式标题级别(MainDocumentPart? mainPart, Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (string.IsNullOrWhiteSpace(styleId))
            return 0;

        return 解析样式标题级别(mainPart, styleId);
    }

    public static int 解析样式标题级别(MainDocumentPart? mainPart, string? styleId)
    {
        if (string.IsNullOrWhiteSpace(styleId))
            return 0;

        var key = styleId.Trim().ToLowerInvariant();
        if (key is "heading1" or "标题1" or "title" or "title1") return 1;
        if (key is "heading2" or "标题2") return 2;
        if (key is "heading3" or "标题3") return 3;
        if (key is "heading4" or "标题4") return 4;

        var styles = mainPart?.StyleDefinitionsPart?.Styles;
        if (styles == null)
            return 0;

        var style = styles.Elements<Style>().FirstOrDefault(x => x.StyleId?.Value == styleId);
        if (style == null)
            return 0;

        var styleName = style.StyleName?.Val?.Value?.Trim().ToLowerInvariant();
        if (styleName is "标题 1" or "一级标题") return 1;
        if (styleName is "标题 2" or "二级标题") return 2;
        if (styleName is "标题 3" or "三级标题") return 3;
        if (styleName is "标题 4" or "四级标题") return 4;
        return 0;
    }

    public static int 解析大纲级别(Paragraph paragraph)
    {
        var outline = paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
        return outline switch
        {
            0 => 1,
            1 => 2,
            2 => 3,
            3 => 4,
            _ => 0
        };
    }

    public static void 添加页码域(Paragraph paragraph, string latinFont, string fontSize)
    {
        var dashLeft = new Run(new Text("—"));
        设置运行格式(dashLeft, "宋体", latinFont, fontSize);
        paragraph.Append(dashLeft);

        var pageRun = new Run();
        设置运行格式(pageRun, "宋体", latinFont, fontSize);
        pageRun.Append(
            new FieldChar { FieldCharType = FieldCharValues.Begin },
            new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve },
            new FieldChar { FieldCharType = FieldCharValues.Separate },
            new Text("1"),
            new FieldChar { FieldCharType = FieldCharValues.End }
        );
        paragraph.Append(pageRun);

        var dashRight = new Run(new Text("—"));
        设置运行格式(dashRight, "宋体", latinFont, fontSize);
        paragraph.Append(dashRight);
    }

    public static void 合并相邻运行(Paragraph paragraph)
    {
        var runs = paragraph.Elements<Run>().ToList();
        if (runs.Count < 2)
            return;

        for (var i = 0; i < runs.Count - 1; i++)
        {
            var current = runs[i];
            var next = runs[i + 1];

            if (!可合并文本运行(current) || !可合并文本运行(next))
                continue;

            if (!运行格式一致(current, next))
                continue;

            var currentText = current.GetFirstChild<Text>()!;
            var nextText = next.GetFirstChild<Text>()!;
            currentText.Text += nextText.Text;
            currentText.Space = currentText.Space?.Value == SpaceProcessingModeValues.Preserve ||
                                nextText.Space?.Value == SpaceProcessingModeValues.Preserve
                ? SpaceProcessingModeValues.Preserve
                : null;
            next.Remove();
            runs.RemoveAt(i + 1);
            i--;
        }
    }

    public static void 合并文档相邻运行(WordprocessingDocument document)
    {
        foreach (var paragraph in 获取全部段落(document))
        {
            合并相邻运行(paragraph);
        }
    }

    private static void 追加可见文本(OpenXmlElement element, StringBuilder builder)
    {
        switch (element)
        {
            case DeletedText:
            case FieldCode:
            case BookmarkStart:
            case BookmarkEnd:
            case CommentReference:
                return;
            case Drawing drawing:
                foreach (var textBox in drawing.Descendants<TextBoxContent>())
                {
                    foreach (var child in textBox.ChildElements)
                    {
                        追加可见文本(child, builder);
                    }
                }
                return;
            case Text text:
                builder.Append(text.Text);
                return;
            case TabChar:
                builder.Append('\t');
                return;
            case Break:
            case CarriageReturn:
                builder.Append('\n');
                return;
            case NoBreakHyphen:
                builder.Append('-');
                return;
            // <w:vanish w:val="false"/> 表示显式不隐藏，不能当隐藏文字跳过
            case Run run when run.RunProperties?.Vanish != null && run.RunProperties.Vanish.Val?.Value != false:
                return;
        }

        foreach (var child in element.ChildElements)
        {
            追加可见文本(child, builder);
        }
    }

    private static bool 可合并文本运行(Run run)
    {
        // 保守方案：仅恰含一个 w:t 的纯文本运行可合并。
        // 含多个 w:t 的运行若只拼第一个就 Remove，其余文本会被静默删除，故不合并。
        return run.Elements<Text>().Count() == 1 &&
               run.ChildElements.All(x => x is RunProperties or Text);
    }

    private static bool 运行格式一致(Run left, Run right)
    {
        var leftProps = left.RunProperties?.OuterXml ?? string.Empty;
        var rightProps = right.RunProperties?.OuterXml ?? string.Empty;
        return string.Equals(leftProps, rightProps, StringComparison.Ordinal);
    }

    private static IEnumerable<Paragraph> 获取全部段落(WordprocessingDocument document)
    {
        if (document.MainDocumentPart?.Document?.Body != null)
        {
            foreach (var paragraph in document.MainDocumentPart.Document.Body.Descendants<Paragraph>())
            {
                yield return paragraph;
            }
        }

        foreach (var headerPart in document.MainDocumentPart?.HeaderParts ?? [])
        {
            foreach (var paragraph in headerPart.Header?.Descendants<Paragraph>() ?? [])
            {
                yield return paragraph;
            }
        }

        foreach (var footerPart in document.MainDocumentPart?.FooterParts ?? [])
        {
            foreach (var paragraph in footerPart.Footer?.Descendants<Paragraph>() ?? [])
            {
                yield return paragraph;
            }
        }
    }
    public static T Pipe<T>(this T value, Func<T, T> transform) => transform(value);
}

file static class 文本管道扩展
{
}
