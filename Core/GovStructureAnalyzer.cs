using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GBT9704_2012排版工具.Core;

public sealed class GovStructureAnalyzer
{
    private static readonly Regex 一级标题正则 = new(@"^[一二三四五六七八九十]+、", RegexOptions.Compiled);
    private static readonly Regex 二级标题正则 = new(@"^[\(（][一二三四五六七八九十]+[\)）]", RegexOptions.Compiled);
    private static readonly Regex 三级标题正则 = new(@"^\d+[\.．、]\s*", RegexOptions.Compiled);
    private static readonly Regex 四级标题正则 = new(@"^[\(（]\d+[\)）]", RegexOptions.Compiled);
    private static readonly Regex 文号正则 = new(@"[〔〔\[]?\d{4}[〕\]）)]?.*号$", RegexOptions.Compiled);
    private static readonly Regex 日期正则 = new(@"^\d{4}年\d{1,2}月\d{1,2}日$", RegexOptions.Compiled);
    private static readonly Regex 附件正则 = new(@"^附件[：:\s]", RegexOptions.Compiled);
    private static readonly string[] 版记关键词 = ["抄送", "印发", "印送", "分送"];
    private static readonly string[] 主送机关排除词 = ["标题", "通知", "决定", "纪要", "请示", "报告", "附件", "抄送", "印发", "日期"];
    private static readonly string[] 公文标题关键词 = ["通知", "通报", "决定", "请示", "报告", "意见", "公告", "通告", "纪要", "函", "令"];

    public GovDocumentStructure 分析(MainDocumentPart mainPart)
    {
        var body = mainPart.Document?.Body ?? throw new InvalidOperationException("文档没有正文。");
        var paragraphs = body.Elements<Paragraph>()
            .Where(x => !x.Ancestors<Table>().Any())
            .ToList();

        var structure = new GovDocumentStructure();
        if (paragraphs.Count == 0)
            return structure;

        structure.是否含页码字段 = 检测页码字段(mainPart);
        识别横向节(body, structure);

        for (var i = 0; i < Math.Min(paragraphs.Count, 20); i++)
        {
            if (是文号(获取段落可见文本(paragraphs[i])))
            {
                structure.文号段索引.Add(i);
                break;
            }
        }

        var 标题起点 = structure.文号段索引.Count > 0 ? structure.文号段索引.Max() + 1 : 0;
        var 已发现标题 = false;
        for (var i = 标题起点; i < Math.Min(paragraphs.Count, 标题起点 + 12); i++)
        {
            var text = 获取段落可见文本(paragraphs[i]);
            if (string.IsNullOrWhiteSpace(text))
                continue;

            var 允许版式特征 = structure.文号段索引.Count > 0;
            if (!疑似标题段(mainPart, paragraphs[i], text, 允许版式特征))
            {
                if (已发现标题)
                    break;
                continue;
            }

            structure.标题段索引.Add(i);
            已发现标题 = true;
            if (structure.标题段索引.Count >= 3)
                break;
        }

        if (structure.标题段索引.Count > 0)
        {
            var recipientStart = structure.标题段索引.Max() + 1;
            for (var i = recipientStart; i < Math.Min(paragraphs.Count, recipientStart + 6); i++)
            {
                var text = 获取段落可见文本(paragraphs[i]);
                if (string.IsNullOrWhiteSpace(text))
                    continue;

                if (是主送机关(text))
                    structure.主送机关段索引.Add(i);
                break;
            }
        }

        structure.文种结果 = 检测文种(null);

        for (var i = paragraphs.Count - 1; i >= Math.Max(0, paragraphs.Count - 20); i--)
        {
            var text = 获取段落可见文本(paragraphs[i]);
            if (string.IsNullOrWhiteSpace(text))
                continue;

            if (是日期(text))
            {
                structure.日期段索引.Add(i);
                break;
            }
        }

        for (var i = 0; i < paragraphs.Count; i++)
        {
            if (structure.标题段索引.Contains(i) || structure.文号段索引.Contains(i) || structure.主送机关段索引.Contains(i))
                continue;

            var level = 解析段落标题级别(mainPart, paragraphs[i]);
            if (level > 0)
            {
                structure.标题级别映射[i] = level;
            }

            if (是附件说明(text: 获取段落可见文本(paragraphs[i])))
            {
                structure.附件段索引.Add(i);
            }

            if (是版记说明(获取段落可见文本(paragraphs[i])))
            {
                structure.版记段索引.Add(i);
            }
        }

        return structure;
    }

    public static GovDocumentKindResult 检测文种(string? text)
    {
        return new GovDocumentKindResult(GovDocumentKind.普通公文, "本版本仅处理一般普通公文。");
    }

    public static int 检测标题级别(string? text)
    {
        var normalized = text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(normalized))
            return 0;

        if (一级标题正则.IsMatch(normalized)) return 1;
        if (二级标题正则.IsMatch(normalized)) return 2;
        if (三级标题正则.IsMatch(normalized)) return 3;
        if (四级标题正则.IsMatch(normalized)) return 4;
        return 0;
    }

    private static bool 疑似标题段(MainDocumentPart mainPart, Paragraph paragraph, string text, bool 允许版式特征)
    {
        if (text.Length > 60 || text.Contains('：') || text.EndsWith("号", StringComparison.Ordinal))
            return false;

        var styledLevel = GovOpenXmlHelper.解析样式标题级别(mainPart, paragraph);
        if (允许版式特征 && styledLevel > 0)
            return true;

        var outlineLevel = GovOpenXmlHelper.解析大纲级别(paragraph);
        if (允许版式特征 && outlineLevel > 0)
            return true;

        var justification = paragraph.ParagraphProperties?.Justification?.Val?.Value;
        if (允许版式特征 && justification == JustificationValues.Center)
            return true;

        if (text.Length <= 40 &&
            !text.Contains('：') &&
            !text.Contains(':') &&
            包含任一关键词(text, 公文标题关键词))
            return true;

        return 允许版式特征 && paragraph.Descendants<RunProperties>()
            .Any(x => int.TryParse(x.FontSize?.Val, out var size) && size >= 36);
    }

    private static int 解析段落标题级别(MainDocumentPart mainPart, Paragraph paragraph)
    {
        var byStyle = GovOpenXmlHelper.解析样式标题级别(mainPart, paragraph);
        if (byStyle > 0)
            return byStyle;

        var byOutline = GovOpenXmlHelper.解析大纲级别(paragraph);
        if (byOutline > 0)
            return byOutline;

        return 检测标题级别(获取段落可见文本(paragraph));
    }

    private static bool 是文号(string text)
    {
        return text.Length <= 40 && text.Contains('号') && 文号正则.IsMatch(text);
    }

    private static bool 是主送机关(string text)
    {
        if (text.Length is 0 or > 30)
            return false;

        if (!text.EndsWith('：') && !text.EndsWith(':'))
            return false;

        return 主送机关排除词.All(x => !text.Contains(x, StringComparison.Ordinal));
    }

    private static bool 是日期(string text)
    {
        return 日期正则.IsMatch(text.Replace(" ", string.Empty));
    }

    private static bool 是附件说明(string text)
    {
        return 附件正则.IsMatch(text);
    }

    private static bool 是版记说明(string text)
    {
        return !string.IsNullOrWhiteSpace(text) && 版记关键词.Any(x => text.StartsWith(x, StringComparison.Ordinal));
    }

    private static bool 检测页码字段(MainDocumentPart mainPart)
    {
        foreach (var footerPart in mainPart.FooterParts)
        {
            if (footerPart.Footer?.Descendants<FieldCode>().Any(x =>
                    (x.Text ?? string.Empty).Contains("PAGE", StringComparison.OrdinalIgnoreCase)) == true)
                return true;

            if (footerPart.Footer?.Descendants<SimpleField>().Any(x =>
                    (x.Instruction?.Value ?? string.Empty).Contains("PAGE", StringComparison.OrdinalIgnoreCase)) == true)
                return true;
        }

        return false;
    }

    private static void 识别横向节(Body body, GovDocumentStructure structure)
    {
        var sections = body.Descendants<SectionProperties>().ToList();
        for (var i = 0; i < sections.Count; i++)
        {
            var pageSize = sections[i].GetFirstChild<PageSize>();
            var isLandscape = pageSize?.Orient?.Value == PageOrientationValues.Landscape ||
                              pageSize?.Width?.Value > pageSize?.Height?.Value;
            if (isLandscape)
            {
                structure.横向节索引.Add(i);
            }
        }
    }

    private static string 获取段落可见文本(Paragraph paragraph)
    {
        return GovOpenXmlHelper.提取归一化可见文本(paragraph);
    }

    private static bool 包含任一关键词(string text, IEnumerable<string> keywords)
    {
        return keywords.Any(x => text.Contains(x, StringComparison.Ordinal));
    }
}

public sealed class GovDocumentStructure
{
    public HashSet<int> 标题段索引 { get; } = [];

    public HashSet<int> 文号段索引 { get; } = [];

    public HashSet<int> 主送机关段索引 { get; } = [];

    public HashSet<int> 日期段索引 { get; } = [];

    public HashSet<int> 附件段索引 { get; } = [];

    public HashSet<int> 横向节索引 { get; } = [];

    public HashSet<int> 版记段索引 { get; } = [];

    public Dictionary<int, int> 标题级别映射 { get; } = [];

    public GovDocumentKindResult 文种结果 { get; set; } = new(GovDocumentKind.普通公文, "未分析文种。");

    public bool 是否含页码字段 { get; set; }
}

public enum GovDocumentKind { 普通公文 }

public sealed record GovDocumentKindResult(GovDocumentKind Kind, string Reason);
