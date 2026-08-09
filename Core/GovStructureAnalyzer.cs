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
    // 关键词之间已剔除包含关系冗余项（“令”覆盖“命令”、“函”覆盖“复函”、“纪要”覆盖“会议纪要”）
    private static readonly string[] 公文标题关键词 = ["通知", "通报", "决定", "请示", "报告", "意见", "公告", "通告", "纪要", "函", "令"];
    private static readonly string[] 信函关键词 = ["函", "商洽", "答复"];
    private static readonly string[] 命令关键词 = ["令"];
    private static readonly string[] 纪要关键词 = ["纪要"];

    public GovDocumentStructure 分析(MainDocumentPart mainPart)
    {
        var body = mainPart.Document.Body ?? throw new InvalidOperationException("文档没有正文。");
        var paragraphs = body.Elements<Paragraph>()
            .Where(x => !x.Ancestors<Table>().Any())
            .ToList();

        var structure = new GovDocumentStructure();
        if (paragraphs.Count == 0)
            return structure;

        structure.是否含页码字段 = 检测页码字段(mainPart);
        识别横向节(body, structure);

        var titleCount = 0;
        for (var i = 0; i < Math.Min(paragraphs.Count, 12); i++)
        {
            var paragraph = paragraphs[i];
            var text = 获取段落可见文本(paragraph);
            if (string.IsNullOrWhiteSpace(text))
                continue;

            if (疑似标题段(mainPart, paragraph, text))
            {
                structure.标题段索引.Add(i);
                titleCount++;
                if (titleCount >= 3)
                    break;
                continue;
            }

            if (titleCount > 0)
                break;
        }

        if (structure.标题段索引.Count > 0)
        {
            var maxTitleIndex = structure.标题段索引.Max();
            for (var i = maxTitleIndex + 1; i < Math.Min(paragraphs.Count, maxTitleIndex + 8); i++)
            {
                var text = 获取段落可见文本(paragraphs[i]);
                if (string.IsNullOrWhiteSpace(text))
                    continue;

                if (是文号(text))
                {
                    structure.文号段索引.Add(i);
                    break;
                }
            }

            var recipientStart = structure.文号段索引.Count > 0
                ? structure.文号段索引.Max() + 1
                : maxTitleIndex + 1;

            for (var i = recipientStart; i < Math.Min(paragraphs.Count, recipientStart + 6); i++)
            {
                var text = 获取段落可见文本(paragraphs[i]);
                if (string.IsNullOrWhiteSpace(text))
                    continue;

                if (是主送机关(text))
                {
                    structure.主送机关段索引.Add(i);
                    break;
                }

                if (检测标题级别(text) > 0)
                    break;
            }
        }

        // 文种识别限定在标题区+文号区文本，避免正文出现“函”“令”等高频字时误判文种
        structure.文种结果 = 检测文种(构建文种判定文本(paragraphs, structure));

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
        var normalized = 归一化文种文本(text);
        if (string.IsNullOrWhiteSpace(normalized))
            return new GovDocumentKindResult(GovDocumentKind.普通公文, "未识别到有效文本，按普通公文兜底。");

        if (包含任一关键词(normalized, 纪要关键词))
            return new GovDocumentKindResult(GovDocumentKind.纪要, "命中“纪要/会议纪要”等纪要特征词。");

        if (normalized.Contains("国务院令", StringComparison.Ordinal) ||
            normalized.Contains("主席令", StringComparison.Ordinal) ||
            normalized.Contains("第", StringComparison.Ordinal) && normalized.Contains("号", StringComparison.Ordinal) &&
            包含任一关键词(normalized, 命令关键词))
            return new GovDocumentKindResult(GovDocumentKind.命令, "命中“令/命令/第×号”等命令文种特征。");

        if (包含任一关键词(normalized, 信函关键词))
            return new GovDocumentKindResult(GovDocumentKind.信函, "命中“函/复函/商洽/答复”等信函特征词。");

        return new GovDocumentKindResult(GovDocumentKind.普通公文, "未命中特定格式特征，按普通公文处理。");
    }

    /// <summary>
    /// 文种识别只取标题区与文号区文本，避免正文中的高频字（如“函”“令”）误判文种；
    /// 未识别到标题/文号时退化为前 12 段版头区文本（与标题扫描窗口一致）。
    /// </summary>
    private static string 构建文种判定文本(List<Paragraph> paragraphs, GovDocumentStructure structure)
    {
        var 候选索引 = structure.标题段索引.Concat(structure.文号段索引).Distinct().OrderBy(x => x).ToList();
        if (候选索引.Count > 0)
            return string.Join("\n", 候选索引.Select(i => 获取段落可见文本(paragraphs[i])));

        return string.Join("\n", paragraphs.Take(12)
            .Select(获取段落可见文本)
            .Where(x => !string.IsNullOrWhiteSpace(x)));
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

    private static bool 疑似标题段(MainDocumentPart mainPart, Paragraph paragraph, string text)
    {
        if (text.Length > 60 || text.Contains('：') || text.EndsWith("号", StringComparison.Ordinal))
            return false;

        var styledLevel = GovOpenXmlHelper.解析样式标题级别(mainPart, paragraph);
        if (styledLevel > 0)
            return true;

        var outlineLevel = GovOpenXmlHelper.解析大纲级别(paragraph);
        if (outlineLevel > 0)
            return true;

        var justification = paragraph.ParagraphProperties?.Justification?.Val?.Value;
        if (justification == JustificationValues.Center)
            return true;

        if (text.Length <= 40 &&
            !text.Contains('：') &&
            !text.Contains(':') &&
            包含任一关键词(text, 公文标题关键词))
            return true;

        return paragraph.Descendants<RunProperties>()
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

    private static string 归一化文种文本(string? text)
    {
        return text?.Replace("\r", "\n").Replace("　", " ").Trim() ?? string.Empty;
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

public enum GovDocumentKind
{
    普通公文,
    信函,
    命令,
    纪要
}

public sealed record GovDocumentKindResult(GovDocumentKind Kind, string Reason);

