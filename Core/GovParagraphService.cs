using DocumentFormat.OpenXml.Wordprocessing;

namespace GBT9704_2012排版工具.Core;

public sealed class GovParagraphService
{
    private const int 固定行距 = 578;
    private const string 标题字体 = "方正小标宋简体";
    private const string 一级标题字体 = "黑体";
    private const string 二级标题字体 = "楷体_GB2312";
    private const string 正文字体 = "仿宋_GB2312";
    private const string 西文字体 = "Times New Roman";
    private const string 标题字号 = "44";
    private const string 正文字号 = "32";

    /// <summary>
    /// 上一轮 格式化 中因含高风险结构（公章图片/超链接/书签/批注/域/内容控件等）而跳过文本重写的段落数。
    /// 这些段落仍做了行距/缩进/对齐等段落属性级调整，调用方应将该计数呈现给用户以便人工复核。
    /// </summary>
    public int 跳过不安全段落数 { get; private set; }

    public void 格式化(Body body, GovDocumentStructure structure)
    {
        跳过不安全段落数 = 0;

        var paragraphs = body.Elements<Paragraph>()
            .Where(x => !x.Ancestors<Table>().Any())
            .ToList();

        for (var i = 0; i < paragraphs.Count; i++)
        {
            var paragraph = paragraphs[i];
            var text = GovOpenXmlHelper.提取可见文本(paragraph).Trim();
            var isEmpty = string.IsNullOrWhiteSpace(text);

            if (structure.文号段索引.Count > 0 && i < structure.文号段索引.Min())
                continue;

            if (isEmpty)
            {
                continue;
            }

            // 防护门：含 Drawing/Hyperlink/书签/批注/域/SdtRun 等结构的段落禁止整体重写文本，
            // 否则成文日期段锚定的公章图片等内容会被静默抹平。此类段落只做上面的 pPr 级调整。
            var 可重写文本 = 结构安全分析器.段落可安全重写(paragraph);
            if (!可重写文本)
            {
                跳过不安全段落数++;
                System.Diagnostics.Trace.WriteLine(
                    $"[GovParagraphService] 第 {i} 段含高风险结构，已跳过文本重写：{text}");
                continue;
            }

            GovOpenXmlHelper.设置段落行距(paragraph, 固定行距);

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
        GovOpenXmlHelper.设置段落缩进(paragraph, leftChars: 2);
        var pPr = GovOpenXmlHelper.确保段落属性(paragraph);
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
        格式化运行(paragraph, font, 西文字体, 正文字号, bold);
    }

    private static void 格式化正文(Paragraph paragraph)
    {
        GovOpenXmlHelper.设置段落缩进(paragraph, firstLineChars: 2);
        var pPr = GovOpenXmlHelper.确保段落属性(paragraph);
        pPr.Justification = new Justification { Val = JustificationValues.Both };
        格式化运行(paragraph, 正文字体, 西文字体, 正文字号);
    }

    private static void 格式化运行(Paragraph paragraph, string 中文字体, string latinFont, string fontSize, bool bold = false)
    {
        foreach (var run in GovOpenXmlHelper.获取文本运行(paragraph))
            GovOpenXmlHelper.设置运行格式(run, 中文字体, latinFont, fontSize, bold);
    }
}
