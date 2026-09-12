using System.IO;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using GBT9704_2012排版工具.Contracts;
using GBT9704_2012排版工具.Core;

namespace GB_T9704_2012排版工具.Tests;

public sealed class 公文排版测试
{
    [Theory]
    [InlineData("某政发[2026]01号", "2026年9月12日", "文号")]
    [InlineData("某政发〔2026〕1号", "2025年2月29日", "日期")]
    public void 原文错误必须进入人工复核(string 文号, string 日期, string 问题)
    {
        var result = new GovDocumentPipeline().Process(new RequestContract
        {
            InputPath = 样本(文号, "关于开展检查的通知", "各单位：", "这是正文。", 日期)
        });
        Assert.True(result.Success, result.Message);
        Assert.True(result.NeedsManualReview);
        Assert.Contains(问题, result.Message);
    }

    [Fact]
    public void 有效闰日不得误报()
    {
        var result = new GovDocumentPipeline().Process(new RequestContract
        {
            InputPath = 样本("某政发〔2024〕1号", "关于开展检查的通知", "这是正文。", "2024年2月29日")
        });
        Assert.True(result.Success, result.Message);
        Assert.False(result.NeedsManualReview, result.Message);
    }

    [Fact]
    public void 附件说明与标签标题不一致必须提示()
    {
        var result = new GovDocumentPipeline().Process(new RequestContract
        {
            InputPath = 样本("关于开展检查的通知", "这是正文。", "附件：1.检查名单", "附件1", "另一份名单", "张三")
        });
        Assert.True(result.Success, result.Message);
        Assert.True(result.NeedsManualReview);
        Assert.Contains("附件", result.Message);
    }

    [Fact]
    public void 四级标题必须可供Word目录使用()
    {
        var result = new GovDocumentPipeline().Process(new RequestContract
        {
            InputPath = 样本("关于开展检查的通知", "一、总体要求", "（一）工作安排", "1.检查内容", "（1）执行步骤", "正文。")
        });
        Assert.True(result.Success, result.Message);
        using var doc = WordprocessingDocument.Open(result.OutputPath, false);
        var ps = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToArray();
        for (var i = 1; i <= 4; i++)
        {
            Assert.Equal(i - 1, ps[i].ParagraphProperties?.OutlineLevel?.Val?.Value);
            Assert.False(string.IsNullOrWhiteSpace(ps[i].ParagraphProperties?.ParagraphStyleId?.Val));
        }
    }

    [Fact]
    public void 检查记录不得冒充Word实测或全规则检查()
    {
        var result = new GovDocumentPipeline().Process(new RequestContract { InputPath = 样本("关于开展检查的通知", "正文。") });
        Assert.True(result.Success, result.Message);
        using var audit = JsonDocument.Parse(File.ReadAllText(result.AuditPath!));
        Assert.True(audit.RootElement.TryGetProperty("Word实测已执行", out var word));
        Assert.False(word.GetBoolean());
        Assert.DoesNotContain(audit.RootElement.GetProperty("CheckedRuleCodes").EnumerateArray(), x => x.GetString() == "IMPRINT");
    }

    [Theory]
    [InlineData(排版模式.电子红头)]
    [InlineData(排版模式.套打)]
    public void 正式模式生成正确版头并保持正文(排版模式 模式)
    {
        var input = 样本("关于开展检查的通知", "各单位：", "应保留的正文与金额1234.50。", "某市人民政府", "2026年9月12日");
        var result = new GovDocumentPipeline().Process(new RequestContract { InputPath = input, 模式 = 模式, 要素 = 要素() });
        Assert.True(result.Success, result.Message);
        using var doc = WordprocessingDocument.Open(result.OutputPath, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.Contains("应保留的正文与金额1234.50。", body.InnerText);
        Assert.Single(body.Descendants<Text>(), t => t.Text == "某政发〔2026〕1号");
        var red = body.Descendants<Color>().Any(c => c.Val?.Value == "FF0000");
        Assert.Equal(模式 == 排版模式.电子红头, red);
        Assert.Equal(模式 == 排版模式.电子红头, body.InnerText.Contains("某市人民政府文件"));
        Assert.Equal(1, body.Descendants<Text>().Count(t => t.Text == "关于开展检查的通知"));
    }

    [Fact]
    public void 复杂版头不得被拆除或叠加()
    {
        var input = 样本("某市人民政府文件", "某政发〔2026〕1号", "关于开展检查的通知", "正文。", "2026年9月12日");
        using (var doc = WordprocessingDocument.Open(input, true))
            doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First().Append(new BookmarkStart { Id = "1", Name = "保留版头" }, new BookmarkEnd { Id = "1" });
        var bytes = File.ReadAllBytes(input);
        var result = new GovDocumentPipeline().Process(new RequestContract { InputPath = input, 模式 = 排版模式.电子红头, 要素 = 要素() });
        Assert.False(result.Success);
        Assert.Contains("复杂", result.Message);
        Assert.Equal(bytes, File.ReadAllBytes(input));
    }

    [Fact]
    public void 附件另页且末页版记使用独立定位()
    {
        var input = 样本("关于开展检查的通知", "正文。", "附件：1.检查名单", "某市人民政府", "2026年9月12日", "附件1", "检查名单", "张三");
        var f = 要素(); f.抄送 = "抄送：有关单位。"; f.印发机关 = "某市人民政府办公室"; f.印发日期 = "2026年9月12日";
        var result = new GovDocumentPipeline().Process(new RequestContract { InputPath = input, 模式 = 排版模式.电子红头, 要素 = f });
        Assert.True(result.Success, result.Message);
        using var doc = WordprocessingDocument.Open(result.OutputPath, false);
        var body = doc.MainDocumentPart!.Document!.Body!;
        Assert.NotNull(body.Elements<Paragraph>().Single(p => p.InnerText == "附件1").ParagraphProperties?.PageBreakBefore);
        Assert.Contains(body.Descendants<TablePositionProperties>(), p => p.VerticalAnchor?.Value == VerticalAnchorValues.Margin);
    }

    [Fact]
    public void 普通材料长标题不得降为正文且主题字体不能覆盖指定字体()
    {
        var title = "关于开展全市有关部门年度重点项目资金使用情况以及资料完整性专项检查并进一步规范后续整改工作的通知";
        var input = 样本(title, "正文。");
        using (var doc = WordprocessingDocument.Open(input, true))
            doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First().GetFirstChild<Run>()!.PrependChild(
                new RunProperties(new RunFonts { EastAsiaTheme = ThemeFontValues.MajorEastAsia }));
        var result = new GovDocumentPipeline().Process(new RequestContract { InputPath = input });
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        var run = output.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First().GetFirstChild<Run>()!;
        Assert.Equal("44", run.RunProperties?.FontSize?.Val?.Value);
        Assert.Null(run.RunProperties?.RunFonts?.EastAsiaTheme);
    }

    [Fact]
    public void 附件说明续项不应识别为三级标题且各项不得拆开()
    {
        var input = 样本("关于开展检查的通知", "正文。", "附件：1.甲", "2.乙", "附件1", "甲", "内容。", "附件2", "乙", "内容。");
        var result = new GovDocumentPipeline().Process(new RequestContract { InputPath = input });
        using var doc = WordprocessingDocument.Open(result.OutputPath, false);
        var ps = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToArray();
        Assert.Null(ps.Single(p => p.InnerText == "2.乙").ParagraphProperties?.OutlineLevel);
        Assert.NotNull(ps.Single(p => p.InnerText == "附件：1.甲").ParagraphProperties?.KeepNext);
    }

    [Fact]
    public void 用户指定的标题换行应保留且不增加标题间空行()
    {
        var f = 要素(); f.标题 = "关于开展\n检查的通知";
        var input = 样本("关于开展检查的通知", "正文。", "2026年9月12日");
        var result = new GovDocumentPipeline().Process(new RequestContract { InputPath = input, 模式 = 排版模式.电子红头, 要素 = f });
        Assert.True(result.Success, result.Message);
        using var doc = WordprocessingDocument.Open(result.OutputPath, false);
        var ps = doc.MainDocumentPart!.Document!.Body!.Elements<Paragraph>();
        Assert.Contains(ps, p => p.InnerText == "关于开展检查的通知" && p.Descendants<Break>().Count() == 1);
    }

    [Fact]
    public void 缺失字体必须同时进入返回值与检查记录()
    {
        var result = new GovDocumentPipeline(() => ["仿宋_GB2312"]).Process(new RequestContract { InputPath = 样本("关于开展检查的通知", "正文。") });
        Assert.True(result.Success, result.Message);
        Assert.True(result.NeedsManualReview);
        Assert.Contains("仿宋_GB2312", result.Message);
        using var audit = JsonDocument.Parse(File.ReadAllText(result.AuditPath!));
        Assert.Equal("人工复核", audit.RootElement.GetProperty("GateResult").GetString());
        Assert.Contains("仿宋_GB2312", audit.RootElement.GetProperty("实际检查记录").ToString());
    }

    private static 公文要素 要素() => new()
    {
        机关标志 = "某市人民政府文件", 标题 = "关于开展检查的通知", 文号 = "某政发〔2026〕1号",
        主送机关 = "各单位：", 署名 = "某市人民政府", 成文日期 = "2026年9月12日", 已确认 = true
    };

    internal static string 样本(params string[] 文本)
    {
        var dir = Path.Combine(Path.GetTempPath(), "GB-T9704-2012-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "输入.docx");
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var body = new Body(文本.Select(x => new Paragraph(new Run(new Text(x)))));
        body.Append(new SectionProperties(new PageSize { Width = 11906, Height = 16838 }, new PageMargin()));
        main.Document = new Document(body);
        return path;
    }
}
