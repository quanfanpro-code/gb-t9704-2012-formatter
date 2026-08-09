using GBT9704_2012排版工具;
using GBT9704_2012排版工具.Core;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace GB_T9704_2012排版工具.Tests;

public class 党政机关规则测试
{
    [Fact]
    public void 程序只处理一般普通公文()
    {
        var result = GovStructureAnalyzer.检测文种("关于印发实施方案的通知");

        Assert.Equal(GovDocumentKind.普通公文, result.Kind);
        Assert.Contains("仅处理一般普通公文", result.Reason);
    }

    [Theory]
    [InlineData("一、总体要求", 1)]
    [InlineData("（一）工作目标", 2)]
    [InlineData("1．主要任务", 3)]
    [InlineData("（1）具体措施", 4)]
    [InlineData("这是普通正文", 0)]
    public void 标题级别识别应符合党政机关规则(string text, int expectedLevel)
    {
        var actual = GovStructureAnalyzer.检测标题级别(text);

        Assert.Equal(expectedLevel, actual);
    }

    [Fact]
    public void 输出文件命名应使用国标排版后缀()
    {
        var path = 输出文件命名规则.生成输出路径(@"C:\Temp\通知.docx");

        Assert.EndsWith("_GB9704-2012排版.docx", path, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void 已排版文件应被识别出来()
    {
        Assert.True(输出文件命名规则.是已排版文件(@"C:\Temp\通知_GB9704-2012排版.docx"));
        Assert.False(输出文件命名规则.是已排版文件(@"C:\Temp\通知.docx"));
    }

    [Fact]
    public void 审计规则只适用于一般普通公文()
    {
        var codes = GovRuleCatalog.GetApplicableRules(GovDocumentKind.普通公文).Select(rule => rule.Code).ToList();

        Assert.Contains("OPENXML_SCHEMA", codes);
        Assert.Contains("PAGE_SIZE", codes);
        Assert.DoesNotContain("DOC_NORMAL_HEADER", codes);
        Assert.DoesNotContain("DOC_LETTER_LAYOUT", codes);
        Assert.DoesNotContain("DOC_ORDER_SERIAL", codes);
        Assert.DoesNotContain("DOC_MEMO_TITLE", codes);
    }

    [Fact]
    public void 含脚注引用的段落不得自动改写()
    {
        var paragraph = new Paragraph(new Run(new Text("说明"), new FootnoteReference { Id = 1 }));

        Assert.False(结构安全分析器.段落可安全重写(paragraph));
    }

    [Fact]
    public void 含公式的段落不得自动改写()
    {
        var paragraph = new Paragraph(new M.OfficeMath(new M.Run(new M.Text("x=1"))));

        Assert.False(结构安全分析器.段落可安全重写(paragraph));
    }
}
