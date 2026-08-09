using GBT9704_2012排版工具;
using GBT9704_2012排版工具.Core;

namespace GB_T9704_2012排版工具.Tests;

public class 党政机关规则测试
{
    [Theory]
    [InlineData("关于印发XX实施方案的通知", GovDocumentKind.普通公文, "未命中特定格式特征，按普通公文处理。")]
    [InlineData("关于商洽XX事项的函", GovDocumentKind.信函, "命中“函/复函/商洽/答复”等信函特征词。")]
    [InlineData("国务院令第777号", GovDocumentKind.命令, "命中“令/命令/第×号”等命令文种特征。")]
    [InlineData("XX专题会议纪要", GovDocumentKind.纪要, "命中“纪要/会议纪要”等纪要特征词。")]
    public void 文种识别应命中预期文种(string text, GovDocumentKind expectedKind, string expectedReasonContains)
    {
        var result = GovStructureAnalyzer.检测文种(text);

        Assert.Equal(expectedKind, result.Kind);
        Assert.Contains(expectedReasonContains, result.Reason);
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

    [Theory]
    [InlineData("1234", false, "1,234.00")]
    [InlineData("-1234.5", false, "-1,234.50")]
    [InlineData("(1234.5)", false, "-1,234.50")]
    [InlineData("1,234.5", false, "1,234.50")]
    [InlineData("不是数字", false, null)]
    [InlineData("0123", false, null)]
    [InlineData("3", true, null)]
    [InlineData("123)", false, null)]
    [InlineData("(123", false, null)]
    public void 表格数字格式化应统一千分位和两位小数且保留受保护值(string rawText, bool isFirstColumn, string? expectedText)
    {
        var actual = GovTableService.格式化纯数字文本(rawText, isFirstColumn);

        Assert.Equal(expectedText, actual);
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

    [Theory]
    [InlineData(GovDocumentKind.普通公文)]
    [InlineData(GovDocumentKind.信函)]
    [InlineData(GovDocumentKind.命令)]
    [InlineData(GovDocumentKind.纪要)]
    public void 审计适用规则不应宣称尚未实现的文种专属校验(GovDocumentKind kind)
    {
        var codes = GovRuleCatalog.GetApplicableRules(kind).Select(rule => rule.Code).ToList();

        Assert.Contains("OPENXML_SCHEMA", codes);
        Assert.Contains("PAGE_SIZE", codes);
        Assert.DoesNotContain("DOC_NORMAL_HEADER", codes);
        Assert.DoesNotContain("DOC_LETTER_LAYOUT", codes);
        Assert.DoesNotContain("DOC_ORDER_SERIAL", codes);
        Assert.DoesNotContain("DOC_MEMO_TITLE", codes);
    }
}
