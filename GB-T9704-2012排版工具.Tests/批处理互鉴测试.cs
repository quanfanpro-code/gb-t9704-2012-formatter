using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using GBT9704_2012排版工具;
using GBT9704_2012排版工具.Contracts;
using GBT9704_2012排版工具.Core;

namespace GB_T9704_2012排版工具.Tests;

public sealed class 批处理互鉴测试
{
    [Theory]
    [InlineData("报告_GB9704-2012排版(2).docx", true)]
    [InlineData("报告_GB9704-2012排版_0123456789abcdef0123456789abcdef.docx", true)]
    [InlineData("报告_GB9704-2012排版_失败_20260912_123456_0123456789abcdef0123456789abcdef.docx", true)]
    [InlineData(".报告.0123456789abcdef0123456789abcdef.tmp.docx", true)]
    [InlineData("~$报告.docx", true)]
    [InlineData("报告_GB9704-2012排版.docx", true)]
    [InlineData("报告_GB9704-2012排版说明.docx", false)]
    [InlineData("失败原因说明.docx", false)]
    public void 生成件不应再次处理且不误伤原稿(string name, bool expected)
        => Assert.Equal(expected, 输出文件命名规则.是已排版文件(name));

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void 取消不发布成稿并保留原稿(bool during)
    {
        var dir = Path.Combine(Path.GetTempPath(), "公文取消_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var input = Path.Combine(dir, "原稿.docx");
        using (var doc = WordprocessingDocument.Create(input, WordprocessingDocumentType.Document))
            doc.AddMainDocumentPart().Document = new Document(new Body(new Paragraph(new Run(new Text("关于核对材料的通知"))), new Paragraph(new Run(new Text("应保留的正文。")))));
        var original = File.ReadAllBytes(input);
        using var cts = new CancellationTokenSource();
        var request = new RequestContract { InputPath = input, OutputPath = Path.Combine(dir, "成稿.docx") };
        var tokenProperty = typeof(RequestContract).GetProperty("CancellationToken");
        Assert.NotNull(tokenProperty);
        tokenProperty.SetValue(request, cts.Token);
        if (!during) cts.Cancel();
        var result = new GovDocumentPipeline(() => { if (during) cts.Cancel(); return Array.Empty<string>(); }).Process(request);
        Assert.False(result.Success);
        Assert.Equal("CANCELLED", result.ErrorCode);
        Assert.False(File.Exists(Path.Combine(dir, "成稿.docx")));
        Assert.Equal(original, File.ReadAllBytes(input));
    }
}
