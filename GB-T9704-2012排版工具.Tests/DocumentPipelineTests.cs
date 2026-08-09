using System.IO;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using GBT9704_2012排版工具.Contracts;
using GBT9704_2012排版工具.Core;

namespace GB_T9704_2012排版工具.Tests;

public sealed class 文档流水线测试
{
    [Fact]
    public void 表格数字和尾部空行必须保持()
    {
        var input = 创建样本文档(body =>
        {
            添加标准版头(body);
            body.Append(创建表格(
                new TableRow(new TableCell(new Paragraph(new Run(new Text("1234"))))),
                new TableRow(new TableCell(new Paragraph(new Run(new Text(string.Empty)))))));
        });

        var result = 处理(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        var table = Assert.Single(获取正文(output).Elements<Table>());
        Assert.Equal(2, table.Elements<TableRow>().Count());
        Assert.Equal("1234", table.Elements<TableRow>().First().InnerText);
    }

    [Fact]
    public void 换行分页和段前分页必须保持()
    {
        var input = 创建样本文档(body =>
        {
            添加标准版头(body);
            body.Append(new Paragraph(
                new ParagraphProperties(new PageBreakBefore()),
                new Run(new Text("第一行"), new Break(), new Text("第二行"), new Break { Type = BreakValues.Page })));
        });

        var result = 处理(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        var paragraph = 获取正文(output).Elements<Paragraph>()
            .Single(x => x.InnerText.Contains("第一行", StringComparison.Ordinal));
        Assert.NotNull(paragraph.ParagraphProperties?.PageBreakBefore);
        Assert.Equal(2, paragraph.Descendants<Break>().Count());
        Assert.Contains(paragraph.Descendants<Break>(), x => x.Type?.Value == BreakValues.Page);
    }

    [Fact]
    public void 机关标志不得被误认为标题()
    {
        var input = 创建样本文档(body => 添加标准版头(body));

        var result = 处理(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        var paragraphs = 获取正文(output).Elements<Paragraph>().ToList();
        var agencyRun = paragraphs.Single(x => x.InnerText == "某某市人民政府文件").GetFirstChild<Run>()!;
        var number = paragraphs.Single(x => x.InnerText == "某政发〔2026〕1号");
        var titleRun = paragraphs.Single(x => x.InnerText == "关于开展专项工作的通知").GetFirstChild<Run>()!;

        Assert.Equal("FF0000", agencyRun.RunProperties?.Color?.Val?.Value);
        Assert.Equal("36", agencyRun.RunProperties?.FontSize?.Val?.Value);
        Assert.Equal(JustificationValues.Center, number.ParagraphProperties?.Justification?.Val?.Value);
        Assert.Equal("44", titleRun.RunProperties?.FontSize?.Val?.Value);
    }

    [Fact]
    public void 页眉必须保留且所有旧页脚必须替换为标准页码()
    {
        var input = 创建样本文档(body => 添加标准版头(body), addHeaderAndFooters: true);

        var result = 处理(input);

        Assert.True(result.Success, result.Message);
        using var output = WordprocessingDocument.Open(result.OutputPath, false);
        var main = output.MainDocumentPart!;
        Assert.Contains(main.HeaderParts, x => x.Header?.InnerText.Contains("保留页眉", StringComparison.Ordinal) == true);
        Assert.DoesNotContain(main.FooterParts, x => x.Footer?.InnerText.Contains("旧页脚业务文字", StringComparison.Ordinal) == true);
        Assert.All(main.FooterParts, footer =>
        {
            var footerXml = Assert.IsType<Footer>(footer.Footer);
            Assert.Contains(footerXml.Descendants<FieldCode>(), x => x.Text.Contains("PAGE", StringComparison.OrdinalIgnoreCase));
            Assert.All(footerXml.Descendants<RunFonts>(), fonts => Assert.Equal("宋体", fonts.Ascii?.Value));
            Assert.All(footerXml.Descendants<FontSize>(), size => Assert.Equal("28", size.Val?.Value));
        });
    }

    [Fact]
    public void 最终输出必须通过OpenXml结构校验且审计反映最终页码()
    {
        var input = 创建样本文档(body =>
        {
            添加标准版头(body);
            body.Append(创建表格(new TableRow(new TableCell(new Paragraph(new Run(new Text("内容")))))));
        });

        var result = 处理(input);

        Assert.True(result.Success, result.Message);
        Assert.True(result.HasPageNumberField);
        using (var output = WordprocessingDocument.Open(result.OutputPath, false))
        {
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(output));
        }

        Assert.NotNull(result.AuditPath);
        using var json = JsonDocument.Parse(File.ReadAllText(result.AuditPath!));
        Assert.True(json.RootElement.GetProperty("HasPageNumberField").GetBoolean());
        Assert.Equal("通过", json.RootElement.GetProperty("GateResult").GetString());
    }

    [Fact]
    public void 已有输出文件不得被覆盖()
    {
        var input = 创建样本文档(body => 添加标准版头(body));
        var existing = Path.Combine(Path.GetDirectoryName(input)!, "已有结果.docx");
        File.WriteAllText(existing, "不得覆盖");

        var result = new GovDocumentPipeline().Process(new RequestContract
        {
            InputPath = input,
            OutputPath = existing
        });

        Assert.False(result.Success);
        Assert.Equal("不得覆盖", File.ReadAllText(existing));
    }

    [Fact]
    public void OpenXml结构错误必须成为阻断错误()
    {
        var path = 创建样本文档(body => body.Append(
            new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("无表格定义"))))))));
        var service = new GovValidationService();

        using var document = WordprocessingDocument.Open(path, false);
        service.验证(document);

        Assert.True(service.HasBlockingErrors);
        Assert.Contains(service.Errors, x => x.Contains("OpenXML 结构错误", StringComparison.Ordinal));
    }

    [Fact]
    public void 批量审计优先写入指定输出目录并使用UTF8BOM()
    {
        var directory = Path.Combine(Path.GetTempPath(), "GB-T9704-2012-tests", Guid.NewGuid().ToString("N"));
        var path = new GovAuditService().写入批量审计清单("不存在的输入目录", [], directory);
        var bytes = File.ReadAllBytes(path);

        Assert.Equal(directory, Path.GetDirectoryName(path));
        Assert.True(bytes.Length >= 3);
        Assert.Equal(new byte[] { 0xEF, 0xBB, 0xBF }, bytes[..3]);
    }

    private static ResponseContract 处理(string input)
    {
        return new GovDocumentPipeline().Process(new RequestContract { InputPath = input });
    }

    private static string 创建样本文档(Action<Body> buildBody, bool addHeaderAndFooters = false)
    {
        var directory = Path.Combine(Path.GetTempPath(), "GB-T9704-2012-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "输入.docx");

        using var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        main.Document = new Document();
        var body = new Body();
        main.Document.Append(body);
        buildBody(body);

        var section = new SectionProperties(
            new PageSize { Width = 11906, Height = 16838 },
            new PageMargin());

        if (addHeaderAndFooters)
        {
            var header = main.AddNewPart<HeaderPart>();
            header.Header = new Header(new Paragraph(new Run(new Text("保留页眉"))));
            section.PrependChild(new HeaderReference
            {
                Type = HeaderFooterValues.Default,
                Id = main.GetIdOfPart(header)
            });

            foreach (var type in new[] { HeaderFooterValues.First, HeaderFooterValues.Default, HeaderFooterValues.Even })
            {
                var footer = main.AddNewPart<FooterPart>();
                footer.Footer = new Footer(new Paragraph(new Run(new Text("旧页脚业务文字"))));
                section.InsertBefore(new FooterReference
                {
                    Type = type,
                    Id = main.GetIdOfPart(footer)
                }, section.GetFirstChild<PageSize>());
            }
        }

        body.Append(section);
        main.Document.Save();
        return path;
    }

    private static void 添加标准版头(Body body)
    {
        body.Append(
            居中段落("某某市人民政府文件", new RunProperties(
                new RunFonts { EastAsia = "方正小标宋简体" },
                new Color { Val = "FF0000" },
                new FontSize { Val = "36" })),
            居中段落("某政发〔2026〕1号"),
            居中段落("关于开展专项工作的通知"),
            new Paragraph(new Run(new Text("各有关单位："))),
            new Paragraph(new Run(new Text("这是正文。"))));
    }

    private static Paragraph 居中段落(string text, RunProperties? properties = null)
    {
        var run = new Run(new Text(text));
        if (properties != null)
            run.PrependChild(properties);

        return new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
            run);
    }

    private static Body 获取正文(WordprocessingDocument document)
    {
        return document.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("测试输出缺少正文。");
    }

    private static Table 创建表格(params TableRow[] rows)
    {
        var table = new Table(
            new TableProperties(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Dxa }),
            new TableGrid(new GridColumn { Width = "5000" }));
        table.Append(rows);
        return table;
    }
}
