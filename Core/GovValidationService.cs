using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GBT9704_2012排版工具.Core;

public sealed class GovValidationService
{
    // 页面常量统一引用 GovPageConstants，避免与排版服务双份硬编码

    private readonly List<string> _warnings = [];

    /// <summary>校验过程中收集的所有警告。</summary>
    public IReadOnlyList<string> Warnings => _warnings;

    public void 验证(WordprocessingDocument document, string outputPath, GovDocumentStructure? structure = null)
    {
        _warnings.Clear();

        if (document.MainDocumentPart?.Document?.Body == null)
        {
            _warnings.Add("文档缺少正文。");
            return;
        }

        var body = document.MainDocumentPart.Document.Body;
        if (body.ChildElements.Count == 0)
        {
            _warnings.Add("文档正文为空。");
            return;
        }

        // Schema 校验（仅告警）
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        foreach (var err in validator.Validate(document).Take(10))
            _warnings.Add($"Schema: {err.Description}");

        var text = GovOpenXmlHelper.提取归一化可见文本(body);
        if (string.IsNullOrWhiteSpace(text))
            _warnings.Add("文档未识别到可用正文内容。");

        校验页面规则(body);

        // 注意：File.Copy 刚复制出的输出文件在 WordprocessingDocument Dispose 前尚未完全落盘，
        // 不能在此处做 File.Exists 判断；可打开性验证请用 输出文件可正常打开，在 Dispose 之后由 pipeline 调用。
    }

    private void 校验页面规则(Body body)
    {
        var sections = body.Descendants<SectionProperties>().ToList();
        if (sections.Count == 0)
        {
            _warnings.Add("页面参数：缺少节属性，无法验证页面参数。");
            return;
        }

        foreach (var section in sections)
        {
            校验页面尺寸(section.GetFirstChild<PageSize>());
            校验页边距(section.GetFirstChild<PageMargin>());
            校验版心行距栅格(section.GetFirstChild<DocGrid>());
        }
    }

    private void 校验页面尺寸(PageSize? pageSize)
    {
        if (pageSize?.Width?.Value == null || pageSize.Height?.Value == null)
        {
            _warnings.Add("页面尺寸：缺少页面尺寸。");
            return;
        }

        var width = pageSize.Width.Value;
        var height = pageSize.Height.Value;
        var isPortrait = pageSize.Orient?.Value != PageOrientationValues.Landscape;
        var valid = isPortrait
            ? width == GovPageConstants.A4宽度 && height == GovPageConstants.A4高度
            : width == GovPageConstants.A4高度 && height == GovPageConstants.A4宽度;

        if (!valid)
            _warnings.Add("页面尺寸：不符合 GB/T 9704—2012 A4 要求。");
    }

    private void 校验页边距(PageMargin? pageMargin)
    {
        if (pageMargin == null)
        {
            _warnings.Add("页边距：缺少页面边距。");
            return;
        }

        if (Math.Abs((int)(pageMargin.Top?.Value ?? 0) - GovPageConstants.上边距) > 1)
            _warnings.Add("页面上边距：不符合 GB/T 9704—2012 要求。");
        if (Math.Abs((int)(pageMargin.Bottom?.Value ?? 0) - GovPageConstants.下边距) > 1)
            _warnings.Add("页面下边距：不符合 GB/T 9704—2012 要求。");
        if (Math.Abs((int)(pageMargin.Left?.Value ?? 0) - GovPageConstants.左边距) > 1)
            _warnings.Add("页面左边距：不符合 GB/T 9704—2012 要求。");
        if (Math.Abs((int)(pageMargin.Right?.Value ?? 0) - GovPageConstants.右边距) > 1)
            _warnings.Add("页面右边距：不符合 GB/T 9704—2012 要求。");
        if (Math.Abs((int)(pageMargin.Header?.Value ?? 0) - GovPageConstants.页眉距离) > 1)
            _warnings.Add("页眉距离：不符合 GB/T 9704—2012 要求。");
        if (Math.Abs((int)(pageMargin.Footer?.Value ?? 0) - GovPageConstants.页脚距离) > 1)
            _warnings.Add("页脚距离：不符合 GB/T 9704—2012 要求。");
    }

    private void 校验版心行距栅格(DocGrid? docGrid)
    {
        if (docGrid?.LinePitch?.Value != GovPageConstants.正文行距)
            _warnings.Add("版心行距栅格：不符合 22 行/面要求。");
    }

    /// <summary>
    /// 试开输出文件验证可正常打开。供 pipeline 在 WordprocessingDocument Dispose 之后调用（此时内容才真正落盘）。
    /// </summary>
    public static bool 输出文件可正常打开(string outputPath)
    {
        try
        {
            using var document = WordprocessingDocument.Open(outputPath, false);
            return document.MainDocumentPart?.Document?.Body != null;
        }
        catch
        {
            return false;
        }
    }
}
