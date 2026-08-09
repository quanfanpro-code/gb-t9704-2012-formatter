using System.IO;
using DocumentFormat.OpenXml.Packaging;
using GBT9704_2012排版工具.Contracts;

namespace GBT9704_2012排版工具.Core;

public sealed class GovDocumentPipeline
{
    private readonly GovStructureAnalyzer _structureAnalyzer = new();
    private readonly GovParagraphService _paragraphService = new();
    private readonly GovTableService _tableService = new();
    private readonly GovHeaderFooterService _headerFooterService = new();
    private readonly GovValidationService _validationService = new();
    private readonly GovAuditService _auditService = new();

    public ResponseContract Process(RequestContract request)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));

        GovDocumentKindResult? currentKind = null;
        GovDocumentStructure? currentStructure = null;
        var outputPath = request.OutputPath;
        var 已创建输出 = false;

        if (!File.Exists(request.InputPath))
            return ResponseContract.Fail($"输入文件不存在：{request.InputPath}");

        try
        {
            outputPath = string.IsNullOrWhiteSpace(request.OutputPath)
                ? 输出文件命名规则.生成输出路径(request.InputPath)
                : request.OutputPath;

            // 输入输出同路径时复制等于自我覆盖，直接拒绝
            if (string.Equals(
                    Path.GetFullPath(request.InputPath),
                    Path.GetFullPath(outputPath),
                    StringComparison.OrdinalIgnoreCase))
                return ResponseContract.Fail("输入文件与输出文件路径相同，拒绝处理。");

            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
            File.Copy(request.InputPath, outputPath, overwrite: true);
            已创建输出 = true;
            // File.Copy 会带走源文件的只读属性，导致下一行以可写方式打开时抛 UnauthorizedAccessException
            File.SetAttributes(outputPath, FileAttributes.Normal);

            using (var document = WordprocessingDocument.Open(outputPath, true))
            {
                var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("文档缺少主部件。");
                var body = mainPart.Document.Body ?? throw new InvalidOperationException("文档缺少正文。");

                var structure = _structureAnalyzer.分析(mainPart);
                currentStructure = structure;
                if (request.PreferredDocumentKind.HasValue)
                {
                    structure.文种结果 = new GovDocumentKindResult(
                        request.PreferredDocumentKind.Value,
                        $"用户在界面中手动改选为“{request.PreferredDocumentKind.Value}”，覆盖自动识别结果。");
                }
                currentKind = structure.文种结果;

                _paragraphService.格式化(body, structure);
                _tableService.格式化(body);
                _headerFooterService.格式化(document);
                GovOpenXmlHelper.合并文档相邻运行(document);

                mainPart.Document.Save();
                _validationService.验证(document, outputPath, structure);
            }

            // using 块结束（Dispose 落盘完成）后再校验输出完整性，打不开视为写盘失败
            if (!GovValidationService.输出文件可正常打开(outputPath))
            {
                try { File.Delete(outputPath); } catch { /* 损坏文件删除失败，留给用户手动清理 */ }
                return ResponseContract.Fail("输出文件无法重新打开，疑似写盘失败");
            }

            request.OutputPath = outputPath;
            var finalStructure = currentStructure!;
            var 校验警告列表 = new List<string>(_validationService.Warnings);
            if (_paragraphService.跳过不安全段落数 > 0)
                校验警告列表.Add($"跳过 {_paragraphService.跳过不安全段落数} 个含图片/超链接/书签等复杂结构的段落：仅调整行距缩进，未重写文本。");
            if (_tableService.跳过不安全单元格数 > 0)
                校验警告列表.Add($"跳过 {_tableService.跳过不安全单元格数} 个含复杂结构的数字单元格：保留原文本未格式化。");
            var 校验警告 = 校验警告列表.Count > 0
                ? string.Join("\n", 校验警告列表)
                : null;
            var success = ResponseContract.Ok(
                outputPath,
                null,
                finalStructure.文种结果.Kind.ToString(),
                finalStructure.文种结果.Reason,
                attachmentCount: finalStructure.附件段索引.Count,
                hasPageNumberField: finalStructure.是否含页码字段,
                landscapeSectionCount: finalStructure.横向节索引.Count,
                imprintCount: finalStructure.版记段索引.Count);
            var result = 校验警告 != null ? success with { Message = 校验警告 } : success;
            var auditPath = 尝试写入单文件审计(request, result, finalStructure);
            if (auditPath == null)
            {
                // 审计写入失败只降级为警告，不翻转排版结果
                const string 审计警告 = "审计文件写入失败，不影响本次排版结果。";
                result = result with
                {
                    Message = string.IsNullOrEmpty(result.Message)
                        ? 审计警告
                        : result.Message + "\n" + 审计警告
                };
            }
            return result with { AuditPath = auditPath };
        }
        catch (Exception ex)
        {
            // 清理半成品：复制成功后中途失败会留下未排版副本，且命名规则会让它下次被误判为已排版而跳过
            if (已创建输出 && !string.IsNullOrEmpty(outputPath))
            {
                try
                {
                    File.Delete(outputPath);
                }
                catch (Exception deleteEx)
                {
                    System.Diagnostics.Debug.WriteLine($"半成品输出文件删除失败：{outputPath}，{deleteEx.Message}");
                }
            }

            var fail = ResponseContract.Fail(
                ex.Message,
                documentKind: currentKind?.Kind.ToString(),
                documentKindReason: currentKind?.Reason,
                attachmentCount: currentStructure?.附件段索引.Count ?? 0,
                hasPageNumberField: currentStructure?.是否含页码字段 ?? false,
                landscapeSectionCount: currentStructure?.横向节索引.Count ?? 0,
                imprintCount: currentStructure?.版记段索引.Count ?? 0);
            request.OutputPath = outputPath;
            var auditPath = 尝试写入单文件审计(request, fail, currentStructure);
            return fail with { AuditPath = auditPath };
        }
    }

    private string? 尝试写入单文件审计(RequestContract request, ResponseContract result, GovDocumentStructure? structure)
    {
        try
        {
            return _auditService.写入单文件审计(request, result, structure);
        }
        catch (Exception ex)
        {
            // 审计写入失败只降级为警告日志，绝不向外抛
            System.Diagnostics.Debug.WriteLine($"审计文件写入失败：{ex.Message}");
            return null;
        }
    }
}

