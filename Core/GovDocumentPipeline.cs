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
        var workingPath = string.Empty;

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

            if (File.Exists(outputPath))
                return ResponseContract.Fail($"输出文件已存在，未执行覆盖：{outputPath}", errorCode: "OUTPUT_EXISTS");

            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
            workingPath = Path.Combine(
                Path.GetDirectoryName(outputPath)!,
                $".{Path.GetFileNameWithoutExtension(outputPath)}.{Guid.NewGuid():N}.tmp.docx");
            File.Copy(request.InputPath, workingPath, overwrite: false);
            // File.Copy 会带走源文件的只读属性，导致下一行以可写方式打开时抛 UnauthorizedAccessException
            File.SetAttributes(workingPath, FileAttributes.Normal);

            using (var document = WordprocessingDocument.Open(workingPath, true))
            {
                var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("文档缺少主部件。");
                var body = mainPart.Document?.Body ?? throw new InvalidOperationException("文档缺少正文。");

                var structure = _structureAnalyzer.分析(mainPart);
                currentStructure = structure;
                currentKind = structure.文种结果;

                _paragraphService.格式化(body, structure);
                _tableService.格式化(body);
                _headerFooterService.格式化(document);
                mainPart.Document.Save();
            }

            GovDocumentStructure finalStructure;
            try
            {
                using var finalDocument = WordprocessingDocument.Open(workingPath, false);
                var finalMainPart = finalDocument.MainDocumentPart ?? throw new InvalidOperationException("输出文档缺少主部件。");
                finalStructure = _structureAnalyzer.分析(finalMainPart);
                _validationService.验证(finalDocument);
            }
            catch (Exception ex)
            {
                return 构建失败结果(request, workingPath, currentStructure, currentKind, $"输出文件无法重新打开：{ex.Message}", "OUTPUT_REOPEN", "最终校验");
            }

            if (_validationService.HasBlockingErrors)
            {
                var message = string.Join("\n", _validationService.Errors);
                return 构建失败结果(request, workingPath, finalStructure, finalStructure.文种结果, message, "OPENXML_SCHEMA", "最终校验");
            }

            File.Move(workingPath, outputPath, overwrite: false);
            workingPath = string.Empty;
            request.OutputPath = outputPath;
            var 校验警告列表 = new List<string>(_validationService.Warnings);
            if (_paragraphService.跳过不安全段落数 > 0)
                校验警告列表.Add($"跳过 {_paragraphService.跳过不安全段落数} 个含复杂结构的段落，内容和版式保持原样，请人工确认。");
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
            var result = 校验警告 != null
                ? success with { Message = 校验警告, NeedsManualReview = true }
                : success;
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
            return 构建失败结果(request, workingPath, currentStructure, currentKind, ex.Message, "PROCESS_FAILED", "排版处理");
        }
    }

    private ResponseContract 构建失败结果(
        RequestContract request,
        string workingPath,
        GovDocumentStructure? structure,
        GovDocumentKindResult? kind,
        string message,
        string errorCode,
        string failureStage)
    {
        if (!string.IsNullOrWhiteSpace(workingPath) && File.Exists(workingPath))
        {
            try
            {
                var directory = Path.GetDirectoryName(request.OutputPath) ?? Path.GetDirectoryName(workingPath)!;
                var name = Path.GetFileNameWithoutExtension(request.OutputPath);
                var failedPath = Path.Combine(directory, $"{name}_失败_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.docx");
                File.Move(workingPath, failedPath, overwrite: false);
                workingPath = failedPath;
                message += $"\n失败文件已保留：{failedPath}";
            }
            catch (Exception ex)
            {
                message += $"\n中间文件保留在：{workingPath}；改名失败：{ex.Message}";
            }
        }

        request.OutputPath = string.IsNullOrWhiteSpace(workingPath) ? request.OutputPath : workingPath;
        var fail = ResponseContract.Fail(
            message,
            errorCode: errorCode,
            failureStage: failureStage,
            ruleCode: errorCode == "OPENXML_SCHEMA" ? "OPENXML_SCHEMA" : null,
            ruleName: errorCode == "OPENXML_SCHEMA" ? "OpenXML结构合法性" : null,
            needsManualReview: true,
            documentKind: kind?.Kind.ToString(),
            documentKindReason: kind?.Reason,
            attachmentCount: structure?.附件段索引.Count ?? 0,
            hasPageNumberField: structure?.是否含页码字段 ?? false,
            landscapeSectionCount: structure?.横向节索引.Count ?? 0,
            imprintCount: structure?.版记段索引.Count ?? 0);
        var auditPath = 尝试写入单文件审计(request, fail, structure);
        return fail with { AuditPath = auditPath };
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
