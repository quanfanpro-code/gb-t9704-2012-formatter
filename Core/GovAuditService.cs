using System.IO;
using System.Text.Json;
using GBT9704_2012排版工具.Contracts;

namespace GBT9704_2012排版工具.Core;

public sealed class GovAuditService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        // 不转义中文等非 ASCII 字符，审计 JSON 保持可读
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    public string 写入单文件审计(RequestContract request, ResponseContract response, GovDocumentStructure? structure = null)
    {
        var outputPath = string.IsNullOrWhiteSpace(response.OutputPath)
            ? request.OutputPath
            : response.OutputPath;
        var auditPath = $"{outputPath}.audit.json";
        var documentKind = 解析文种(response.DocumentKind);
        var rules = GovRuleCatalog.GetApplicableRules(documentKind);
        var matchedRules = 获取命中规则(structure);
        var summary = new AuditSummaryContract(
            request.InputPath,
            outputPath,
            response.DocumentKind,
            response.DocumentKindReason,
            structure?.附件段索引.Count ?? 0,
            structure?.是否含页码字段 ?? false,
            structure?.横向节索引.Count ?? 0,
            structure?.版记段索引.Count ?? 0,
            rules.Select(x => x.Code).ToList(),
            rules.Select(x => x.Name).ToList(),
            matchedRules.Select(x => x.Code).ToList(),
            matchedRules.Select(x => x.Name).ToList(),
            response.Success ? "通过" : "阻断",
            response.FailureStage,
            response.RuleCode,
            response.RuleName,
            response.NeedsManualReview,
            response.Message);

        var json = JsonSerializer.Serialize(summary, JsonOptions);
        File.WriteAllText(auditPath, json);
        return auditPath;
    }

    public string 写入批量审计清单(string rootPath, IReadOnlyList<BatchAuditItemContract> items)
    {
        // 优先写到输入目录；rootPath 不是目录时退到其所在目录（即输入文件所在目录）
        var directoryPath = Directory.Exists(rootPath)
            ? rootPath
            : Path.GetDirectoryName(rootPath);
        if (string.IsNullOrEmpty(directoryPath) || !Directory.Exists(directoryPath))
        {
            // 输入目录也不可用时退到用户本地数据目录，绝不落进 CWD（快捷方式启动时 CWD 不可预测）
            directoryPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "GB-T9704-2012排版工具", "审计");
            Directory.CreateDirectory(directoryPath);
        }
        // 时间戳精确到毫秒，避免同一秒内两批互相覆盖
        var filePath = Path.Combine(
            directoryPath,
            $"GB-T9704-2012批量审计清单_{DateTime.Now:yyyyMMdd_HHmmss_fff}.json");
        var summary = new BatchAuditSummaryContract(
            rootPath,
            items.Count,
            items.Count(x => x.GateResult == "通过"),
            items.Count(x => x.GateResult != "通过"),
            items);

        var json = JsonSerializer.Serialize(summary, JsonOptions);
        File.WriteAllText(filePath, json);
        return filePath;
    }

    private static GovDocumentKind? 解析文种(string? text)
    {
        return text switch
        {
            nameof(GovDocumentKind.普通公文) => GovDocumentKind.普通公文,
            nameof(GovDocumentKind.信函) => GovDocumentKind.信函,
            nameof(GovDocumentKind.命令) => GovDocumentKind.命令,
            nameof(GovDocumentKind.纪要) => GovDocumentKind.纪要,
            _ => null
        };
    }

    private static IReadOnlyList<GovRuleDefinition> 获取命中规则(GovDocumentStructure? structure)
    {
        var results = new List<GovRuleDefinition>();
        if (structure == null)
            return results;

        if (structure.附件段索引.Count > 0)
            results.Add(GovRuleCatalog.附件标识);
        if (structure.是否含页码字段)
            results.Add(GovRuleCatalog.页码字段);
        if (structure.版记段索引.Count > 0)
            results.Add(GovRuleCatalog.版记区域);

        return results;
    }
}
