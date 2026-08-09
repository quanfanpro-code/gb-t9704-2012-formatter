namespace GBT9704_2012排版工具.Contracts;

public sealed record BatchAuditItemContract(
    string InputPath,
    string? OutputPath,
    int AttachmentCount,
    bool HasPageNumberField,
    int LandscapeSectionCount,
    int ImprintCount,
    string GateResult,
    string? FailureStage,
    string? RuleCode,
    string? RuleName,
    bool NeedsManualReview,
    string? Message,
    string? DocumentKind,
    string? AuditPath
);
