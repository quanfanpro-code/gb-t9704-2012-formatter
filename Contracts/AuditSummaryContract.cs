namespace GBT9704_2012排版工具.Contracts;

public sealed record AuditSummaryContract(
    string InputPath,
    string OutputPath,
    string? DocumentKind,
    string? DocumentKindReason,
    int AttachmentCount,
    bool HasPageNumberField,
    int LandscapeSectionCount,
    int ImprintCount,
    IReadOnlyList<string> CheckedRuleCodes,
    IReadOnlyList<string> CheckedRuleNames,
    IReadOnlyList<string> MatchedRuleCodes,
    IReadOnlyList<string> MatchedRuleNames,
    string GateResult,
    string? FailureStage,
    string? RuleCode,
    string? RuleName,
    bool NeedsManualReview,
    string? Message
);
