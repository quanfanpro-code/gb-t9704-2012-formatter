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
)
{
    public string 排版模式 { get; init; } = "普通材料";
    public bool Word实测已执行 { get; init; }
    public IReadOnlyList<公文检查项> 实际检查记录 { get; init; } = [];
}
