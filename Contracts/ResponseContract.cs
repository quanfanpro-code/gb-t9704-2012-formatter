namespace GBT9704_2012排版工具.Contracts;

public sealed record ResponseContract(
    bool Success,
    string OutputPath,
    string? AuditPath,
    string? DocumentKind,
    string? DocumentKindReason,
    int AttachmentCount,
    bool HasPageNumberField,
    int LandscapeSectionCount,
    int ImprintCount,
    string? ErrorCode,
    string? Message,
    string? FailureStage,
    string? RuleCode,
    string? RuleName,
    bool NeedsManualReview
)
{
    public static ResponseContract Ok(
        string outputPath,
        string? auditPath,
        string? documentKind,
        string? documentKindReason,
        int attachmentCount = 0,
        bool hasPageNumberField = false,
        int landscapeSectionCount = 0,
        int imprintCount = 0)
    {
        return new ResponseContract(true, outputPath, auditPath, documentKind, documentKindReason, attachmentCount, hasPageNumberField, landscapeSectionCount, imprintCount, null, null, null, null, null, false);
    }

    public static ResponseContract Fail(
        string message,
        string? errorCode = null,
        string? failureStage = null,
        string? ruleCode = null,
        string? ruleName = null,
        bool needsManualReview = false,
        string? auditPath = null,
        string? documentKind = null,
        string? documentKindReason = null,
        int attachmentCount = 0,
        bool hasPageNumberField = false,
        int landscapeSectionCount = 0,
        int imprintCount = 0)
    {
        return new ResponseContract(false, string.Empty, auditPath, documentKind, documentKindReason, attachmentCount, hasPageNumberField, landscapeSectionCount, imprintCount, errorCode, message, failureStage, ruleCode, ruleName, needsManualReview);
    }
}
