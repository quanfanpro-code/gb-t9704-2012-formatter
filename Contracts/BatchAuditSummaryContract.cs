namespace GBT9704_2012排版工具.Contracts;

public sealed record BatchAuditSummaryContract(
    string RootPath,
    int TotalCount,
    int SuccessCount,
    int FailCount,
    IReadOnlyList<BatchAuditItemContract> Items
);
