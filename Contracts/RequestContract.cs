using GBT9704_2012排版工具.Core;

namespace GBT9704_2012排版工具.Contracts;

public sealed class RequestContract
{
    public string InputPath { get; set; } = string.Empty;
    public string OutputPath { get; set; } = string.Empty;
    public GovDocumentKind? PreferredDocumentKind { get; set; }
}

