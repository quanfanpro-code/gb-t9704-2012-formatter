namespace GBT9704_2012排版工具.Contracts;

public sealed class RequestContract
{
    public string InputPath { get; set; } = string.Empty;
    public string OutputPath { get; set; } = string.Empty;
    public 排版模式 模式 { get; set; }
    public 公文要素? 要素 { get; set; }
    public double 套打红线距纸顶毫米 { get; set; } = 110;
    public List<公文检查项> 检查记录 { get; } = [];
    [System.Text.Json.Serialization.JsonIgnore]
    public CancellationToken CancellationToken { get; set; }
}

