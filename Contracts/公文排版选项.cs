namespace GBT9704_2012排版工具.Contracts;

public enum 排版模式 { 普通材料, 电子红头, 套打 }

public sealed class 公文要素
{
    public string 机关标志 { get; set; } = "";
    public string 标题 { get; set; } = "";
    public string 文号 { get; set; } = "";
    public string 主送机关 { get; set; } = "";
    public string 署名 { get; set; } = "";
    public string 成文日期 { get; set; } = "";
    public string 抄送 { get; set; } = "";
    public string 印发机关 { get; set; } = "";
    public string 印发日期 { get; set; } = "";
    public bool 预留盖章 { get; set; } = true;
    public bool 已确认 { get; set; }
}

public sealed record 公文检查项(string 编号, string 名称, string 结果, string 说明);
