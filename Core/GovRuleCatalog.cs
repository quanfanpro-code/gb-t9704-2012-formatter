namespace GBT9704_2012排版工具.Core;

public sealed record GovRuleDefinition(
    string Code,
    string Name,
    IReadOnlyList<GovDocumentKind> DocumentKinds
);

public static class GovRuleCatalog
{
    private static IReadOnlyList<GovDocumentKind> 所有文种 { get; } = [GovDocumentKind.普通公文];

    public static readonly GovRuleDefinition 正文存在性 = new(
        "BODY_MISSING",
        "正文存在性",
        所有文种);

    public static readonly GovRuleDefinition OpenXml结构合法性 = new(
        "OPENXML_SCHEMA",
        "OpenXml结构合法性",
        所有文种);

    public static readonly GovRuleDefinition 可见正文识别 = new(
        "BODY_VISIBLE_TEXT",
        "可见正文识别",
        所有文种);

    public static readonly GovRuleDefinition 页面尺寸 = new(
        "PAGE_SIZE",
        "页面尺寸",
        所有文种);

    public static readonly GovRuleDefinition 页面上边距 = new(
        "PAGE_MARGIN_TOP",
        "页面上边距",
        所有文种);

    public static readonly GovRuleDefinition 页面下边距 = new(
        "PAGE_MARGIN_BOTTOM",
        "页面下边距",
        所有文种);

    public static readonly GovRuleDefinition 页面左边距 = new(
        "PAGE_MARGIN_LEFT",
        "页面左边距",
        所有文种);

    public static readonly GovRuleDefinition 页面右边距 = new(
        "PAGE_MARGIN_RIGHT",
        "页面右边距",
        所有文种);

    public static readonly GovRuleDefinition 页眉距离 = new(
        "PAGE_HEADER_DISTANCE",
        "页眉距离",
        所有文种);

    public static readonly GovRuleDefinition 页脚距离 = new(
        "PAGE_FOOTER_DISTANCE",
        "页脚距离",
        所有文种);

    public static readonly GovRuleDefinition 版心行距栅格 = new(
        "PAGE_DOC_GRID",
        "版心行距栅格",
        所有文种);

    public static readonly GovRuleDefinition 附件标识 = new(
        "DOC_APPENDIX_MARK",
        "附件标识",
        所有文种);

    public static readonly GovRuleDefinition 页码字段 = new(
        "DOC_PAGE_NUMBER_FIELD",
        "页码字段",
        所有文种);

    public static readonly GovRuleDefinition 版记区域 = new(
        "DOC_IMPRINT_SECTION",
        "版记区域",
        所有文种);

    public static IReadOnlyList<GovRuleDefinition> All { get; } =
    [
        正文存在性,
        OpenXml结构合法性,
        可见正文识别,
        页面尺寸,
        页面上边距,
        页面下边距,
        页面左边距,
        页面右边距,
        页眉距离,
        页脚距离,
        版心行距栅格,
        附件标识,
        页码字段,
        版记区域
    ];

    public static IReadOnlyList<GovRuleDefinition> GetApplicableRules(GovDocumentKind? kind)
    {
        return All;
    }
}
