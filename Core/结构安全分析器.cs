using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GBT9704_2012排版工具.Core;

public static class 结构安全分析器
{
    public static bool 段落可安全重写(Paragraph? paragraph)
    {
        if (paragraph == null)
            return false;

        return !存在高风险结构(paragraph);
    }

    public static bool 单元格可安全重写(TableCell? cell)
    {
        if (cell == null)
            return false;

        if (cell.Elements<Paragraph>().Skip(1).Any())
            return false;

        return !存在高风险结构(cell);
    }

    private static bool 存在高风险结构(OpenXmlElement element)
    {
        if (element.Descendants<Hyperlink>().Any())
            return true;
        if (element.Descendants<BookmarkStart>().Any() || element.Descendants<BookmarkEnd>().Any())
            return true;
        if (element.Descendants<CommentRangeStart>().Any() || element.Descendants<CommentRangeEnd>().Any())
            return true;
        if (element.Descendants<CommentReference>().Any())
            return true;
        if (element.Descendants<InsertedRun>().Any() || element.Descendants<DeletedRun>().Any())
            return true;
        if (element.Descendants<MoveFromRun>().Any() || element.Descendants<MoveToRun>().Any())
            return true;
        if (element.Descendants<FieldCode>().Any() || element.Descendants<SimpleField>().Any())
            return true;
        if (element.Descendants<Drawing>().Any())
            return true;
        if (element.Descendants<SdtRun>().Any() || element.Descendants<SdtBlock>().Any() || element.Descendants<SdtCell>().Any())
            return true;

        return false;
    }
}

