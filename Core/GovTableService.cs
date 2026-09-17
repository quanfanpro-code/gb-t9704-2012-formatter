using DocumentFormat.OpenXml.Wordprocessing;

namespace GBT9704_2012排版工具.Core;

public sealed class GovTableService
{
    public void 格式化(Body body)
    {
        foreach (var table in body.Elements<Table>())
        {
            foreach (var row in table.Elements<TableRow>())
            {
                var cells = row.Elements<TableCell>().ToList();
                for (var column = 0; column < cells.Count; column++)
                {
                    var cell = cells[column];
                    cell.TableCellProperties ??= new TableCellProperties();
                    cell.TableCellProperties.TableCellVerticalAlignment =
                        new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center };

                    foreach (var paragraph in cell.Elements<Paragraph>())
                    {
                        GovOpenXmlHelper.设置段落行距(paragraph, 578);
                        GovOpenXmlHelper.设置段落缩进(paragraph);
                        GovOpenXmlHelper.确保段落属性(paragraph).Justification =
                            new Justification
                            {
                                Val = column == 0
                                    ? JustificationValues.Left
                                    : JustificationValues.Center
                            };

                        foreach (var run in GovOpenXmlHelper.获取可格式化文字运行(paragraph))
                            GovOpenXmlHelper.设置运行格式(run, "仿宋_GB2312", "Times New Roman", "28");
                    }
                }
            }
        }
    }
}
