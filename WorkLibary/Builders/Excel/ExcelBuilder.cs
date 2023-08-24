using OfficeOpenXml;

namespace WorkLibary.Builders.Excel;

public class ExcelBuilder
{
    private ExcelPackage package = new();
    private Dictionary<string, PageBuilder> pages = new();

    public PageBuilder AddPage(string worksheet)
    {
        if (pages.ContainsKey(worksheet))
            throw new Exception("Can't add excisted page");

        var pageWorksheet = package.Workbook.Worksheets.Add(worksheet);
        var pageBuilder = new PageBuilder(pageWorksheet);
        pages.Add(worksheet, pageBuilder);

        return pageBuilder;
    }

    public PageBuilder? GetPage(string worksheet) =>
        !pages.ContainsKey(worksheet) ? null : pages[worksheet];

    public void Build(string fileName)
    {
        package.SaveAs(fileName);
        package.Dispose();
    }
}