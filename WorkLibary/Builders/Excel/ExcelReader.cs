using OfficeOpenXml;

namespace WorkLibary.Builders.Excel;

public class ExcelReader
{
    private readonly ExcelWorksheet excelWorksheet;

    public ExcelReader(ExcelWorksheet excelWorksheet)
    {
        this.excelWorksheet = excelWorksheet;
    }

    public string[] ReadLine(int row)
    {
        var result = new string[5];

        result[0] = (string) excelWorksheet.Cells[row, 3].Value + ":;:" + (string) excelWorksheet.Cells[row, 32].Value;

        result[1] = (string) excelWorksheet.Cells[row, 5].Value + ":;:" + (string) excelWorksheet.Cells[row, 6].Value +
                    ":;:" +
                    (string) excelWorksheet.Cells[row, 7].Value + ":;:" + (string) excelWorksheet.Cells[row, 8].Value +
                    ":;:" +
                    (string) excelWorksheet.Cells[row, 9].Value + ":;:" + (string) excelWorksheet.Cells[row, 10].Value;

        result[2] = (string) excelWorksheet.Cells[row, 12].Value + ":;:" + (string) excelWorksheet.Cells[row, 13].Value +
                    ":;:" +
                    (string) excelWorksheet.Cells[row, 14].Value + ":;:" + (string) excelWorksheet.Cells[row, 15].Value +
                    ":;:" +
                    (string) excelWorksheet.Cells[row, 16].Value + ":;:" + (string) excelWorksheet.Cells[row, 17].Value;

        result[3] = (string) excelWorksheet.Cells[row, 19].Value + ":;:" + (string) excelWorksheet.Cells[row, 20].Value +
                    ":;:" +
                    (string) excelWorksheet.Cells[row, 21].Value + ":;:" + (string) excelWorksheet.Cells[row, 22].Value +
                    ":;:" +
                    (string) excelWorksheet.Cells[row, 23].Value + ":;:" + (string) excelWorksheet.Cells[row, 24].Value;

        result[4] = (string) excelWorksheet.Cells[row, 26].Value + ":;:" + (string) excelWorksheet.Cells[row, 27].Value +
                    ":;:" +
                    (string) excelWorksheet.Cells[row, 28].Value + ":;:" + (string) excelWorksheet.Cells[row, 29].Value +
                    ":;:" +
                    (string) excelWorksheet.Cells[row, 30].Value + ":;:" + (string) excelWorksheet.Cells[row, 31].Value;

        return result;
    }

    public IEnumerable<string[]> ReadAllLines(int rowStart = 2, int rowTo = 100)
    {
        var row = rowStart;
            
        while (true)
        {
            if (excelWorksheet.Cells[row, 1].Value is null)
                break;
            
            yield return ReadLine(row++);
        }
    }

    public IEnumerable<string> ReadCellsById(int columnId, int rowStart = 1)
    {
        var row = rowStart;
        
        while (true)
        {
            if (excelWorksheet.Cells[row, columnId].Value is null)
                break;

            yield return (string) excelWorksheet.Cells[row++, columnId].Value;
        }
    }
}