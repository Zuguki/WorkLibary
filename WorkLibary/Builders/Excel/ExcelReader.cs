using System.Text;
using OfficeOpenXml;

namespace WorkLibary.Builders.Excel;

public static class ExcelReader
{
    public static string ReadLine(ExcelWorksheet excelWorksheet, int row)
    {
        var stringBuilder = new StringBuilder();
        var lastColumn = GetLastColumn(excelWorksheet, row);

        for (var column = 1; column <= lastColumn; column++)
            stringBuilder.Append(excelWorksheet.Cells[row, column].Value + ":;:");

        return stringBuilder.ToString();
    }

    public static IEnumerable<string> ReadAllLines(ExcelWorksheet excelWorksheet, int rowStart = 2)
    {
        for (var row = rowStart; row < GetLastRow(excelWorksheet, 1); row++)
            yield return ReadLine(excelWorksheet, row);
    }

    public static IEnumerable<string> ReadCellsById(ExcelWorksheet excelWorksheet, int columnId, int rowStart = 1)
    {
        for (var row = rowStart; row < GetLastRow(excelWorksheet, 1); row++)
        {
            
            yield return (string) excelWorksheet.Cells[row, columnId].Value;
        }
    }

    public static int GetLastColumn(ExcelWorksheet excelWorksheet, int row)
    {
        var countOfEmptyCells = 0;
        
        for (var column = 1; column < excelWorksheet.Columns.EndColumn; column++)
        {
            if (excelWorksheet.Cells[row, column].Value is null)
            {
                if (countOfEmptyCells++ >= 5)
                    return column - 5;
            }

            if (countOfEmptyCells > 0)
                countOfEmptyCells = 0;
        }

        return excelWorksheet.Columns.EndColumn;
    }
    
    public static int GetLastRow(ExcelWorksheet excelWorksheet, int column)
    {
        var countOfEmptyCells = 0;
        
        for (var row = 1; row < excelWorksheet.Rows.EndRow; row++)
        {
            if (excelWorksheet.Cells[row, column].Value is null)
            {
                if (countOfEmptyCells++ >= 5)
                    return row - 5;
                
                continue;
            }

            if (countOfEmptyCells > 0)
                countOfEmptyCells = 0;
        }

        return excelWorksheet.Rows.EndRow;
    }
}