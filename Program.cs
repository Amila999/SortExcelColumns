using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office.Word;

class Program
{
    static void Main(string[] args)
    {
        string filePath = "Anu.xlsx";
        var workbook = new XLWorkbook(filePath);
        var worksheet = workbook.Worksheet(1);

        // Get the range of cells from the second column to the last column in seventh row
        var dataRange = worksheet.Range(
            worksheet.Cell(7, 2),
            worksheet.Cell(7, worksheet.LastColumnUsed().ColumnNumber() - 1)
        );
        Console.WriteLine(dataRange);

        List<List<object>> dataList = new List<List<object>>();
        ListOfRowContents(dataRange, dataList);

        dataList = dataList.OrderBy(innerList => (int)innerList[0]).ToList();

        dataList = dataList.OrderBy(innerList => innerList[0].ToString().Substring(0, 2))
                          .ThenBy(innerList => int.Parse(innerList[0].ToString().Substring(2)))
                          .ToList();

        foreach (var innerList in dataList)
        {
            //Console.WriteLine($"Integer: {innerList[0]}, IXLColumn: {innerList[1]}");
        }

        // Create a new workbook to save the sorted data
        var sortedWorkbook = new XLWorkbook();
        var sortedWorksheet = sortedWorkbook.Worksheets.Add("Sorted Data");
        int rowNum = 1;
        int columnNum = 1;
        foreach (var innerList in dataList)
        {
            sortedWorksheet.Cell(1, columnNum).Value = Convert.ToInt32(innerList[0].ToString());
            var columnData = (IXLColumn)innerList[1];
            //IXLCell vv =;
            var bb = sortedWorksheet.Cell(1, columnNum).WorksheetColumn();
            columnData.CopyTo(bb);
            Console.WriteLine(bb);
            columnNum++;
        }

        var NewdataRange = sortedWorksheet.Range(
            sortedWorksheet.Cell(7, 1),
            sortedWorksheet.Cell(7, worksheet.LastColumnUsed().ColumnNumber() - 1)
        );

        RenameSeventhRow(NewdataRange);

        //foreach (var innerList in dataList[0][1]) { }

        // Save the sorted data to a new Excel file
        sortedWorkbook.SaveAs("sorted_excel_file.xlsx");


    }
    private static void ListOfRowContents(IXLRange dataRange, List<List<object>> dataList)
    {
        foreach (var cell in dataRange.Cells())
        {

            List<object> rowData = [Convert.ToInt32(cell.Value.ToString())];
            var entireColumn = cell.WorksheetColumn();

            // Add entire column data to the inner list
            rowData.Add(entireColumn);

            // Add the inner list to the main list
            dataList.Add(rowData);
        }
    }

    static void RenameSeventhRow(IXLRange dataRange)
    {
        foreach (var cell in dataRange.CellsUsed())
        {
            var cellValue = cell.GetString();
            if (!string.IsNullOrEmpty(cellValue))
            {
                string newValue = "V" + cellValue[0] + "T" + cellValue[1] + "R" + cellValue.Substring(2);
                cell.Value = newValue;
            }
        }
    }
}
