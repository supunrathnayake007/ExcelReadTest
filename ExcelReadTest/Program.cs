using System;
using System.IO;
using OfficeOpenXml;

// Change the license context for EPPlus. This is required for non-commercial
// use of the library.
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var filePath = @"C:\Users\Supun\Downloads\V1HBL_PD-WCM_Term Loan_Jul 2024.xlsx";
var password = "EY@FaaS";
var sheetName = "2023-Jul";
var columnHeader = "2023-Jul DPD Bucket";

if (!File.Exists(filePath))
{
    Console.WriteLine($"File not found: {filePath}");
    return;
}

using var package = new ExcelPackage(new FileInfo(filePath), password);
var worksheet = package.Workbook.Worksheets[sheetName];
if (worksheet == null)
{
    Console.WriteLine($"Worksheet '{sheetName}' not found.");
    return;
}

// Determine the column index of the specified header.
int columnIndex = -1;
//int headerRow = worksheet.Dimension.Start.Row;
int headerRow = 5; 
for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
{
    if (worksheet.Cells[headerRow, col].Text.Trim() == columnHeader)
    {
        columnIndex = col;
        break;
    }
}

if (columnIndex == -1)
{
    Console.WriteLine($"Column '{columnHeader}' not found.");
    return;
}

// Iterate over rows and count how many contain the value "Current".
int count = 0;
for (int row = headerRow + 1; row <= worksheet.Dimension.End.Row; row++)
{
    var value = worksheet.Cells[row, columnIndex].Text.Trim();
    if (string.Equals(value, "Current", StringComparison.OrdinalIgnoreCase))
    {
        count++;
    }
}

Console.WriteLine($"Rows where '{columnHeader}' = 'Current': {count}");
