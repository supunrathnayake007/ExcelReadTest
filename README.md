# ExcelReadTest

This sample console application demonstrates how to read data from an Excel
file using the EPPlus library.

## Running the sample

1. Install the .NET 8 SDK if you don't already have it.
2. From the repository root, run:

   ```bash
   dotnet run --project ExcelReadTest
   ```

The program reads the file located at
`C:\Users\SupunRathnayake\Downloads\V1HBL_PD-WCM_Term Loan_Jul 2024_ai.xlsx`
and prints the number of rows in the `2023-Jul` sheet where the
`2023-Jul DPD Bucket` column has the value `Current`.
