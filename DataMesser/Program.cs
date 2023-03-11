using OfficeOpenXml;
using System.Drawing;


internal class Program
{
    private static void Main(string[] args)
    {

        // set the license context for EPPlus
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // open the Excel workbook using EPPlus
        ExcelPackage excel = new ExcelPackage(new FileInfo(@"C:\Users\johnraesly\Downloads\Zone5.xlsx"));

        // get the worksheet by its index or name
        ExcelWorksheet sourceWorksheet = excel.Workbook.Worksheets["Cut Flowers"];

        // create a new worksheet
        ExcelWorksheet newWorksheet = excel.Workbook.Worksheets.Add($"New {sourceWorksheet.Name}");

        // initialize the row index in the new worksheet
        int newRow = 1;

        // iterate through each worksheet in the source Excel workbook
        //foreach (ExcelWorksheet sourceWorksheet in worksheet.Workbook.Worksheets)
        //{
            // get the value of the first cell in column A of the source worksheet
            string groupValue = sourceWorksheet.Cells[1, 1].Value?.ToString();

            // initialize the start row index for each group
            int startRow = 1;

            // iterate through each row in column A of the source worksheet
            for (int row = 1; row <= sourceWorksheet.Dimension.End.Row; row++)
            {
                // get the value of the cell in column A
                string value = sourceWorksheet.Cells[row, 1].Value?.ToString();

                // check if the cell value is not null and is different from the previous group value
                if (value != null && value != groupValue)
                {
                    // add the group header to the new worksheet
                    newWorksheet.Cells[newRow, 1].Value = groupValue;
                    newWorksheet.Cells[newRow, 1, newRow, 27].Style.Font.Bold = true;
                    newRow++;

                    // iterate through each row in the current group and copy columns D through AA to the new worksheet
                    for (int i = startRow; i < row; i++)
                    {
                        for (int col = 4; col < 27; col++)
                        {
                            newWorksheet.Cells[newRow, col - 3].Value = sourceWorksheet.Cells[i, col].Value;
                        }
                        newRow++;
                    }

                    // update the group value and start row index
                    groupValue = value;
                    startRow = row;
                }
            }

            // add the last group to the new worksheet
            newWorksheet.Cells[newRow, 1].Value = groupValue;
            newWorksheet.Cells[newRow, 1, newRow, 27].Style.Font.Bold = true;
            newRow++;

            for (int i = startRow; i <= sourceWorksheet.Dimension.End.Row; i++)
            {
                for (int col = 4; col <= 27; col++)
                {
                    newWorksheet.Cells[newRow, col - 3].Value = sourceWorksheet.Cells[i, col].Value;
                }
                newRow++;
            }
        //}
        // transpose the cells in the new worksheet
        //newWorksheet.Cells[1, 2, 27, newRow - 1].Copy(newWorksheet.Cells[2, 1, newRow - 1, 27]);

        // save the changes and close the workbook
        excel.Save();
        excel.Dispose();
    }
}