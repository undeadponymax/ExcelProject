using System;
using System.IO;
using GemBox.Spreadsheet;

class Program
{
   static void Main()
    {
        // Set license key to use GemBox.Spreadsheet in Free mode.
        SpreadsheetInfo.SetLicense("SN-2020Apr27-MV0CW7jiIJlsiZzwfdRaklP7MfRc6QM05dKX3Hzjam7BONU/IFJNXBhExD2+nu6wQaJLlen0h1/29EuBwfDriVZY4yQ==A");

        // Load Excel workbook from file's path.
        ExcelFile book = ExcelFile.Load("D:/Apps and files/Для Макса/Для Макса.xlsx");

        for (int sheetIndex = 0; sheetIndex < book.Worksheets.Count; sheetIndex++)
        {
            // Get Excel worksheet using zero-based index.
            ExcelWorksheet worksheet = book.Worksheets[sheetIndex];
            for (int rowIndex = 1; rowIndex < worksheet.Rows.Count; rowIndex++)
            {
                // Get Excel row using zero-based index.
                ExcelRow row = worksheet.Rows[rowIndex];
                string start = $"{row.Cells[0].Value}";
                ExcelFile workbook = new ExcelFile();
                if (File.Exists("D:/Apps and files/Для Макса/Сравнительная группа/" + start + ".xlsx"))
                {
                    workbook = ExcelFile.Load("D:/Apps and files/Для Макса/Сравнительная группа/" + start + ".xlsx");
                }
                else
                {
                    workbook.Worksheets.Add("Общие данные");
                }
                int length = workbook.Worksheets["Общие данные"].Rows.Count;
                int len1 = -1;
                bool doo = true;
                while(doo)
                {
                    len1 += 1;
                    if(workbook.Worksheets["Общие данные"].Rows[len1].Cells[0].Value == null)
                    {
                        doo = false;
                    }
                }
                for (int columnIndex = 0; columnIndex < 26; columnIndex++)
                {
                    ExcelCell cell = row.Cells[columnIndex];
                    workbook.Worksheets["Общие данные"].Rows[len1].Cells[columnIndex].Value = cell.Value;
                }
                workbook.Save("D:/Apps and files/Для Макса/Сравнительная группа/" + start + ".xlsx");
            }
        }
    }
}