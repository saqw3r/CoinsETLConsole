using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace CoinsETLConsole
{
    class Program
    {
        public static DateTime GetDatetimeFromCell(object cellValue)
        {
            long dateNum = long.Parse(cellValue.ToString());
            var res = DateTime.FromOADate(dateNum);// Cast<DateTime>()
            return res;
        }

        public static void ExcelWrite(string filePath, List<ReportingItem> itemsToWrite, DateTime startDate, DateTime endDate)
        {
            // Creating an instance 
            // of ExcelPackage 
            ExcelPackage excel = new ExcelPackage();

            string[] projectNames = itemsToWrite.Select(x => x.Project).Distinct().ToArray();

            foreach (var projectName in projectNames)
            {
                // name of the sheet 
                var workSheet = excel.Workbook.Worksheets.Add(projectName);

                // setting the properties 
                // of the work sheet  
                //workSheet.TabColor = System.Drawing.Color;
                workSheet.DefaultRowHeight = 12;

                // Setting the properties 
                // of the first row 
                workSheet.Row(1).Height = 20;
                workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Row(1).Style.Font.Bold = true;
                workSheet.Row(1).Style.Font.Color.SetColor(System.Drawing.Color.Black);

                workSheet.Cells[1, 1].Value = "Period";
                workSheet.Cells[1, 2].Value = $"{startDate.ToShortDateString()} - {endDate.ToShortDateString()} ";

                workSheet.Row(2).Height = 20;
                workSheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Row(2).Style.Fill.PatternType = ExcelFillStyle.Solid;
                workSheet.Row(2).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                workSheet.Row(2).Style.Font.Color.SetColor(System.Drawing.Color.White);

                //workSheet.Row(4).Style.Font.Bold = ;

                // Header of the Excel sheet 
                workSheet.Cells[2, 1].Value = "Date";
                workSheet.Cells[2, 2].Value = "Reporter";
                workSheet.Cells[2, 3].Value = "Category";
                workSheet.Cells[2, 4].Value = "Task";
                workSheet.Cells[2, 5].Value = "Description";
                workSheet.Cells[2, 6].Value = "Hours";

                // Inserting the article data into excel 
                // sheet by using the for each loop 
                // As we have values to the first row  
                // we will start with second row 
                int recordIndex = 3;

                var itemsForCurrentProject = itemsToWrite.Where(x => x.Project == projectName).ToArray();

                foreach (var row in itemsForCurrentProject)
                {
                    workSheet.Cells[recordIndex, 1].Value = row.Date.ToShortDateString();
                    workSheet.Cells[recordIndex, 2].Value = row.Reporter;
                    workSheet.Cells[recordIndex, 3].Value = row.Category;
                    workSheet.Cells[recordIndex, 4].Value = row.Task;
                    workSheet.Cells[recordIndex, 5].Value = row.Description;
                    workSheet.Cells[recordIndex, 6].Value = row.Hours;
                    recordIndex++;
                }

                // By default, the column width is not  
                // set to auto fit for the content 
                // of the range, so we are using 
                // AutoFit() method here.  
                workSheet.Column(1).AutoFit();
                workSheet.Column(2).AutoFit();
                workSheet.Column(3).AutoFit();
                workSheet.Column(4).AutoFit();
                workSheet.Column(5).AutoFit();
                workSheet.Column(6).AutoFit();
            }

            // file name with .xlsx extension  

            if (File.Exists(filePath))
                File.Delete(filePath);

            // Create excel file on physical disk  
            FileStream objFileStrm = File.Create(filePath);
            objFileStrm.Close();

            // Write content to excel file  
            File.WriteAllBytes(filePath, excel.GetAsByteArray());
            Console.ReadKey();
        }

        static void Main(string[] args)
        {
            const string inputPath = @"C:\Users\ssurnin\Downloads\OneDrive_1_2-7-2020\Example - input.xlsx";
            //const string inputPath = @"D:\Sources\ETL_For_COINS\OneDrive_1_07.02.2020\Example - input.xlsx";
            const string outputPath = @"C:\Users\ssurnin\Downloads\OneDrive_1_2-7-2020\Output.xlsx";
            //const string outputPath = @"D:\Sources\ETL_For_COINS\OneDrive_1_07.02.2020\Output.xlsx";

            string[] dirs = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.xlsx");
            Console.WriteLine("The number of files starting with c is {0}.", dirs.Length);
            foreach (string dir in dirs)
            {
                Console.WriteLine(dir);
            }

            //read the Excel file as byte array
            byte[] bin = File.ReadAllBytes(inputPath);

            //create a new Excel package in a memorystream
            using (MemoryStream stream = new MemoryStream(bin))
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                //load worksheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];

                DateTime startDate = (DateTime)worksheet.Cells[4, 2].Value; //rows[3][1].Cast<DateTime>();
                DateTime endDate = (DateTime)worksheet.Cells[5, 2].Value; //rows[4][1].Cast<DateTime>();
                //DateTime startDate = Program.GetDatetimeFromCell(worksheet.Cells[4, 2].Value); //rows[3][1].Cast<DateTime>();
                //DateTime endDate = Program.GetDatetimeFromCell(worksheet.Cells[5, 2].Value); //rows[4][1].Cast<DateTime>();

                if (worksheet!=null)
                {
                    int rowsLength = worksheet.Dimension.End.Row - worksheet.Dimension.Start.Row; //rows.Count();

                    int startIndex = -1;
                    //for (int i = 0; i < rowsLength; i++)
                    //{
                    //    RowNoHeader row = rows[i];
                    //    if (row[0] == "Project")
                    //    {
                    //        startIndex = i + 1;
                    //        break;
                    //    }
                    //}

                    //loop all rows
                    for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                    {
                        if ((string)worksheet.Cells[i, 1].Value == "Project")
                        {
                            startIndex = i + 1;
                            break;
                        }
                        
                    }

                    List<InputExcelRow> readRows = new List<InputExcelRow>();
                    List<ReportingItem> reportedItems = new List<ReportingItem>();

                    for (int i = startIndex; i < rowsLength; i++)
                    {
                        string[] row = new string[9];
                       
                        //loop all columns in a row
                        for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                        {
                            //add the cell data to the List
                            if (worksheet.Cells[i, j].Value != null)
                            {
                                row[j-1] = worksheet.Cells[i, j].Value.ToString();
                            }
                        }
                        var input = new InputExcelRow(row);

                        var reportingItem = new ReportingItem(input);

                        var parsedTasks = ReportingItem.ParseComment(input.Comment, reportingItem);

                        readRows.Add(input);
                        reportedItems.AddRange(parsedTasks);
                    }

                    //print rows here
                    foreach (var item in reportedItems)
                    {
                        string itemToPrint = item.ToString();
                        Console.WriteLine(itemToPrint);

                        Console.WriteLine();
                        Console.WriteLine();
                    }

                    ExcelWrite(outputPath, reportedItems, startDate, endDate);
                }
            }

            Console.WriteLine("Execution is finished");
        }
    }
}
