﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using LinqToExcel;
using LinqToExcel.Attributes;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace CoinsETLConsole
{
    internal class InputExcelRow
    {
        public InputExcelRow(RowNoHeader row)
        {
            this.Row = row;
            this.Project = row[0];
            this.ProjectPhase = row[1];
            this.Date = row[2].Cast<DateTime>();
            this.Worker = row[3];
            this.Task = row[4];
            this.Comment = row[5];
            this.ReportedHours = row[6];
            this.TimeType = row[7];
            this.Billable = row[8];
        }

        public RowNoHeader Row { get; set; }

        [ExcelColumn("A")] // Project
        public string Project { get; set; }

        [ExcelColumn("B")] // Project Phase
        public string ProjectPhase { get; set; }

        [ExcelColumn("C")] // Date
        public DateTime Date { get; set; }

        [ExcelColumn("D")] // Worker
        public string Worker { get; set; }

        [ExcelColumn("E")] // Project Task
        public string Task { get; set; }

        [ExcelColumn("F")] // Comment
        public string Comment { get; set; }

        [ExcelColumn("G")] // Reported Hours
        public string ReportedHours { get; set; }

        [ExcelColumn("H")] // Time Type
        public string TimeType { get; set; }

        [ExcelColumn("I")] // Billable
        public string Billable { get; set; }

        private PropertyInfo[] _PropertyInfos = null;

        public override string ToString()
        {
            if (_PropertyInfos == null)
                _PropertyInfos = this.GetType().GetProperties();

            var sb = new StringBuilder();

            foreach (var info in _PropertyInfos)
            {
                var value = info.GetValue(this, null) ?? "(null)";
                sb.AppendLine(info.Name + ": " + value.ToString());
            }

            return sb.ToString();
        }
    }

    internal class ReportingItem
    {
        public static string ExtractShortProjectName(string longProjectName)
        {
            int indexStart = 0;
            string endingPattern = "COINS CCCA - ";
            int indexEnd = longProjectName.IndexOf(endingPattern);
            
            if (indexEnd == -1)
            {
                endingPattern = "COINS - ";
                indexEnd = longProjectName.IndexOf(endingPattern);
                if (indexEnd == -1)
                {
                    endingPattern = "COINS ";
                    indexEnd = longProjectName.IndexOf(endingPattern);
                }
            }

            string result = longProjectName.Substring(indexEnd + endingPattern.Length);

            result = result.Replace(" project", "").Replace(" Project", "");

            return result;
        }


        public static List<ReportingItem> ParseComment(string comment, ReportingItem current)
        {
            //Story 123: lorem ipsum
            //Story - 123: lorem ipsum
            //Task 123: lorem ipsum
            //Task - 123: lorem ipsum
            //Test Case 123: lorem ipsum
            //Test Case-123: lorem ipsum
            //Test Case 22234: lorem ipsum
            //US 123: lorem ipsum
            //US - 123: lorem ipsum
            //User Story 22566: lorem ipsum
            //Bug 22565: lorem ipsum
            //Task 22777: lorem ipsum
            //Defect 456: lorem ipsum
            //Defect - 456: lorem ipsum
            //Bug 456: lorem ipsum
            //Bug - 456: lorem ipsum
            comment = comment.TrimStart(' ').TrimEnd(' ');
            int length = comment.Length;

            List<ReportingItem> result = new List<ReportingItem>();
            for (int i = 0; i < length; i++)
            {

            }

        }

        public ReportingItem(ReportingItem itemToCopy)
        {
            Project = itemToCopy.Project;
            Date = itemToCopy.Date;
            Reporter = itemToCopy.Reporter;
            Category = itemToCopy.Category;
        }

        public ReportingItem(InputExcelRow input)
        {
            Project = ExtractShortProjectName(input.Project);

            //long dateNum = long.Parse(input.Date);
            //Date = DateTime.FromOADate(dateNum);

            Date = input.Date;

            Reporter = Regex.Replace(input.Worker, @" \(.*?\)", "");

            Category = input.Task;

            input.ReportedHours = input.ReportedHours.TrimEnd('h', 'H');

            double tmp;
            double.TryParse(input.ReportedHours, NumberStyles.Any, CultureInfo.InvariantCulture, out tmp);

            this.Hours = tmp;
        }

        public string Project { get; set; }

        public DateTime Date { get; set; }

        public string Reporter { get; set; }

        public string Category { get; set; }

        public string Task { get; set; }

        public string Description { get; set; }

        public double Hours { get; set; }

        private PropertyInfo[] _PropertyInfos = null;

        public override string ToString()
        {
            if (_PropertyInfos == null)
                _PropertyInfos = this.GetType().GetProperties();

            var sb = new StringBuilder();

            foreach (var info in _PropertyInfos)
            {
                var value = info.GetValue(this, null) ?? "(null)";
                sb.AppendLine(info.Name + ": " + value.ToString());
            }

            return sb.ToString();
        }
    }

    class Program
    {
        public static void ExcelWrite(string filePath, List<ReportingItem> itemsToWrite)
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
                workSheet.TabColor = System.Drawing.Color.Black;
                workSheet.DefaultRowHeight = 12;

                // Setting the properties 
                // of the first row 
                workSheet.Row(1).Height = 20;
                workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Row(1).Style.Font.Bold = true;

                // Header of the Excel sheet 
                workSheet.Cells[1, 1].Value = "Date";
                workSheet.Cells[1, 2].Value = "Reporter";
                workSheet.Cells[1, 3].Value = "Category";
                workSheet.Cells[1, 4].Value = "Task";
                workSheet.Cells[1, 5].Value = "Description";
                workSheet.Cells[1, 6].Value = "Hours";

                // Inserting the article data into excel 
                // sheet by using the for each loop 
                // As we have values to the first row  
                // we will start with second row 
                int recordIndex = 2;

                foreach (var row in itemsToWrite)
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
            //const string inputPath = @"C:\Users\ssurnin\Downloads\OneDrive_1_2-7-2020\Example - input.xlsx";
            const string inputPath = @"D:\Sources\ETL_For_COINS\OneDrive_1_07.02.2020\Example - input.xlsx";
            //const string outputPath = @"C:\Users\ssurnin\Downloads\OneDrive_1_2-7-2020\Output.xlsx";
            const string outputPath = @"D:\Sources\ETL_For_COINS\OneDrive_1_07.02.2020\Output.xlsx";

            var excel = new ExcelQueryFactory(inputPath);

            var worksheet = excel.WorksheetNoHeader("Sheet1");

            var rows = worksheet.ToArray();

            int rowsLength = rows.Count();
            int startIndex = -1;

            for (int i = 0; i < rowsLength; i++)
            {
                RowNoHeader row = rows[i];
                if (row[0] == "Project")
                {
                    startIndex = i + 1;
                    break;
                }
            }

            List<InputExcelRow> readRows = new List<InputExcelRow>();
            List<ReportingItem> reportedItems = new List<ReportingItem>();

            for (int i = startIndex; i < rowsLength; i++)
            {
                //read data here
                var row = rows[i];
                var input = new InputExcelRow(row);
                var reportingItem = new ReportingItem(input);
               
                readRows.Add(input);
                reportedItems.Add(reportingItem);

                string itemToPrint = reportingItem.ToString();
                Console.WriteLine(itemToPrint);

                //for (int j = 0; j < row.Count; j++)
                //{
                //    Console.Write($"{row[j]} ");
                //}

                Console.WriteLine();
                Console.WriteLine();
            }

            ExcelWrite(outputPath, reportedItems);

            Console.WriteLine("Hello World!");
        }


    }
}