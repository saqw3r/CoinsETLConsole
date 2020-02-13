﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace CoinsETLConsole
{
    internal class ReportingItem
    {
        public static string ExtractShortProjectName(string longProjectName)
        {
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


        public static List<ReportingItem> ParseComment(string comment, ReportingItem baseReportingItem)
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
            comment = comment.Trim();
            int length = comment.Length;

            List<ReportingItem> result = new List<ReportingItem>();

            char[] digits = new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
            string[] stringSeparators = new string[] { "h", "hour", "hours", "\n\r" };

            int last_match_index = 0;

            List<string> tasks = new List<string>();

            for (int i = 0; i < length; i++)
            {
                foreach (var separator in stringSeparators)
                {
                    if (separator == "hour" && (i + separator.Length + 1) < comment.Length 
                        && comment[i + separator.Length + 1] == 's')
                    {
                        continue;
                    }

                    if (i + separator.Length <= length )
                    {
                        string s = comment.Substring(i, separator.Length);
                        string commentSubracted;
                        if (s.Equals(separator, StringComparison.InvariantCultureIgnoreCase))
                        {
                            if (i - 1 > last_match_index && digits.Contains(comment[i-1]))
                            {
                                commentSubracted = comment.Substring(last_match_index, i - last_match_index);
                                commentSubracted = commentSubracted.Trim().TrimStart('\n', '\r');
                                tasks.Add(commentSubracted);
                                last_match_index = i + separator.Length;
                            }
                            else if (i - 2 > last_match_index && digits.Contains(comment[i - 2]))
                            {
                                commentSubracted = comment.Substring(last_match_index, i - last_match_index - 1);
                                commentSubracted = commentSubracted.Trim().TrimStart('\n', '\r');
                                tasks.Add(commentSubracted);
                                last_match_index = i + separator.Length;
                            }
                        }
                    }
                }
            }

            if (!tasks.Any())
            {
                var reportingItemToAdd = new ReportingItem(baseReportingItem);
                reportingItemToAdd.Description = comment;
                reportingItemToAdd.Hours = null;
                if (reportingItemToAdd.Description != null)
                {
                    ExtractTaskOutOfDescription(reportingItemToAdd);
                    result.Add(reportingItemToAdd);
                }
            }

            foreach (var task in tasks)
            {
                var reportingItemToAdd = new ReportingItem(baseReportingItem);
                reportingItemToAdd.Hours = null;
                ExtractHoursAtTheEndingOfString(task, reportingItemToAdd);

                if (reportingItemToAdd.Description != null)
                {
                    ExtractTaskOutOfDescription(reportingItemToAdd);
                    result.Add(reportingItemToAdd);
                }
            }

            return result;
        }

        public static void ExtractTaskOutOfDescription(ReportingItem reportingItem)
        {
            string[] taskIds = new string[] { "Story", "Task", "Test Case", "US", "User Story", "Bug", "Defect" };

            string description = reportingItem.Description.Trim();

            bool isTaskCouldBeDefined = taskIds.Any(s => description.StartsWith(s));

            if (isTaskCouldBeDefined)
            {
                int index = description.IndexOf(':');
                if (index > 0)
                {
                    reportingItem.Task = description.Substring(0, index).Trim();
                    reportingItem.Description = description.Substring(index+1).Trim();
                }
            }
        }

        public static void ExtractHoursAtTheEndingOfString(string taskComment, ReportingItem toUpdateByParsing)
        {
            int length = taskComment.Length;
            int index;

            char[] allowedForDoubleValue = new char [] { '0', '1', '2', '3','4','5','6','7','8','9', '.', ','};
            for (index = length-1; index > -1; index--)
            {
                if (!allowedForDoubleValue.Contains(taskComment[index]))
                    break;
            }

            if (index < 0 || index + 1 >= length)
            {
                toUpdateByParsing.Hours = null;
                return;
            }
                
            string tmpToParse = taskComment.Substring(index+1);
            double parsedValue = double.Parse(tmpToParse, CultureInfo.InvariantCulture);
            toUpdateByParsing.Hours = parsedValue;
            toUpdateByParsing.Description = taskComment.Substring(0, index).TrimEnd(':', '-').TrimEnd();
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

            input.ReportedHours = input.ReportedHours?.TrimEnd('h', 'H');

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

        public double? Hours { get; set; }

        private PropertyInfo[] _PropertyInfos = null;

        public override string ToString()
        {
            if (_PropertyInfos == null)
                _PropertyInfos = this.GetType().GetProperties();

            var sb = new StringBuilder();

            foreach (var info in _PropertyInfos)
            {
                var value = info.GetValue(this, null) ?? "?";
                sb.AppendLine(info.Name + ": " + value.ToString());
            }

            return sb.ToString();
        }
    }
}