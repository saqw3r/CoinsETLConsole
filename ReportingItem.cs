using System;
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
        static char[] digits = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };

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

            string[] stringSeparators = { "\n\r", "\n" };

            int last_match_index = 0;

            List<string> tasks = new List<string>();

            for (int i = 0; i < length; i++)
            {
                foreach (var separator in stringSeparators)
                {
                    string commentSubracted = "";
                    if (i + separator.Length <= length)
                    {
                        string s = comment.Substring(i, separator.Length);
                        if (s.Equals(separator, StringComparison.InvariantCultureIgnoreCase))
                        {
                            commentSubracted = comment.Substring(last_match_index, i - last_match_index);
                            commentSubracted = commentSubracted.Trim().TrimStart('\n', '\r');
                            last_match_index = i + separator.Length;
                            if (commentSubracted.Length > 0)
                            {
                                tasks.Add(commentSubracted);
                            }
                        }
                    }
                    else 
                    {
                        commentSubracted = comment.Substring(last_match_index);
                        commentSubracted = commentSubracted.Trim().TrimStart('\n', '\r');
                        if (commentSubracted.Length > 0)
                        {
                            tasks.Add(commentSubracted);
                        }
                    }
                }
            }
            
            if (!tasks.Any())
            {
                var reportingItemToAdd = new ReportingItem(baseReportingItem);
                if (comment != null)
                {
                    ExtractActivitiesFromPartOfComment(comment, reportingItemToAdd);
                    if (reportingItemToAdd.Description != null)
                    {
                        ExtractTaskOutOfDescription(reportingItemToAdd);
                    }
                    result.Add(reportingItemToAdd);
                }
            }

            foreach (var task in tasks)
            {
                var reportingItemToAdd = new ReportingItem(baseReportingItem);
                reportingItemToAdd.Hours = null;
                if (task != null)
                {
                    ExtractActivitiesFromPartOfComment(task, reportingItemToAdd);
                    if (reportingItemToAdd.Description != null)
                    {
                        ExtractTaskOutOfDescription(reportingItemToAdd);
                    }
                    result.Add(reportingItemToAdd);
                }
            }

            return result;
        }

        private static void ExtractActivitiesFromPartOfComment(string commentToParse, ReportingItem itemToUpdate)
        {
            string[] timeMarkers = { "h", "hour", "hours" };

            commentToParse = commentToParse.TrimEnd(' ', '\r', '\n');

            foreach (string timeMarker in timeMarkers)
            {
                if (commentToParse.EndsWith(timeMarker))
                {
                    commentToParse = commentToParse.Remove(commentToParse.Length - timeMarker.Length).TrimEnd();
                    itemToUpdate.Description = commentToParse;

                    ExtractHoursAtTheEndingOfString(itemToUpdate.Description, itemToUpdate);
                    return;
                }
            }

            itemToUpdate.Description = commentToParse;
            return;
        }

        public static void ExtractTaskOutOfDescription(ReportingItem reportingItem)
        {
            if (reportingItem.Description!=null)
            {
                //string[] taskIds = new string[] { "Story", "Task", "Test Case", "US", "User Story", "Bug", "Defect" };

                string description = reportingItem.Description.Trim();

                //bool isTaskCouldBeDefined = taskIds.Any(s => description.StartsWith(s));

                //if (isTaskCouldBeDefined)
                //{

                char[] arrayOfSeparators = {' ', ':'};
                int index = description.IndexOf(':');
                if (index > 0)
                {
                    if (digits.Contains(description[index - 1]) ||
                        arrayOfSeparators.Contains(description[index-1]) && index > 1 && digits.Contains(description[index - 2]))
                    {
                        reportingItem.Task = description.Substring(0, index).Trim();
                        reportingItem.Description = description.Substring(index + 1).Trim();
                    }
                }
                //}
            }
        }

        public static void ExtractHoursAtTheEndingOfString(string taskComment, ReportingItem toUpdateByParsing)
        {
            int length = taskComment.Length;
            int index;

            string timeToParse = "";
            char[] digits = new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9'};

            char[] allowedForDoubleValue = new char [] { '0', '1', '2', '3','4','5','6','7','8','9', '.', ','};
            if (length > 0)
            {
                for (index = length - 1; index > -1; index--)
                {

                    if (!allowedForDoubleValue.Contains(taskComment[index]))
                        break;
                    else
                    {
                        timeToParse = taskComment[index] + timeToParse;
                    }
                }

                if (index < 0 || index + 1 >= length)
                {
                    toUpdateByParsing.Hours = null;
                    return;
                }

                bool containsAnyDigit = timeToParse.IndexOfAny(digits) > -1;
                if (containsAnyDigit)
                {
                    double parsedValue = double.Parse(timeToParse, CultureInfo.InvariantCulture);
                    toUpdateByParsing.Hours = parsedValue;
                    toUpdateByParsing.Description = taskComment.Substring(0, index).TrimEnd(':', '-').TrimEnd();
                }
                else
                {
                    toUpdateByParsing.Description = taskComment;
                    toUpdateByParsing.Hours = null;
                }
            }
            else
            {
                toUpdateByParsing.Description = null;
                toUpdateByParsing.Hours = null;
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

            input.ReportedHours = input.ReportedHours?.TrimEnd('h', 'H');

            double tmp;
            double.TryParse(input.ReportedHours, NumberStyles.Any, CultureInfo.InvariantCulture, out tmp);

            this.Hours = tmp;
        }

        public string Project { get; set; }

        public DateTime Date { get; set; }

        public string DateToExcel
        {
            get
            {
                if (Date == null)
                    return "?";

                return Date.ToShortDateString();
            }
        }

        public string Reporter { get; set; }

        public string ReporterToExcel
        {
            get
            {
                if (Reporter == null)
                    return "?";

                return Reporter;
            }
        }

        public string Category { get; set; }

        public string CategoryToExcel
        {
            get
            {
                if (Category == null)
                    return "?";

                return Category;
            }
        }

        public string Task { get; set; }

        public string TaskToExcel 
        {
            get {
                if (Task == null)
                    return "n/a";

                return Task;
            }
        }

        public string Description { get; set; }

        public string DescriptionToExcel
        {
            get
            {
                if (Description == null)
                    return "?";

                return Description;
            }
        }

        public double? Hours { get; set; }

        public string HoursToExcel
        {
            get
            {
                if (Hours == null)
                    return "?";

                return Hours.ToString();
            }
        }

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
