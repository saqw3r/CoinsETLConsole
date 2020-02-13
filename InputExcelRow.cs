using System;
using System.Reflection;
using System.Text;
using LinqToExcel;
using LinqToExcel.Attributes;
using OfficeOpenXml;

namespace CoinsETLConsole
{
    internal class InputExcelRow
    {
        public InputExcelRow(RowNoHeader row)
        {
            this.Row = row;
            this.Project = row[0];
            this.ProjectPhase = row[1];
            this.Date = Program.GetDatetimeFromCell(row[2]); //row[2].Cast<DateTime>();
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
                var value = info.GetValue(this, null) ?? "?";
                sb.AppendLine(info.Name + ": " + value.ToString());
            }

            return sb.ToString();
        }
    }
}
