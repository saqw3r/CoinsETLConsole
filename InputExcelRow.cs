using System;
using System.Reflection;
using System.Text;

namespace CoinsETLConsole
{
    internal class InputExcelRow
    {
        public InputExcelRow(string[] row)
        {
            this.Row = row;
            this.Project = row[0];
            this.ProjectPhase = row[1];
            this.Date = DateTime.Parse(row[2]); //Program.GetDatetimeFromCell(row[2]); //row[2].Cast<DateTime>();
            this.Worker = row[3];
            this.Task = row[4];
            this.Comment = row[5];
            this.ReportedHours = row[6];
            this.TimeType = row[7];
            this.Billable = row[8];
        }

        public string[] Row { get; set; }

        // Project
        public string Project { get; set; }

        // Project Phase
        public string ProjectPhase { get; set; }

        // Date
        public DateTime Date { get; set; }

        // Worker
        public string Worker { get; set; }

        // Project Task
        public string Task { get; set; }

        // Comment
        public string Comment { get; set; }

        // Reported Hours
        public string ReportedHours { get; set; }

        // Time Type
        public string TimeType { get; set; }

        // Billable
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
