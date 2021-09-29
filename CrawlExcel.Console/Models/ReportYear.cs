using System;

namespace CrawlExcel.Console.Models
{
    public class ReportYear : IComparable<ReportYear>
    {
        public int Quarter { get; set; }

        public int Year { get; set; }

        public int CompareTo(ReportYear? year)
        {
            if (year == null) throw new ArgumentNullException(nameof(year));
            var yearCompare = Year.CompareTo(year.Year);
            return yearCompare != 0 ? yearCompare : Quarter.CompareTo(year.Quarter);
        }
    }
}