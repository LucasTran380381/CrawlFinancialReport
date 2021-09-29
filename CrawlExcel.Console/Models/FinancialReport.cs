using System.Collections.Generic;

namespace CrawlExcel.Console.Models
{
    public class FinancialReport
    {
        public string CompanyCode { get; set; }

        public ReportYear Year { get; set; }

        public List<Account> Accounts { get; set; }
    }
}