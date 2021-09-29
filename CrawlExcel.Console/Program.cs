using System.Collections.Generic;
using System.IO;
using System.Linq;
using CrawlExcel.Console.Models;
using Fizzler.Systems.HtmlAgilityPack;
using HtmlAgilityPack;
using OfficeOpenXml;

namespace CrawlExcel.Console
{
    class Program
    {
        static void Main()
        {
            var reports = new List<FinancialReport>();
            
            reports.AddRange(GetReports("DPM", new ReportYear() { Quarter = 2, Year = 2021 },
                new ReportYear() { Quarter = 2, Year = 2021 }));
            // reports.AddRange(GetReports("DMC", new ReportYear() { Quarter = 2, Year = 2010 },
            //     new ReportYear() { Quarter = 2, Year = 2021 }));
            
                SubmitReport("/Users/lucas/IdeaProjects/CrawlExcel/CrawlExcel.Console/Submit.xlsx",
                    reports.Where(report => report.Year.Quarter is 4 or 2).ToArray());
        }

        private static FinancialReport GetReport(string stockCode, ReportYear year)
        {
            var accounts = new List<Account>();
            accounts.AddRange(GetBalanceSheetAccounts(stockCode, year));
            accounts.AddRange(GetIncomeStatusAccounts(stockCode, year));
            accounts.Add(GetDepreciationAccount(stockCode, year));

            return new FinancialReport()
            {
                CompanyCode = stockCode,
                Year = year,
                Accounts = accounts
            };
        }

        private static IEnumerable<FinancialReport> GetReports(string stocksCode, ReportYear beginYear,
            ReportYear endYear)
        {
            return GetReportYearInRange(beginYear, endYear).Select(year => GetReport(stocksCode, year));
        }

        private static List<ReportYear> GetReportYearInRange(ReportYear beginYear, ReportYear endYear)
        {
            var years = new List<ReportYear>();
            while (beginYear.CompareTo(endYear) <= 0)
            {
                years.Add(new ReportYear() { Quarter = beginYear.Quarter, Year = beginYear.Year });
                if (beginYear.Quarter == 4)
                {
                    beginYear.Quarter = 1;
                    beginYear.Year++;
                }
                else
                {
                    beginYear.Quarter++;
                }
            }

            return years;
        }

        private static IEnumerable<Account> GetBalanceSheetAccounts(string stockCode, ReportYear year)
        {
            var dom = new HtmlWeb().Load(
                $"https://s.cafef.vn/bao-cao-tai-chinh/{stockCode}/BSheet/{year.Year}/{year.Quarter}/1/0/bao-cao-tai-chinh-tong-cong-ty-phan-bon-va-hoa-chat-dau-khictcp.chn");

            return dom.DocumentNode.QuerySelectorAll("#tableContent tr")
                .Where(node => !string.IsNullOrWhiteSpace(node.Id))
                .Select(node => new Account()
                {
                    Code = node.Id,
                    Value = node.QuerySelector("td:nth-child(5)").InnerText.Replace(",", string.Empty).Trim()
                }).ToList();
        }

        private static IEnumerable<Account> GetIncomeStatusAccounts(string stockCode, ReportYear year)
        {
            var dom = new HtmlWeb().Load(
                $"https://s.cafef.vn/bao-cao-tai-chinh/{stockCode}/IncSta/{year.Year}/{year.Quarter}/0/0/bao-cao-tai-chinh-tong-cong-ty-phan-bon-va-hoa-chat-dau-khictcp.chn");

            return dom.DocumentNode.QuerySelectorAll("#tableContent tr")
                .Where(node => !string.IsNullOrWhiteSpace(node.Id) && node.Id != "02").Select(node => new Account()
                {
                    Code = node.Id,
                    Value = node.QuerySelector("td:nth-child(5)").InnerText.Replace(",", string.Empty).Trim()
                }).ToList();
        }

        private static Account GetDepreciationAccount(string stockCode, ReportYear year)
        {
            var dom = new HtmlWeb().Load(
                $"https://s.cafef.vn/bao-cao-tai-chinh/{stockCode}/CashFlow/{year.Year}/{year.Quarter}/0/0/luu-chuyen-tien-te-gian-tiep-tong-cong-ty-phan-bon-va-hoa-chat-dau-khictcp.chn");

            return new Account()
            {
                Code = "02",
                Value = dom.DocumentNode.QuerySelector("#02 td:nth-child(5)").InnerText.Replace(",", string.Empty)
                    .Trim()
            };
        }

        private static void SubmitReport(string submitPath, params FinancialReport[] reports)
        {
            var submitFile = new FileInfo(submitPath);
            if (!submitFile.Exists)
            {
                System.Console.WriteLine("not exist file");
                return;
            }

            using var package = new ExcelPackage(submitFile);
            var worksheet = package.Workbook.Worksheets[0];
            var insertRow = 4;
            foreach (var report in reports)
            {
                for (var column = 1; column <= worksheet.Dimension.End.Column; column++)
                {
                    object? value = column switch
                    {
                        1 => report.CompanyCode,
                        2 => FormatYear(report.Year),
                        _ => report.Accounts.FirstOrDefault(account =>
                                account.Code == worksheet.Cells[2, column].Value.ToString())
                            ?.Value
                    };

                    worksheet.Cells[insertRow, column].Value = value;
                }

                insertRow++;
            }

            package.Save();
        }

        private static string FormatYear(ReportYear year)
        {
            var quarter = year.Quarter == 2 ? "06" : "12";
            return $"{quarter}{year.Year}";
        }
    }
}