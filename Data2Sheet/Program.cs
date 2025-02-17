using System;
using System.Collections.Generic;
using System.Data;

namespace Data2Sheet
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            Console.WriteLine("Data2Sheet");

            try
            {
                ExcelFile excelReport = new ExcelFile("Londan" + DateTime.Now.ToString("ddMMyyyhhmmss") + ".xlsx");
                DataTable dt = ReportQuery.StockDataTable("Londan stock report");
                PopulateData populateData = new PopulateData();
                populateData.PopulateDataStock(dt, "234");
                excelReport.CreateStockReport(dt, "Londan stock report", "Londan Office");

                excelReport = new ExcelFile("Manchester" + DateTime.Now.ToString("ddMMyyyhhmmss") + ".xlsx");
                dt = ReportQuery.StockDataTable("Manchester Stock report");
                populateData.PopulateDataStock(dt, "236");
                excelReport.CreateStockReport(dt, "Manchester Stock report", "Manchester Office");

                excelReport = new ExcelFile("Edinburgh" + DateTime.Now.ToString("ddMMyyyhhmmss") + ".xlsx");
                dt = ReportQuery.EdinburghDataTable("Edinburgh billing report");
                populateData.PopulateDataEdinburgh(dt, "190");
                excelReport.CreateEdinburghReport(dt, "Edinburgh billing report", "Edinburgh Office");

                Dictionary<string, string> list = new Dictionary<string, string>();

                excelReport = new ExcelFile("MK" + DateTime.Now.ToString("ddMMyyyhhmmss") + ".xlsx");
                populateData.PopulateDataMK(list, out dt);
                excelReport.CreateMKReport(dt);

                list = new Dictionary<string, string>();
                excelReport = new ExcelFile("Bristol" + DateTime.Now.ToString("ddMMyyyhhmmss") + ".xlsx");
                populateData.PopulateDataBristol(list, out dt);
                excelReport.CreateBristolReport(dt);

                list = new Dictionary<string, string>();
                excelReport = new ExcelFile("Kent" + DateTime.Now.ToString("ddMMyyyhhmmss") + ".xlsx");
                populateData.PopulateDataKent(list, out dt);
                excelReport.CreateKentReport(dt);

                populateData.PopulateEmail_MK_LowStockLaptop();
                Console.WriteLine("End");
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
            }
        }
    }
}