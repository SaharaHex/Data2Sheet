using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;

namespace Data2Sheet
{
    /// <summary>
    /// Test random data for the application.
    /// To run this set config flag TestRandom = yes
    /// e.g., can try out with No Database Connection
    /// </summary>
    internal class Testing
    {
        public static DataTable StockGetData(DataTable dt)
        {
            DataRow workRow = dt.NewRow();
            workRow["Asset Tag"] = "001247";
            workRow["Item Type"] = "Laptop";
            workRow["Item Code"] = "Lenovo 20M5-0013UK";
            workRow["Item Description"] = "ThinkPad  L380";
            workRow["Manufacturer"] = "Lenovo";
            workRow["Status"] = "Stock";
            workRow["Location"] = "J7";
            workRow["Condition"] = "B : Cosmetic Damage";
            workRow["Re Issue"] = "Yes";
            workRow["Has Charger"] = "No";
            workRow["CMAR"] = "No";
            workRow["PO Number"] = "";
            workRow["Purchase Date"] = "";
            workRow["Warranty Start"] = "20/12/2017 00:00:00";
            workRow["Warranty End"] = "";
            workRow["Last Audited"] = "";
            dt.Rows.Add(workRow);

            workRow = dt.NewRow();
            workRow["Asset Tag"] = "017801";
            workRow["Item Type"] = "Monitor";
            workRow["Item Code"] = "Philips V Line 24 Monitor 243V7";
            workRow["Item Description"] = "Philips V Line Full HD LCD monitor";
            workRow["Manufacturer"] = "Philips";
            workRow["Status"] = "Stock";
            workRow["Location"] = "Window Wall";
            workRow["Condition"] = "A+ : Brand New";
            workRow["Re Issue"] = "Yes";
            workRow["Has Charger"] = "N/A";
            workRow["CMAR"] = "";
            workRow["PO Number"] = "BC13663";
            workRow["Purchase Date"] = "20/09/2022 00:00:00";
            workRow["Warranty Start"] = "20/09/2022 00:00:00";
            workRow["Warranty End"] = "20/09/2023 00:00:00";
            workRow["Last Audited"] = "10/02/2023 00:00:00";
            dt.Rows.Add(workRow);
            Console.WriteLine("StockGetData added from Testing");
            return dt;
        }

        public static DataTable EdinburghGetData(DataTable dt)
        {
            Random rnd = new Random();
            int num = rnd.Next(71); // creates a number between 0 and 70

            string _consoleOutPut = num.ToString();
            DataRow workRow = dt.NewRow();
            workRow["DespatchDate"] = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
            workRow["OrderRef"] = (num.ToString() + 1);
            workRow["ClientID"] = (num.ToString() + 2);
            workRow["AssetTag"] = (num.ToString() + 3);
            workRow["ItemType"] = ("Laptop");
            workRow["UserID"] = (num.ToString() + 100);
            workRow["SerialNo"] = "LR0BPLS" + num.ToString() + 3;
            workRow["IMEI"] = "";
            dt.Rows.Add(workRow);

            workRow = dt.NewRow();
            workRow["DespatchDate"] = DateTime.Now.AddDays(-2).ToString("yyyy-MM-dd");
            workRow["OrderRef"] = (num.ToString() + 5);
            workRow["ClientID"] = (num.ToString() + 5);
            workRow["AssetTag"] = (num.ToString() + 5);
            workRow["ItemType"] = ("Mobile Phone");
            workRow["UserID"] = (num.ToString() + 105);
            workRow["SerialNo"] = "";
            workRow["IMEI"] = "3531291" + num.ToString() + 3;
            dt.Rows.Add(workRow);
            Console.WriteLine(_consoleOutPut);
            return dt;
        }

        public static Dictionary<string, string> AddDataToDictionary(Dictionary<string, string> dt, string rowName)
        {
            Random rnd = new Random();
            int num = rnd.Next(71); // creates a number between 0 and 70
            string _consoleOutPut = num.ToString();
            dt.Add(rowName, num.ToString());
            Console.WriteLine(_consoleOutPut);
            Thread.Sleep(500); //0.5 seconds
            return dt;
        }

        public static void MK_LowStockLaptop(int quantityTrigger)
        {
            string totalCount = "5";
            int.TryParse(totalCount, out int i);
            string message = "Low Stock Laptop: " + totalCount;

            if (i <= quantityTrigger)
            {
                Email em = new Email("test100@test.com", message, "Subject Line", "");
                em.SendMail();
            }
            Console.WriteLine(totalCount);
        }
    }
}
