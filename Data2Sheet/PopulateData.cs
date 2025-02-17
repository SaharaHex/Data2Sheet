using MySql.Data.MySqlClient;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;

namespace Data2Sheet
{
    /// <summary>
    /// Populate Data for the application.
    /// Prepare the data in the format needed for Excel file creation.
    /// Note set database connection up in config
    /// </summary>
    public class PopulateData
    {
        private DBConnection dbCon;

        public PopulateData() 
        {
            var appSettings = ConfigurationManager.AppSettings;
            dbCon = DBConnection.Instance();
            dbCon.Server = appSettings["Server"];
            dbCon.DatabaseName = appSettings["DatabaseName"];
            dbCon.UserName = appSettings["UserName"];
            dbCon.Password = appSettings["Password"];
        }

        public void PopulateDataStock(DataTable dt, string clientID)
        {
            var appSettings = ConfigurationManager.AppSettings;

            if (appSettings["TestRandom"] == "yes")
            {
                Testing.StockGetData(dt);
            }
            else
            {
                dbCon.Connection = null;
                if (dbCon.IsConnect())
                {
                    string query = ReportQuery.StockQuery(clientID);
                    var cmd = new MySqlCommand(query, dbCon.Connection);
                    var reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        ReportQuery.StockGetData(dt, reader);
                    }
                    dbCon.Close();
                }
            }
        }

        public void PopulateDataEdinburgh(DataTable dt, string clientID)
        {
            var appSettings = ConfigurationManager.AppSettings;

            if (appSettings["TestRandom"] == "yes")
            {
                Testing.EdinburghGetData(dt);
            }
            else
            {
                dbCon.Connection = null;
                if (dbCon.IsConnect())
                {
                    string query = ReportQuery.EdinburghQuery(clientID);
                    var cmd = new MySqlCommand(query, dbCon.Connection);
                    var reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        ReportQuery.EdinburghGetData(dt, reader);
                    }
                    dbCon.Close();
                }
            }
        }

        static DataTable ToDataTable(List<Dictionary<string, string>> list)
        {
            DataTable result = new DataTable();
            if (list.Count == 0)
                return result;

            //Put data into a single row for each fields in Dictionary list(s)
            var columnNames = list.SelectMany(dict => dict.Keys).Distinct(); //heading 
            result.Columns.AddRange(columnNames.Select(c => new DataColumn(c)).ToArray());
            foreach (Dictionary<string, string> item in list)
            {
                var row = result.NewRow();
                foreach (var key in item.Keys)
                {
                    row[key] = item[key]; //value
                }

                result.Rows.Add(row);
            }

            return result;
        }

        public void PopulateDataMK(Dictionary<string, string> list, out DataTable dt)
        {
            var appSettings = ConfigurationManager.AppSettings;

            ReadData(ReportQuery.MKQuery_ChromebooksNew("uk", "UKChromebooksNew"), "UKChromebooksNew");
            ReadData(ReportQuery.MKQuery_ChromebooksUsed("uk", "UKChromebooksUsed"), "UKChromebooksUsed");
            ReadData(ReportQuery.MKQuery_DellNew("uk", "UKDellNew"), "UKDellNew");
            ReadData(ReportQuery.MKQuery_DellUsed("uk", "UKDellUsed"), "UKDellUsed");
            ReadData(ReportQuery.MKQuery_Dell65W("uk", "UKDell65W"), "UKDell65W");
            ReadData(ReportQuery.MKQuery_Dell130W("uk", "UKDell130W"), "UKDell130W");
            ReadData(ReportQuery.MKQuery_HPNew("uk", "UKHPNew"), "UKHPNew");
            ReadData(ReportQuery.MKQuery_HPUsed("uk", "UKHPUsed"), "UKHPUsed");
            ReadData(ReportQuery.MKQuery_HP45W("uk", "UKHP45W"), "UKHP45W");
            ReadData(ReportQuery.MKQuery_LenovoNew("uk", "UKLenovoNew"), "UKLenovoNew");
            ReadData(ReportQuery.MKQuery_LenovoUsed("uk", "UKLenovoUsed"), "UKLenovoUsed");
            ReadData(ReportQuery.MKQuery_Lenovo65W("uk", "UKLenovo65W"), "UKLenovo65W");
            ReadData(ReportQuery.MKQuery_MobileNew("uk", "UKMobileNew"), "UKMobileNew");
            ReadData(ReportQuery.MKQuery_MobileUsed("uk", "UKMobileUsed"), "UKMobileUsed");
            ReadData(ReportQuery.MKQuery_MobileCharger("uk", "UKMobileCharger"), "UKMobileCharger"); 
            ReadData(ReportQuery.MKQuery_UKSIMCards(), "UKSIMCards");
            ReadData(ReportQuery.MKQuery_UKDellCharger65(), "UKDellCharger65");
            ReadData(ReportQuery.MKQuery_UKDellCharger130(), "UKDellCharger130");
            ReadData(ReportQuery.MKQuery_UKHPCharger(), "UKHPCharger");
            ReadData(ReportQuery.MKQuery_UKLenovoCharger(), "UKLenovoCharger");
            ReadData(ReportQuery.MKQuery_UKMobileUSBCCharger(), "UKMobileUSBCCharger");

            ReadData(ReportQuery.MKQuery_ChromebooksNew("us", "USChromebooksNew"), "USChromebooksNew");
            ReadData(ReportQuery.MKQuery_ChromebooksUsed("us", "USChromebooksUsed"), "USChromebooksUsed");
            ReadData(ReportQuery.MKQuery_DellNew("us", "USDellNew"), "USDellNew");
            ReadData(ReportQuery.MKQuery_DellUsed("us", "USDellUsed"), "USDellUsed");
            ReadData(ReportQuery.MKQuery_Dell65W("us", "USDell65W"), "USDell65W");
            ReadData(ReportQuery.MKQuery_Dell130W("us", "USDell130W"), "USDell130W");
            ReadData(ReportQuery.MKQuery_HPNew("us", "USHPNew"), "USHPNew");
            ReadData(ReportQuery.MKQuery_HPUsed("us", "USHPUsed"), "USHPUsed");
            ReadData(ReportQuery.MKQuery_HP45W("us", "USHP45W"), "USHP45W");
            ReadData(ReportQuery.MKQuery_LenovoNew("us", "USLenovoNew"), "USLenovoNew");
            ReadData(ReportQuery.MKQuery_LenovoUsed("us", "USLenovoUsed"), "USLenovoUsed");
            ReadData(ReportQuery.MKQuery_Lenovo65W("us", "USLenovo65W"), "USLenovo65W");
            ReadData(ReportQuery.MKQuery_MobileNew("us", "USMobileNew"), "USMobileNew");
            ReadData(ReportQuery.MKQuery_MobileUsed("us", "USMobileUsed"), "USMobileUsed");
            ReadData(ReportQuery.MKQuery_MobileCharger("us", "USMobileCharger"), "USMobileCharger");

            ReadData(ReportQuery.MKQuery_ChromebooksNew("eu", "EUChromebooksNew"), "EUChromebooksNew");
            ReadData(ReportQuery.MKQuery_ChromebooksUsed("eu", "EUChromebooksUsed"), "EUChromebooksUsed");
            ReadData(ReportQuery.MKQuery_DellNew("eu", "EUDellNew"), "EUDellNew");
            ReadData(ReportQuery.MKQuery_DellUsed("eu", "EUDellUsed"), "EUDellUsed");
            ReadData(ReportQuery.MKQuery_Dell65W("eu", "EUDell65W"), "EUDell65W");
            ReadData(ReportQuery.MKQuery_Dell130W("eu", "EUDell130W"), "EUDell130W");
            ReadData(ReportQuery.MKQuery_HPNew("eu", "EUHPNew"), "EUHPNew");
            ReadData(ReportQuery.MKQuery_HPUsed("eu", "EUHPUsed"), "EUHPUsed");
            ReadData(ReportQuery.MKQuery_HP45W("eu", "EUHP45W"), "EUHP45W");
            ReadData(ReportQuery.MKQuery_LenovoNew("eu", "EULenovoNew"), "EULenovoNew");
            ReadData(ReportQuery.MKQuery_LenovoUsed("eu", "EULenovoUsed"), "EULenovoUsed");
            ReadData(ReportQuery.MKQuery_Lenovo65W("eu", "EULenovo65W"), "EULenovo65W");
            ReadData(ReportQuery.MKQuery_MobileNew("eu", "EUMobileNew"), "EUMobileNew");
            ReadData(ReportQuery.MKQuery_MobileUsed("eu", "EUMobileUsed"), "EUMobileUsed");
            ReadData(ReportQuery.MKQuery_MobileCharger("eu", "EUMobileCharger"), "EUMobileCharger");

            List<Dictionary<string, string>> it = new List<Dictionary<string, string>>();
            it.Add(list);
            dt = ToDataTable(it);
            dt.TableName = "MK stock report";
            
            void ReadData(string query, string rowName)
            {
                if (appSettings["TestRandom"] == "yes")
                {
                    Testing.AddDataToDictionary(list, rowName);
                }
                else
                {
                    dbCon.Connection = null;
                    if (dbCon.IsConnect())
                    {
                        var cmd = new MySqlCommand(query, dbCon.Connection);
                        var reader = cmd.ExecuteReader();

                        while (reader.Read())
                        {
                            ReportQuery.AddDataToDictionary(list, reader, rowName);
                        }
                        dbCon.Close();
                    }
                }
            }
        }

        public void PopulateDataBristol(Dictionary<string, string> list, out DataTable dt)
        {
            var appSettings = ConfigurationManager.AppSettings;

            ReadData(ReportQuery.BristolQuery_KeyboardNew(), "KeyboardNew");
            ReadData(ReportQuery.BristolQuery_KeyboardUsed(), "KeyboardUsed");
            ReadData(ReportQuery.BristolQuery_HeadsetsNew(), "HeadsetsNew");
            ReadData(ReportQuery.BristolQuery_HeadsetsUsed(), "HeadsetsUsed");
            ReadData(ReportQuery.BristolQuery_PhoneCase(), "PhoneCase");
            ReadData(ReportQuery.BristolQuery_Protector(), "Protector");
            ReadData(ReportQuery.BristolQuery_TypeCHubNew(), "TypeCHubNew");
            ReadData(ReportQuery.BristolQuery_TypeCHubUsed(), "TypeCHubUsed");
            ReadData(ReportQuery.BristolQuery_GeomaticsNew(), "GeomaticsNew");
            ReadData(ReportQuery.BristolQuery_GeomaticsUsed(), "GeomaticsUsed");
            ReadData(ReportQuery.BristolQuery_GeomaticsUSBC(), "GeomaticsUSBC");
            ReadData(ReportQuery.BristolQuery_FinanceNew(), "FinanceNew");
            ReadData(ReportQuery.BristolQuery_FinanceUsed(), "FinanceUsed");
            ReadData(ReportQuery.BristolQuery_FinanceUSBC(), "FinanceUSBC");
            ReadData(ReportQuery.BristolQuery_StandardNew(), "StandardNew");
            ReadData(ReportQuery.BristolQuery_StandardUsed(), "StandardUsed");
            ReadData(ReportQuery.BristolQuery_DellUSBC(), "DellUSBC");
            ReadData(ReportQuery.BristolQuery_LenovoUSBC(), "LenovoUSBC");
            ReadData(ReportQuery.BristolQuery_SamsungNew(), "SamsungNew");
            ReadData(ReportQuery.BristolQuery_SamsungUsed(), "SamsungUsed");
            ReadData(ReportQuery.BristolQuery_SamsungChargers(), "SamsungChargers");
            ReadData(ReportQuery.BristolQuery_AppleNew(), "AppleNew");
            ReadData(ReportQuery.BristolQuery_AppleUsed(), "AppleUsed");
            ReadData(ReportQuery.BristolQuery_AppleChargers(), "AppleChargers");
            ReadData(ReportQuery.BristolQuery_63WCharger(), "63WCharger");
            ReadData(ReportQuery.BristolQuery_DellAdapter(), "DellAdapter");
            ReadData(ReportQuery.BristolQuery_LenovoAdapter(), "LenovoAdapter");
            ReadData(ReportQuery.BristolQuery_LaptopBag(), "LaptopBag");
            ReadData(ReportQuery.BristolQuery_MobileSIM(), "MobileSIM");

            List<Dictionary<string, string>> it = new List<Dictionary<string, string>>();
            it.Add(list);
            dt = ToDataTable(it);
            dt.TableName = "Bristol Stock report";

            void ReadData(string query, string rowName)
            {
                if (appSettings["TestRandom"] == "yes")
                {
                    Testing.AddDataToDictionary(list, rowName);
                }
                else
                {
                    dbCon.Connection = null;
                    if (dbCon.IsConnect())
                    {
                        var cmd = new MySqlCommand(query, dbCon.Connection);
                        var reader = cmd.ExecuteReader();

                        while (reader.Read())
                        {
                            ReportQuery.AddDataToDictionary(list, reader, rowName);
                        }
                        dbCon.Close();
                    }
                }
            }
        }

        public void PopulateDataKent(Dictionary<string, string> list, out DataTable dt)
        {
            var appSettings = ConfigurationManager.AppSettings;

            ReadData(ReportQuery.KentQuery_KeyboardNew(), "KeyboardNew");
            ReadData(ReportQuery.KentQuery_KeyboardUsed(), "KeyboardUsed");
            ReadData(ReportQuery.KentQuery_HeadsetsNew(), "HeadsetsNew");
            ReadData(ReportQuery.KentQuery_HeadsetsUsed(), "HeadsetsUsed");
            ReadData(ReportQuery.KentQuery_BAUNew(), "BAUNew");
            ReadData(ReportQuery.KentQuery_BAUUsed(), "BAUUsed");
            ReadData(ReportQuery.KentQuery_BAUChargers(), "BAUChargers");
            ReadData(ReportQuery.KentQuery_StockLaptopsNew(), "StockLaptopsNew");
            ReadData(ReportQuery.KentQuery_StockLaptopsUsed(), "StockLaptopsUsed");
            ReadData(ReportQuery.KentQuery_StockLaptopsChargers(), "StockLaptopsChargers");
            ReadData(ReportQuery.KentQuery_MonitorNew(), "MonitorNew");
            ReadData(ReportQuery.KentQuery_MonitorUsed(), "MonitorUsed");
            ReadData(ReportQuery.KentQuery_MobilePhoneNew(), "MobilePhoneNew");
            ReadData(ReportQuery.KentQuery_MobilePhoneUsed(), "MobilePhoneUsed");
            ReadData(ReportQuery.KentQuery_MobilePhoneChargers(), "MobilePhoneChargers");
            ReadData(ReportQuery.KentQuery_MobileSim(), "MobileSim");

            List<Dictionary<string, string>> it = new List<Dictionary<string, string>>();
            it.Add(list);
            dt = ToDataTable(it);
            dt.TableName = "Kent Stock report";

            void ReadData(string query, string rowName)
            {
                if (appSettings["TestRandom"] == "yes")
                {
                    Testing.AddDataToDictionary(list, rowName);
                }
                else
                {
                    dbCon.Connection = null;
                    if (dbCon.IsConnect())
                    {
                        var cmd = new MySqlCommand(query, dbCon.Connection);
                        var reader = cmd.ExecuteReader();

                        while (reader.Read())
                        {
                            ReportQuery.AddDataToDictionary(list, reader, rowName);
                        }
                        dbCon.Close();
                    }
                }
            }
        }

        #region Notification
        public void PopulateEmail_MK_LowStockLaptop()
        {
            var appSettings = ConfigurationManager.AppSettings;

            ReadData(ReportQuery.MKQuery_LowStockLaptop());

            void ReadData(string query)
            {
                if (appSettings["TestRandom"] == "yes")
                {
                    Testing.MK_LowStockLaptop(5);
                }
                else
                {
                    dbCon.Connection = null;
                    if (dbCon.IsConnect())
                    {
                        var cmd = new MySqlCommand(query, dbCon.Connection);
                        var reader = cmd.ExecuteReader();

                        while (reader.Read())
                        {
                            Notification.MK_LowStockLaptop(reader, 10);
                        }
                        dbCon.Close();
                    }
                }
            }
        }
        #endregion
    }
}
