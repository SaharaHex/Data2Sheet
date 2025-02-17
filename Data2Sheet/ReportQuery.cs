using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Data2Sheet
{
    /// <summary>
    /// All Queries for the application.
    /// Improvements: write Queries in stored procedure, as only had read access this was not done. The database was outside our network.
    /// </summary>

    public class ReportQuery
    {
        public static Dictionary<string, string> AddDataToDictionary(Dictionary<string, string> dt, MySqlDataReader reader, string rowName)
        {
            string _consoleOutPut = reader.GetString(0);
            dt.Add(rowName, reader.GetString(0));
            Console.WriteLine(_consoleOutPut);
            return dt;
        }

        #region Stock for Londan & Manchester & Edinburgh
        public static string StockQuery(string clientID)
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT AssetTag, ItemType, ItemCode, ItemDescription, Manufacturer, AssetStatus, AssetLocation, AssetCondition, ReIssue, HasCharger, CMARStock, PONumber, PurchaseDate, WarrantyStart, WarrantyEnd, LastAudited ");
            query.Append(" FROM client_asset_register ");
            query.Append(string.Format(" WHERE ClientID = '{0}' AND AssetStatus = 'Stock';", clientID));
            return query.ToString();
        }

        public static string EdinburghQuery(string clientID)
        {
            DateTime lastMonth = DateTime.Today.AddMonths(-1); //run for last month

            var firstDayOfMonth = new DateTime(lastMonth.Year, lastMonth.Month, 1);
            var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);

            StringBuilder query = new StringBuilder();
            query.Append(" SELECT od.DespatchDate, od.OrderRef, od.ClientID, oi.AssetTag, ar.ItemType, ar.UserID, ar.SerialNo");
            query.Append(" FROM tbl_order_details AS od " +
                "JOIN tbl_order_items AS oi ON oi.OrderID = od.OrderID " +
                "JOIN client_asset_register AS ar ON oi.AssetTag = ar.AssetTag ");
            query.Append(string.Format(" WHERE od.ClientID = '{0}' AND od.DespatchDate >= '{1}' AND od.DespatchDate <= '{2}' AND (ar.ItemType = 'Mobile Phone' OR ar.ItemType = 'Laptop') ;", clientID, firstDayOfMonth.ToString("yyyy-MM-dd"), lastDayOfMonth.ToString("yyyy-MM-dd")));
            return query.ToString();
        }

        public static DataTable StockDataTable(string tableName)
        {
            DataTable dt = new DataTable { TableName = tableName };
            dt.Columns.Add("Asset Tag", typeof(string));
            dt.Columns.Add("Item Type", typeof(string));
            dt.Columns.Add("Item Code", typeof(string));
            dt.Columns.Add("Item Description", typeof(string));
            dt.Columns.Add("Manufacturer", typeof(string));
            dt.Columns.Add("Status", typeof(string));
            dt.Columns.Add("Location", typeof(string));
            dt.Columns.Add("Condition", typeof(string));
            dt.Columns.Add("Re Issue", typeof(string));
            dt.Columns.Add("Has Charger", typeof(string));
            dt.Columns.Add("CMAR", typeof(string));
            dt.Columns.Add("PO Number", typeof(string));
            dt.Columns.Add("Purchase Date", typeof(string));
            dt.Columns.Add("Warranty Start", typeof(string));
            dt.Columns.Add("Warranty End", typeof(string));
            dt.Columns.Add("Last Audited", typeof(string));
            return dt;
        }

        public static DataTable EdinburghDataTable(string tableName)
        {
            DataTable dt = new DataTable { TableName = tableName };
            dt.TableName = tableName;
            dt.Columns.Add("DespatchDate", typeof(string));
            dt.Columns.Add("OrderRef", typeof(string));
            dt.Columns.Add("ClientID", typeof(string));
            dt.Columns.Add("AssetTag", typeof(string));
            dt.Columns.Add("ItemType", typeof(string));
            dt.Columns.Add("UserID", typeof(string));
            dt.Columns.Add("SerialNo", typeof(string));
            dt.Columns.Add("IMEI", typeof(string));
            return dt;
        }

        public static DataTable StockGetData(DataTable dt, MySqlDataReader reader)
        {
            string _consoleOutPut = reader.GetString(0);
            DataRow workRow = dt.NewRow();
            workRow["Asset Tag"] = (reader.IsDBNull(0) ? "" : reader.GetString(0));
            workRow["Item Type"] = (reader.IsDBNull(1) ? "" : reader.GetString(1));
            workRow["Item Code"] = (reader.IsDBNull(2) ? "" : reader.GetString(2));
            workRow["Item Description"] = (reader.IsDBNull(3) ? "" : reader.GetString(3));
            workRow["Manufacturer"] = (reader.IsDBNull(4) ? "" : reader.GetString(4));
            workRow["Status"] = (reader.IsDBNull(5) ? "" : reader.GetString(5));
            workRow["Location"] = (reader.IsDBNull(6) ? "" : reader.GetString(6));
            workRow["Condition"] = (reader.IsDBNull(7) ? "" : reader.GetString(7));
            workRow["Re Issue"] = (reader.IsDBNull(8) ? "" : reader.GetString(8));
            workRow["Has Charger"] = (reader.IsDBNull(9) ? "" : reader.GetString(9));
            workRow["CMAR"] = (reader.IsDBNull(10) ? "" : reader.GetString(10));
            workRow["PO Number"] = (reader.IsDBNull(11) ? "" : reader.GetString(11));
            workRow["Purchase Date"] = (reader.IsDBNull(12) ? "" : reader.GetString(12));
            workRow["Warranty Start"] = (reader.IsDBNull(13) ? "" : reader.GetString(13));
            workRow["Warranty End"] = (reader.IsDBNull(14) ? "" : reader.GetString(14));
            workRow["Last Audited"] = (reader.IsDBNull(15) ? "" : reader.GetString(15));
            dt.Rows.Add(workRow);
            Console.WriteLine(_consoleOutPut);
            return dt;
        }

        public static DataTable EdinburghGetData(DataTable dt, MySqlDataReader reader)
        {
            string _consoleOutPut = reader.GetString(0);
            DataRow workRow = dt.NewRow();
            workRow["DespatchDate"] = (reader.IsDBNull(0) ? "" : reader.GetString(0));
            workRow["OrderRef"] = (reader.IsDBNull(1) ? "" : reader.GetString(1));
            workRow["ClientID"] = (reader.IsDBNull(2) ? "" : reader.GetString(2));
            workRow["AssetTag"] = (reader.IsDBNull(3) ? "" : reader.GetString(3));
            workRow["ItemType"] = (reader.IsDBNull(4) ? "" : reader.GetString(4));
            workRow["UserID"] = (reader.IsDBNull(5) ? "" : reader.GetString(5));
            if (!reader.IsDBNull(6))
            {
                if (reader.GetString(6).Length >= 15)
                {
                    workRow["SerialNo"] = "";
                    workRow["IMEI"] = (reader.IsDBNull(6) ? "" : reader.GetString(6));
                }
                else
                {
                    workRow["SerialNo"] = (reader.IsDBNull(6) ? "" : reader.GetString(6));
                    workRow["IMEI"] = "";
                }
            }

            dt.Rows.Add(workRow);
            Console.WriteLine(_consoleOutPut);
            return dt;
        }
        #endregion

        #region MK
        private static string AssetLocationString(string location)
        {
            string output;
            switch (location)
            {
                case "uk":
                    output = "AND (AssetLocation != 'A1 - Boston Secure storage' AND AssetLocation != 'A1 - Dublin Secure storage')";
                    break;

                case "us":
                    output = "AND(AssetLocation = 'A1 - Boston Secure storage')";
                    break;

                case "eu":
                    output = "AND(AssetLocation = 'A1 - Dublin Secure storage')";
                    break;

                default:
                    output = "";
                    break;
            }
            return output;
        }

        public static string MKQuery_ChromebooksNew(string assetLocation, string fieldName)
        {
            StringBuilder query = new StringBuilder();
            query.Append(string.Format(" SELECT COUNT(AssetTag) AS '{0}' ", fieldName));
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' AND AssetCondition = 'A+ : Brand New' " +
                "AND (ItemCode LIKE '%chrome%' OR ItemDescription LIKE '%chrome%') " + AssetLocationString(assetLocation) +
                ";");
            return query.ToString();
        }

        public static string MKQuery_ChromebooksUsed(string assetLocation, string fieldName)
        {
            StringBuilder query = new StringBuilder();
            query.Append(string.Format(" SELECT COUNT(AssetTag) AS '{0}' ", fieldName));
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND (ItemCode LIKE '%chrome%' OR ItemDescription LIKE '%chrome%') " + AssetLocationString(assetLocation) +
                ";");
            return query.ToString();
        }

        public static string MKQuery_DellNew(string assetLocation, string fieldName)
        {
            StringBuilder query = new StringBuilder();
            query.Append(string.Format(" SELECT COUNT(AssetTag) AS '{0}' ", fieldName));
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' AND AssetCondition = 'A+ : Brand New' " +
                "AND Manufacturer = 'Dell' AND ItemType = 'Laptop' " + AssetLocationString(assetLocation) +
                ";");
            return query.ToString();
        }

        public static string MKQuery_DellUsed(string assetLocation, string fieldName)
        {
            StringBuilder query = new StringBuilder();
            query.Append(string.Format(" SELECT COUNT(AssetTag) AS '{0}' ", fieldName));
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND Manufacturer = 'Dell' AND ItemType = 'Laptop' " + AssetLocationString(assetLocation) +
                ";");
            return query.ToString();
        }

        public static string MKQuery_Dell65W(string assetLocation, string fieldName)
        {
            StringBuilder query = new StringBuilder();
            query.Append(string.Format(" SELECT COUNT(AssetTag) AS '{0}' ", fieldName));
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND HasCharger = 'No' " +
                "AND ItemType = 'Laptop' " +
                "AND (ItemCode LIKE '%DELL 3520%') " + AssetLocationString(assetLocation) +
                ";");
            return query.ToString();
        }

        public static string MKQuery_Dell130W(string assetLocation, string fieldName)
        {
            StringBuilder query = new StringBuilder();
            query.Append(string.Format(" SELECT COUNT(AssetTag) AS '{0}' ", fieldName));
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND HasCharger = 'No' " +
                "AND ItemType = 'Laptop' " +
                "AND (ItemCode LIKE '%Dell P104F001%') " + AssetLocationString(assetLocation) +
                ";");
            return query.ToString();
        }

        public static string MKQuery_HPNew(string assetLocation, string fieldName)
        {
            StringBuilder query = new StringBuilder();
            query.Append(string.Format(" SELECT COUNT(AssetTag) AS '{0}' ", fieldName));
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' AND AssetCondition = 'A+ : Brand New' " +
                "AND Manufacturer = 'HP' AND ItemType = 'Laptop' " + AssetLocationString(assetLocation) +
                ";");
            return query.ToString();
        }

        public static string MKQuery_HPUsed(string assetLocation, string fieldName)
        {
            StringBuilder query = new StringBuilder();
            query.Append(string.Format(" SELECT COUNT(AssetTag) AS '{0}' ", fieldName));
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND Manufacturer = 'HP' AND ItemType = 'Laptop' " + AssetLocationString(assetLocation) +
                ";");
            return query.ToString();
        }

        public static string MKQuery_HP45W(string assetLocation, string fieldName)
        {
            StringBuilder query = new StringBuilder();
            query.Append(string.Format(" SELECT COUNT(AssetTag) AS '{0}' ", fieldName));
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND HasCharger = 'No'" +
                "AND Manufacturer = 'HP' AND ItemType = 'Laptop' " + AssetLocationString(assetLocation) +
                ";");
            return query.ToString();
        }

        public static string MKQuery_LenovoNew(string assetLocation, string fieldName)
        {
            StringBuilder query = new StringBuilder();
            query.Append(string.Format(" SELECT COUNT(AssetTag) AS '{0}' ", fieldName));
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' AND AssetCondition = 'A+ : Brand New' " +
                "AND Manufacturer = 'Lenovo' AND ItemType = 'Laptop' " + AssetLocationString(assetLocation) +
                ";");
            return query.ToString();
        }

        public static string MKQuery_LenovoUsed(string assetLocation, string fieldName)
        {
            StringBuilder query = new StringBuilder();
            query.Append(string.Format(" SELECT COUNT(AssetTag) AS '{0}' ", fieldName));
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND Manufacturer = 'Lenovo' AND ItemType = 'Laptop' " + AssetLocationString(assetLocation) +
                ";");
            return query.ToString();
        }

        public static string MKQuery_Lenovo65W(string assetLocation, string fieldName)
        {
            StringBuilder query = new StringBuilder();
            query.Append(string.Format(" SELECT COUNT(AssetTag) AS '{0}' ", fieldName));
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND HasCharger = 'No'" +
                "AND Manufacturer = 'Lenovo' AND ItemType = 'Laptop' " + AssetLocationString(assetLocation) +
                ";");
            return query.ToString();
        }

        public static string MKQuery_MobileNew(string assetLocation, string fieldName)
        {
            StringBuilder query = new StringBuilder();
            query.Append(string.Format(" SELECT COUNT(AssetTag) AS '{0}' ", fieldName));
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' AND AssetCondition = 'A+ : Brand New' " +
                "AND ItemType = 'Mobile Phone' " + AssetLocationString(assetLocation) +
                ";");
            return query.ToString();
        }

        public static string MKQuery_MobileUsed(string assetLocation, string fieldName)
        {
            StringBuilder query = new StringBuilder();
            query.Append(string.Format(" SELECT COUNT(AssetTag) AS '{0}' ", fieldName));
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND ItemType = 'Mobile Phone' " + AssetLocationString(assetLocation) +
                ";");
            return query.ToString();
        }

        public static string MKQuery_MobileCharger(string assetLocation, string fieldName)
        {
            StringBuilder query = new StringBuilder();
            query.Append(string.Format(" SELECT COUNT(AssetTag) AS '{0}' ", fieldName));
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND HasCharger = 'No'" +
                "AND ItemType = 'Mobile Phone' " + AssetLocationString(assetLocation) +
                ";");
            return query.ToString();
        }

        #region MK UK
        public static string MKQuery_UKSIMCards()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS UKSIMCards ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND ItemType = 'Mobile SIM' " +
                "AND (AssetLocation != 'A1 - Boston Secure storage' AND AssetLocation != 'A1 - Dublin Secure storage');");
            return query.ToString();
        }

        public static string MKQuery_UKDellCharger65()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS UKDellCharger65 ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND Manufacturer = 'Dell' AND ItemType = 'Charger' " +                
                "AND (ItemCode LIKE '%65w%' OR ItemDescription LIKE '%65w%') " +
                "AND (AssetLocation != 'A1 - Boston Secure storage' AND AssetLocation != 'A1 - Dublin Secure storage');");
            return query.ToString();
        }

        public static string MKQuery_UKDellCharger130()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS UKDellCharger130 ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND Manufacturer = 'Dell' AND ItemType = 'Charger' " +
                "AND (ItemCode LIKE '%130w%' OR ItemDescription LIKE '%130w%') " +
                "AND (AssetLocation != 'A1 - Boston Secure storage' AND AssetLocation != 'A1 - Dublin Secure storage');");
            return query.ToString();
        }

        public static string MKQuery_UKHPCharger()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS UKHPCharger ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND Manufacturer = 'HP' AND ItemType = 'Charger' " +
                "AND (AssetLocation != 'A1 - Boston Secure storage' AND AssetLocation != 'A1 - Dublin Secure storage');");
            return query.ToString();
        }

        public static string MKQuery_UKLenovoCharger()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS UKHPCharger ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND Manufacturer = 'Lenovo' AND ItemType = 'Charger' " +
                "AND (AssetLocation != 'A1 - Boston Secure storage' AND AssetLocation != 'A1 - Dublin Secure storage');");
            return query.ToString();
        }

        public static string MKQuery_UKMobileUSBCCharger()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS UKMobileUSBCCharger ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 274 AND AssetStatus = 'Stock' " +
                "AND ItemType = 'Charger' " +
                "AND (ItemDescription LIKE '%Mobile%') " +
                "AND (AssetLocation != 'A1 - Boston Secure storage' AND AssetLocation != 'A1 - Dublin Secure storage');");
            return query.ToString();
        }
        #endregion

        #endregion

        #region Bristol
        public static string BristolQuery_KeyboardNew()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS KeyboardNew ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND ItemType = 'Keyboard/Mouse' " +
                "AND AssetCondition = 'A+ : Brand New';");
            return query.ToString();
        }

        public static string BristolQuery_KeyboardUsed()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS KeyboardUsed ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND ItemType = 'Keyboard/Mouse' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage');");
            return query.ToString();
        }

        public static string BristolQuery_HeadsetsNew()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS HeadsetsNew ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND ItemType = 'Headsets/earphones' " +
                "AND AssetCondition = 'A+ : Brand New';");
            return query.ToString();
        }

        public static string BristolQuery_HeadsetsUsed()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS HeadsetsUsed ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND ItemType = 'Headsets/earphones' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage');");
            return query.ToString();
        }

        public static string BristolQuery_PhoneCase()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS PhoneCase ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND (ItemCode LIKE '%Phone Case%');");
            return query.ToString();
        }

        public static string BristolQuery_Protector()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS Protector ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND ItemType = 'Screen Protector' AND (ItemDescription LIKE '%Samsung%');");
            return query.ToString();
        }

        public static string BristolQuery_TypeCHubNew()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS TypeCHubNew ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND ItemType = 'Dock' " +
                "AND AssetCondition = 'A+ : Brand New';");
            return query.ToString();
        }

        public static string BristolQuery_TypeCHubUsed()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS TypeCHubUsed ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND ItemType = 'Dock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage');");
            return query.ToString();
        }

        public static string BristolQuery_GeomaticsNew()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS GeomaticsNew ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND AssetCondition = 'A+ : Brand New' " +
                "AND ItemType = 'Laptop' AND Manufacturer = 'Dell' AND (ItemCode LIKE '%75%' OR ItemCode LIKE '%55%');");
            return query.ToString();
        }

        public static string BristolQuery_GeomaticsUsed()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS GeomaticsUsed ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND ItemType = 'Laptop' AND Manufacturer = 'Dell' AND (ItemCode LIKE '%75%' OR ItemCode LIKE '%55%');");
            return query.ToString();
        }

        public static string BristolQuery_GeomaticsUSBC()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS GeomaticsUSBC ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND HasCharger = 'No' " +
                "AND ItemType = 'Laptop' AND Manufacturer = 'Dell' AND (ItemCode LIKE '%75%' OR ItemCode LIKE '%55%');");
            return query.ToString();
        }

        public static string BristolQuery_FinanceNew()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS FinanceNew ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND AssetCondition = 'A+ : Brand New' " +
                "AND ItemType = 'Laptop' AND Manufacturer = 'Dell' AND (ItemCode LIKE '%3530%' OR ItemCode LIKE '%5520%' OR ItemCode LIKE '%5530%');");
            return query.ToString();
        }

        public static string BristolQuery_FinanceUsed()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS FinanceUsed ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND ItemType = 'Laptop' AND Manufacturer = 'Dell' AND (ItemCode LIKE '%3530%' OR ItemCode LIKE '%5520%' OR ItemCode LIKE '%5530%');");
            return query.ToString();
        }

        public static string BristolQuery_FinanceUSBC()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS FinanceUSBC ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND HasCharger = 'No' " +
                "AND ItemType = 'Laptop' AND Manufacturer = 'Dell' AND (ItemCode LIKE '%3530%' OR ItemCode LIKE '%5520%' OR ItemCode LIKE '%5530%');");
            return query.ToString();
        }

        public static string BristolQuery_StandardNew()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS StandardNew ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND AssetCondition = 'A+ : Brand New' " +
                "AND ((ItemType = 'Laptop' AND Manufacturer = 'Dell' AND (ItemCode LIKE '%54%')) " +
                "OR (ItemType = 'Laptop' AND Manufacturer = 'Lenovo'));");
            return query.ToString();
        }

        public static string BristolQuery_StandardUsed()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS StandardUsed ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND ((ItemType = 'Laptop' AND Manufacturer = 'Dell' AND (ItemCode LIKE '%54%')) " +
                "OR (ItemType = 'Laptop' AND Manufacturer = 'Lenovo'));");
            return query.ToString();
        }

        public static string BristolQuery_DellUSBC()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS DellUSBC ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND HasCharger = 'No' " +
                "AND ItemType = 'Laptop' AND Manufacturer = 'Dell' AND (ItemCode LIKE '%54%');");
            return query.ToString();
        }

        public static string BristolQuery_LenovoUSBC()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS LenovoUSBC ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND HasCharger = 'No' " +
                "AND ItemType = 'Laptop' AND Manufacturer = 'Lenovo';");
            return query.ToString();
        }

        public static string BristolQuery_SamsungNew()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS SamsungNew ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND AssetCondition = 'A+ : Brand New' " +
                "AND ItemType = 'Mobile Phone' AND Manufacturer = 'Samsung' ;");
            return query.ToString();
        }

        public static string BristolQuery_SamsungUsed()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS SamsungUsed ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND ItemType = 'Mobile Phone' AND Manufacturer = 'Samsung' ;");
            return query.ToString();
        }

        public static string BristolQuery_SamsungChargers()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS SamsungChargers ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND HasCharger = 'No' " +
                "AND ItemType = 'Mobile Phone' AND Manufacturer = 'Samsung' ;");
            return query.ToString();
        }

        public static string BristolQuery_AppleNew()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS AppleNew ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND AssetCondition = 'A+ : Brand New' " +
                "AND ItemType = 'Mobile Phone' AND Manufacturer = 'Apple' AND (ItemCode LIKE '%SE%') ;");
            return query.ToString();
        }

        public static string BristolQuery_AppleUsed()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS AppleUsed ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND ItemType = 'Mobile Phone' AND Manufacturer = 'Apple' AND (ItemCode LIKE '%SE%') ;");
            return query.ToString();
        }

        public static string BristolQuery_AppleChargers()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS AppleChargers ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND HasCharger = 'No' " +
                "AND ItemType = 'Mobile Phone' AND Manufacturer = 'Apple' AND (ItemCode LIKE '%SE%') ;");
            return query.ToString();
        }

        public static string BristolQuery_63WCharger()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS 63WCharger ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND ItemType = 'Charger' AND (ItemCode LIKE '%63W%') ;");
            return query.ToString();
        }

        public static string BristolQuery_DellAdapter()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS DellAdapter ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND ItemType = 'Adapter' AND Manufacturer = 'Dell' ;");
            return query.ToString();
        }

        public static string BristolQuery_LenovoAdapter()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS LenovoAdapter ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND ItemType = 'Adapter' AND Manufacturer = 'Lenovo' ;");
            return query.ToString();
        }

        public static string BristolQuery_LaptopBag()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS LaptopBag ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND ItemType = 'Laptop bag/Rucksack' ;");
            return query.ToString();
        }

        public static string BristolQuery_MobileSIM()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS MobileSIM ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 161 AND AssetStatus = 'Stock' " +
                "AND ItemType = 'Mobile SIM' ;");
            return query.ToString();
        }
        #endregion

        #region Kent
        public static string KentQuery_KeyboardNew()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS KeyboardNew ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND AssetCondition = 'A+ : Brand New' " +
                "AND ItemType = 'Keyboard/Mouse' ;");
            return query.ToString();
        }

        public static string KentQuery_KeyboardUsed()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS KeyboardUsed ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND ItemType = 'Keyboard/Mouse' ;");
            return query.ToString();
        }

        public static string KentQuery_HeadsetsNew()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS HeadsetsNew ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND AssetCondition = 'A+ : Brand New' " +
                "AND ItemType = 'Headsets/earphones' ;");
            return query.ToString();
        }

        public static string KentQuery_HeadsetsUsed()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS HeadsetsUsed ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND ItemType = 'Headsets/earphones' ;");
            return query.ToString();
        }

        public static string KentQuery_BAUNew()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS BAUNew ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND AssetCondition = 'A+ : Brand New' " +
                "AND ItemType = 'Laptop' AND (ItemCode LIKE '%BAU%') ;");
            return query.ToString();
        }

        public static string KentQuery_BAUUsed()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS BAUUsed ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND ItemType = 'Laptop' AND (ItemCode LIKE '%BAU%') ;");
            return query.ToString();
        }

        public static string KentQuery_BAUChargers()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS BAUChargers ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND HasCharger = 'No' " +
                "AND ItemType = 'Laptop' AND (ItemCode LIKE '%BAU%') ;");
            return query.ToString();
        }

        public static string KentQuery_StockLaptopsNew()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS StockLaptopsNew ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND AssetCondition = 'A+ : Brand New' " +
                "AND ItemType = 'Laptop' AND (ItemCode NOT LIKE '%BAU%') ;");
            return query.ToString();
        }

        public static string KentQuery_StockLaptopsUsed()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS StockLaptopsUsed ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND ItemType = 'Laptop' AND (ItemCode NOT LIKE '%BAU%') ;");
            return query.ToString();
        }

        public static string KentQuery_StockLaptopsChargers()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS StockLaptopsChargers ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND HasCharger = 'No' " +
                "AND ItemType = 'Laptop' AND (ItemCode NOT LIKE '%BAU%') ;");
            return query.ToString();
        }

        public static string KentQuery_MonitorNew()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS MonitorNew ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND AssetCondition = 'A+ : Brand New' " +
                "AND ItemType = 'Monitor' ;");
            return query.ToString();
        }

        public static string KentQuery_MonitorUsed()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS MonitorUsed ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND ItemType = 'Monitor' ;");
            return query.ToString();
        }

        public static string KentQuery_MobilePhoneNew()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS MobilePhoneNew ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND AssetCondition = 'A+ : Brand New' " +
                "AND ItemType = 'Mobile Phone' ;");
            return query.ToString();
        }

        public static string KentQuery_MobilePhoneUsed()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS MobilePhoneUsed ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND ItemType = 'Mobile Phone' ;");
            return query.ToString();
        }

        public static string KentQuery_MobilePhoneChargers()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS MobilePhoneChargers ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND HasCharger = 'No' " +
                "AND ItemType = 'Mobile Phone' ;");
            return query.ToString();
        }

        public static string KentQuery_MobileSim()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS MobileSim ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND ItemType = 'Mobile Sim' ;");
            return query.ToString();
        }
        #endregion

        #region Queries of Low stock triggers
        public static string MKQuery_LowStockLaptop()
        {
            StringBuilder query = new StringBuilder();
            query.Append(" SELECT COUNT(AssetTag) AS LowStockLaptop ");
            query.Append(" FROM client_asset_register ");
            query.Append(" WHERE ClientID = 287 AND AssetStatus = 'Stock' " +
                "AND (AssetCondition = 'A+ : Brand New' OR AssetCondition = 'A- : As New' OR AssetCondition = 'B : Cosmetic Damage') " +
                "AND (ItemCode NOT LIKE '%chrome%' AND ItemDescription NOT LIKE '%chrome%') " +
                "AND (AssetLocation != 'A1 - Boston Secure storage' AND AssetLocation != 'A1 - Dublin Secure storage') " +
                "AND ItemType = 'Laptop' ;");
            return query.ToString();
        }
        #endregion
    }
}
