using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

namespace Data2Sheet
{
    /// <summary>
    /// Creating the Excel file for the application.
    /// </summary>
    public class ExcelFile
    {
        private readonly string _filepath;

        public ExcelFile (string filepath)
        {
            _filepath = filepath;
        }

        private void AddCellString(SheetData sheetData, string columnIndex, UInt32 rowIndex, string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                var row = new Row() { RowIndex = rowIndex };

                Cell heading = new Cell
                {
                    CellReference = columnIndex + (rowIndex),
                    CellValue = new CellValue(value),
                    DataType = new EnumValue<CellValues>(CellValues.String)                    
                };

                row.AppendChild(heading);

                sheetData.AppendChild(row);
            }
        }

        private void AddCellInt(SheetData sheetData, string columnIndex, UInt32 rowIndex, string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                var row = new Row() { RowIndex = rowIndex };

                Cell heading = new Cell
                {
                    CellReference = columnIndex + (rowIndex),
                    CellValue = new CellValue(value),
                    DataType = new EnumValue<CellValues>(CellValues.Number)
                };

                row.AppendChild(heading);

                sheetData.AppendChild(row);
            }
        }

        public void CreateStockReport(DataTable dt, string sheetName, string clientName)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            // https://learn.microsoft.com/en-us/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(_filepath, SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetName                
                };

                sheets.Append(sheet);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                UInt32 rowIndex = 1;
      
                AddCellString(sheetData, "A", rowIndex, "Client Name");
                AddCellString(sheetData, "D", rowIndex, clientName);
                AddCellString(sheetData, "D", 2, "Data From: " + DateTime.Now.ToString("F"));

                rowIndex = 4;
                AddCellString(sheetData, "A", rowIndex, "Asset Tag");
                AddCellString(sheetData, "B", rowIndex, "Item Type");
                AddCellString(sheetData, "C", rowIndex, "Item Code");
                AddCellString(sheetData, "D", rowIndex, "Item Description");
                AddCellString(sheetData, "E", rowIndex, "Manufacturer");
                AddCellString(sheetData, "F", rowIndex, "Status");
                AddCellString(sheetData, "G", rowIndex, "Location");
                AddCellString(sheetData, "H", rowIndex, "Condition");
                AddCellString(sheetData, "I", rowIndex, "Re Issue");
                AddCellString(sheetData, "J", rowIndex, "Has Charger");
                AddCellString(sheetData, "K", rowIndex, "CMAR");
                AddCellString(sheetData, "L", rowIndex, "PO Number");
                AddCellString(sheetData, "M", rowIndex, "Purchase Date");
                AddCellString(sheetData, "N", rowIndex, "Warranty Start");
                AddCellString(sheetData, "O", rowIndex, "Warranty End");
                AddCellString(sheetData, "P", rowIndex, "Last Audited");

                rowIndex = 5;
                foreach (DataRow d in dt.Rows)
                {
                    AddCellInt(sheetData, "A", rowIndex, d["Asset Tag"].ToString());
                    AddCellString(sheetData, "B", rowIndex, d["Item Type"].ToString());
                    AddCellString(sheetData, "C", rowIndex, d["Item Code"].ToString());
                    AddCellString(sheetData, "D", rowIndex, d["Item Description"].ToString());

                    AddCellString(sheetData, "E", rowIndex, d["Manufacturer"].ToString());
                    AddCellString(sheetData, "F", rowIndex, d["Status"].ToString());
                    AddCellString(sheetData, "G", rowIndex, d["Location"].ToString());
                    AddCellString(sheetData, "H", rowIndex, d["Condition"].ToString());

                    AddCellString(sheetData, "I", rowIndex, d["Re Issue"].ToString());
                    AddCellString(sheetData, "J", rowIndex, d["Has Charger"].ToString());
                    AddCellString(sheetData, "K", rowIndex, d["CMAR"].ToString());
                    AddCellString(sheetData, "L", rowIndex, d["PO Number"].ToString());

                    AddCellString(sheetData, "M", rowIndex, d["Purchase Date"].ToString());
                    AddCellString(sheetData, "N", rowIndex, d["Warranty Start"].ToString());
                    AddCellString(sheetData, "O", rowIndex, d["Warranty End"].ToString());
                    AddCellString(sheetData, "P", rowIndex, d["Last Audited"].ToString());

                    rowIndex++;
                }

                workbookpart.Workbook.Save();
            }
        }

        public void CreateEdinburghReport(DataTable dt, string sheetName, string clientName)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            // https://learn.microsoft.com/en-us/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(_filepath, SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetName
                };

                sheets.Append(sheet);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                UInt32 rowIndex = 1;

                AddCellString(sheetData, "A", rowIndex, "Client Name");
                AddCellString(sheetData, "D", rowIndex, clientName);
                AddCellString(sheetData, "D", 2, "Data From: " + DateTime.Now.ToString("F"));

                rowIndex = 4;
                AddCellString(sheetData, "A", rowIndex, "Dispatch date");
                AddCellString(sheetData, "B", rowIndex, "Ticket");
                AddCellString(sheetData, "C", rowIndex, "Item type");
                AddCellString(sheetData, "D", rowIndex, "Assigned user");
                AddCellString(sheetData, "E", rowIndex, "Asset Tag");
                AddCellString(sheetData, "F", rowIndex, "Serial number");
                AddCellString(sheetData, "G", rowIndex, "IMEI");                

                rowIndex = 5;
                foreach (DataRow d in dt.Rows)
                {
                    AddCellString(sheetData, "A", rowIndex, d["DespatchDate"].ToString());
                    AddCellInt(sheetData, "B", rowIndex, d["OrderRef"].ToString());
                    AddCellString(sheetData, "C", rowIndex, d["ItemType"].ToString());
                    AddCellInt(sheetData, "D", rowIndex, d["UserID"].ToString());
                    AddCellInt(sheetData, "E", rowIndex, d["AssetTag"].ToString());
                    AddCellString(sheetData, "F", rowIndex, d["SerialNo"].ToString());
                    AddCellInt(sheetData, "G", rowIndex, d["IMEI"].ToString());

                    rowIndex++;
                }

                workbookpart.Workbook.Save();
            }
        }

        public void CreateMKReport(DataTable dt)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            // https://learn.microsoft.com/en-us/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(_filepath, SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "MK stock report"
                };

                sheets.Append(sheet);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                UInt32 rowIndex = 1;

                AddCellString(sheetData, "A", rowIndex, "Client Name");
                AddCellString(sheetData, "D", rowIndex, "MK Office");
                AddCellString(sheetData, "D", 2, "Data From: " + DateTime.Now.ToString("F"));

                AddCellString(sheetData, "A", 3, "UK Stock");
                AddCellString(sheetData, "A", 4, "Item");
                AddCellString(sheetData, "B", 4, "New in stock");
                AddCellString(sheetData, "C", 4, "Used in stock");
                AddCellString(sheetData, "D", 4, "Chargers required");
                AddCellString(sheetData, "A", 5, "Chromebooks");

                foreach (DataRow d in dt.Rows)
                {
                    AddCellInt(sheetData, "B", 5, d["UKChromebooksNew"].ToString());
                    AddCellInt(sheetData, "C", 5, d["UKChromebooksUsed"].ToString());
                    AddCellString(sheetData, "A", 6, "Dell laptops");
                    AddCellInt(sheetData, "B", 6, d["UKDellNew"].ToString());
                    AddCellInt(sheetData, "C", 6, d["UKDellUsed"].ToString());
                    AddCellString(sheetData, "D", 6, "(" + d["UKDell65W"].ToString() + ") Dell 65-Watt USB-C AC Adapter - UK");
                    AddCellString(sheetData, "E", 6, "(" + d["UKDell130W"].ToString() + ") Dell USB-C 130 W AC Adapter");
                    AddCellString(sheetData, "A", 7, "HP laptops");
                    AddCellInt(sheetData, "B", 7, d["UKHPNew"].ToString());
                    AddCellInt(sheetData, "C", 7, d["UKHPUsed"].ToString());
                    AddCellString(sheetData, "D", 7, "(" + d["UKHP45W"].ToString() + ") 710412-001 Smart AC power adapter (45 watt) 4.5mm barrel connector");
                    AddCellString(sheetData, "A", 8, "Lenovo laptops");
                    AddCellInt(sheetData, "B", 8, d["UKLenovoNew"].ToString());
                    AddCellInt(sheetData, "C", 8, d["UKLenovoUsed"].ToString());
                    AddCellString(sheetData, "D", 8, "(" + d["UKLenovo65W"].ToString() + ") Lenovo 65W Standard AC Adapter (USB Type-C)- UK/Ireland");
                    AddCellString(sheetData, "A", 9, "Total laptops");
                    int TotalLaptops = Convert.ToInt32(d["UKChromebooksNew"]) + Convert.ToInt32(d["UKChromebooksUsed"]) + Convert.ToInt32(d["UKDellNew"]) + Convert.ToInt32(d["UKDellUsed"]) + Convert.ToInt32(d["UKHPNew"]) + Convert.ToInt32(d["UKHPUsed"]) + Convert.ToInt32(d["UKLenovoNew"]) + Convert.ToInt32(d["UKLenovoUsed"]);
                    AddCellInt(sheetData, "B", 9, TotalLaptops.ToString());
                    AddCellString(sheetData, "A", 10, "Mobile phones");
                    AddCellInt(sheetData, "B", 10, d["UKMobileNew"].ToString());
                    AddCellInt(sheetData, "C", 10, d["UKMobileUsed"].ToString());
                    AddCellString(sheetData, "D", 10, "(" + d["UKMobileCharger"].ToString() + ") Mobile phone USB-C charger");
                    AddCellString(sheetData, "A", 11, "SIM cards");
                    AddCellInt(sheetData, "B", 11, d["UKSIMCards"].ToString());
                    AddCellString(sheetData, "A", 12, "Dell 65-Watt USB-C AC Adapter - UK");
                    AddCellInt(sheetData, "B", 12, d["UKDellCharger65"].ToString());
                    AddCellString(sheetData, "A", 13, "Dell USB-C 130 W AC Adapter");
                    AddCellInt(sheetData, "B", 13, d["UKDellCharger130"].ToString());
                    AddCellString(sheetData, "A", 14, "HP 710412-001 Smart AC power adapter (45 watt) 4.5mm barrel connector");
                    AddCellInt(sheetData, "B", 14, d["UKHPCharger"].ToString());
                    AddCellString(sheetData, "A", 15, "Lenovo 65W Standard AC Adapter (USB Type-C)- UK/Ireland");
                    AddCellInt(sheetData, "B", 15, d["UKLenovoCharger"].ToString());
                    AddCellString(sheetData, "A", 16, "Mobile phone USB-C charger");
                    AddCellInt(sheetData, "B", 16, d["UKMobileUSBCCharger"].ToString());
                    
                    AddCellString(sheetData, "A", 18, "US Stock");
                    AddCellString(sheetData, "A", 19, "Item");
                    AddCellString(sheetData, "B", 19, "New in stock");
                    AddCellString(sheetData, "C", 19, "Used in stock");
                    AddCellString(sheetData, "D", 19, "Chargers required");
                    AddCellString(sheetData, "A", 20, "Chromebooks");
                    AddCellInt(sheetData, "B", 20, d["USChromebooksNew"].ToString());
                    AddCellInt(sheetData, "C", 20, d["USChromebooksUsed"].ToString());
                    AddCellString(sheetData, "A", 21, "Dell laptops");
                    AddCellInt(sheetData, "B", 21, d["USDellNew"].ToString());
                    AddCellInt(sheetData, "C", 21, d["USDellUsed"].ToString());
                    AddCellString(sheetData, "D", 21, "(" + d["USDell65W"].ToString() + ") Dell 65-Watt USB-C AC Adapter");
                    AddCellString(sheetData, "E", 21, "(" + d["USDell130W"].ToString() + ") Dell USB-C 130 W AC Adapter");
                    AddCellString(sheetData, "A", 22, "HP laptops");
                    AddCellInt(sheetData, "B", 22, d["USHPNew"].ToString());
                    AddCellInt(sheetData, "C", 22, d["USHPUsed"].ToString());
                    AddCellString(sheetData, "D", 22, "(" + d["USHP45W"].ToString() + ") 710412-001 Smart AC power adapter (45 watt) 4.5mm barrel connector");
                    AddCellString(sheetData, "A", 23, "Lenovo laptops");
                    AddCellInt(sheetData, "B", 23, d["USLenovoNew"].ToString());
                    AddCellInt(sheetData, "C", 23, d["USLenovoUsed"].ToString());
                    AddCellString(sheetData, "D", 23, "(" + d["USLenovo65W"].ToString() + ") Lenovo 65W Standard AC Adapter (USB Type-C)");
                    AddCellString(sheetData, "A", 24, "Total laptops");
                    TotalLaptops = Convert.ToInt32(d["USChromebooksNew"]) + Convert.ToInt32(d["USChromebooksUsed"]) + Convert.ToInt32(d["USDellNew"]) + Convert.ToInt32(d["USDellUsed"]) + Convert.ToInt32(d["USHPNew"]) + Convert.ToInt32(d["USHPUsed"]) + Convert.ToInt32(d["USLenovoNew"]) + Convert.ToInt32(d["USLenovoUsed"]);
                    AddCellInt(sheetData, "B", 24, TotalLaptops.ToString());
                    AddCellString(sheetData, "A", 25, "Mobile phones");
                    AddCellInt(sheetData, "B", 25, d["USMobileNew"].ToString());
                    AddCellInt(sheetData, "C", 25, d["USMobileUsed"].ToString());
                    AddCellString(sheetData, "D", 25, "(" + d["USMobileCharger"].ToString() + ") Mobile phone USB-C charger");

                    AddCellString(sheetData, "A", 27, "EU Stock");
                    AddCellString(sheetData, "A", 28, "Item");
                    AddCellString(sheetData, "B", 28, "New in stock");
                    AddCellString(sheetData, "C", 28, "Used in stock");
                    AddCellString(sheetData, "D", 28, "Chargers required");
                    AddCellString(sheetData, "A", 29, "Chromebooks");
                    AddCellInt(sheetData, "B", 29, d["EUChromebooksNew"].ToString());
                    AddCellInt(sheetData, "C", 29, d["EUChromebooksUsed"].ToString());
                    AddCellString(sheetData, "A", 30, "Dell laptops");
                    AddCellInt(sheetData, "B", 30, d["EUDellNew"].ToString());
                    AddCellInt(sheetData, "C", 30, d["EUDellUsed"].ToString());
                    AddCellString(sheetData, "D", 30, "(" + d["EUDell65W"].ToString() + ") Dell 65-Watt USB-C AC Adapter");
                    AddCellString(sheetData, "E", 30, "(" + d["EUDell130W"].ToString() + ") Dell USB-C 130 W AC Adapter");
                    AddCellString(sheetData, "A", 31, "HP laptops");
                    AddCellInt(sheetData, "B", 31, d["EUHPNew"].ToString());
                    AddCellInt(sheetData, "C", 31, d["EUHPUsed"].ToString());
                    AddCellString(sheetData, "D", 31, "(" + d["EUHP45W"].ToString() + ") 710412-001 Smart AC power adapter (45 watt) 4.5mm barrel connector");
                    AddCellString(sheetData, "A", 32, "Lenovo laptops");
                    AddCellInt(sheetData, "B", 32, d["EULenovoNew"].ToString());
                    AddCellInt(sheetData, "C", 32, d["EULenovoUsed"].ToString());
                    AddCellString(sheetData, "D", 32, "(" + d["EULenovo65W"].ToString() + ") Lenovo 65W Standard AC Adapter (USB Type-C)");
                    AddCellString(sheetData, "A", 33, "Total laptops");
                    TotalLaptops = Convert.ToInt32(d["EUChromebooksNew"]) + Convert.ToInt32(d["EUChromebooksUsed"]) + Convert.ToInt32(d["EUDellNew"]) + Convert.ToInt32(d["EUDellUsed"]) + Convert.ToInt32(d["EUHPNew"]) + Convert.ToInt32(d["EUHPUsed"]) + Convert.ToInt32(d["EULenovoNew"]) + Convert.ToInt32(d["EULenovoUsed"]);
                    AddCellInt(sheetData, "B", 33, TotalLaptops.ToString());
                    AddCellString(sheetData, "A", 34, "Mobile phones");
                    AddCellInt(sheetData, "B", 34, d["EUMobileNew"].ToString());
                    AddCellInt(sheetData, "C", 34, d["EUMobileUsed"].ToString());
                    AddCellString(sheetData, "D", 34, "(" + d["EUMobileCharger"].ToString() + ") Mobile phone USB-C charger");
                }

                workbookpart.Workbook.Save();
            }
            
            MergeCells(_filepath, "MK stock report", "B9", "C9"); //Total laptops UK
            MergeCells(_filepath, "MK stock report", "B11", "C11"); //SIM cards
            MergeCells(_filepath, "MK stock report", "B12", "C12"); //Dell 65-Watt USB-C AC Adapter - UK
            MergeCells(_filepath, "MK stock report", "B13", "C13"); //Dell USB-C 130 W AC Adapter
            MergeCells(_filepath, "MK stock report", "B14", "C14"); //HP 710412-001 Smart AC power adapter (45 watt) 4.5mm barrel connector
            MergeCells(_filepath, "MK stock report", "B15", "C15"); //Lenovo 65W Standard AC Adapter (USB Type-C)- UK/Ireland
            MergeCells(_filepath, "MK stock report", "B16", "C16");//Mobile phone USB-C charger
            MergeCells(_filepath, "MK stock report", "B24", "C24"); //Total laptops US
            MergeCells(_filepath, "MK stock report", "B33", "C33"); //Total laptops EU
        }

        public void CreateBristolReport(DataTable dt)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            // https://learn.microsoft.com/en-us/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(_filepath, SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Bristol Stock report"
                };

                sheets.Append(sheet);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                UInt32 rowIndex = 1;

                AddCellString(sheetData, "A", rowIndex, "Client Name");
                AddCellString(sheetData, "D", rowIndex, "Bristol Office");
                AddCellString(sheetData, "D", 2, "Data From: " + DateTime.Now.ToString("F"));

                AddCellString(sheetData, "A", 4, "Item");
                AddCellString(sheetData, "B", 4, "New Qty");
                AddCellString(sheetData, "C", 4, "Used Qty");
                AddCellString(sheetData, "D", 4, "Chargers required");
                AddCellString(sheetData, "A", 5, "Keyboard/mouse sets");

                foreach (DataRow d in dt.Rows)
                {
                    AddCellInt(sheetData, "B", 5, d["KeyboardNew"].ToString());
                    AddCellInt(sheetData, "C", 5, d["KeyboardUsed"].ToString());
                    AddCellString(sheetData, "A", 6, "Headsets");
                    AddCellInt(sheetData, "B", 6, d["HeadsetsNew"].ToString());
                    AddCellInt(sheetData, "C", 6, d["HeadsetsUsed"].ToString());
                    AddCellString(sheetData, "A", 7, "Phone cases");
                    AddCellInt(sheetData, "B", 7, d["PhoneCase"].ToString());                   
                    AddCellString(sheetData, "A", 8, "Mobile screen protectors");
                    AddCellInt(sheetData, "B", 8, d["Protector"].ToString());
                    AddCellString(sheetData, "A", 9, "6-in-1 USB Type-C Hub (for Geomatics)");
                    AddCellInt(sheetData, "B", 9, d["TypeCHubNew"].ToString());
                    AddCellInt(sheetData, "C", 9, d["TypeCHubUsed"].ToString());
                    AddCellString(sheetData, "A", 10, "Geomatics laptops");
                    AddCellInt(sheetData, "B", 11, d["GeomaticsNew"].ToString());
                    AddCellInt(sheetData, "C", 11, d["GeomaticsUsed"].ToString());
                    AddCellString(sheetData, "D", 11, "(" + d["GeomaticsUSBC"].ToString() + ") Dell USB-C");
                    AddCellString(sheetData, "A", 11, "DELL precision 7560");
                    AddCellString(sheetData, "A", 12, "Dell Precision 5560/5550/5570");
                    AddCellString(sheetData, "A", 13, "Finance laptops");
                    AddCellInt(sheetData, "B", 14, d["FinanceNew"].ToString());
                    AddCellInt(sheetData, "C", 14, d["FinanceUsed"].ToString());
                    AddCellString(sheetData, "D", 14, "(" + d["FinanceUSBC"].ToString() + ") Dell USB-C");
                    AddCellString(sheetData, "A", 14, "Dell Precision 3530");
                    AddCellString(sheetData, "A", 15, "Dell latitude 5520/5530");
                    AddCellString(sheetData, "A", 16, "Standard spec laptops");
                    AddCellString(sheetData, "A", 17, "Dell Latitude 5410/5420/5430");
                    AddCellInt(sheetData, "B", 17, d["StandardNew"].ToString());
                    AddCellInt(sheetData, "C", 17, d["StandardUsed"].ToString());
                    AddCellString(sheetData, "D", 17, "(" + d["DellUSBC"].ToString() + ") Dell USB-C");
                    AddCellString(sheetData, "A", 18, "Lenovo laptops");
                    AddCellString(sheetData, "D", 18, "(" + d["LenovoUSBC"].ToString() + ") Lenovo USB-C");
                    AddCellString(sheetData, "A", 19, "Samsung mobiles");
                    AddCellInt(sheetData, "B", 19, d["SamsungNew"].ToString());
                    AddCellInt(sheetData, "C", 19, d["SamsungUsed"].ToString());
                    AddCellInt(sheetData, "D", 19, d["SamsungChargers"].ToString());
                    AddCellString(sheetData, "A", 20, "Apple iPhone SE");
                    AddCellInt(sheetData, "B", 20, d["AppleNew"].ToString());
                    AddCellInt(sheetData, "C", 20, d["AppleUsed"].ToString());
                    AddCellInt(sheetData, "D", 20, d["AppleChargers"].ToString());
                    AddCellString(sheetData, "A", 21, "63W two port GaN wall charger (for mobiles)");
                    AddCellInt(sheetData, "B", 21, d["63WCharger"].ToString());
                    AddCellString(sheetData, "A", 22, "AC Adapter USB type C 65W – Dell Chargers");
                    AddCellInt(sheetData, "B", 22, d["DellAdapter"].ToString());
                    AddCellString(sheetData, "A", 23, "AC Adapter USB type C 65W – Lenovo Chargers");
                    AddCellInt(sheetData, "B", 23, d["LenovoAdapter"].ToString());
                    AddCellString(sheetData, "A", 24, "Laptop bags");
                    AddCellInt(sheetData, "B", 24, d["LaptopBag"].ToString());
                    AddCellString(sheetData, "A", 25, "SIM cards");
                    AddCellInt(sheetData, "B", 25, d["MobileSIM"].ToString());
                }

                workbookpart.Workbook.Save();
            }

            MergeCells(_filepath, "Bristol Stock report", "B7", "C7"); //Phone cases
            MergeCells(_filepath, "Bristol Stock report", "B8", "C8"); //Screen protectors
            MergeCells(_filepath, "Bristol Stock report", "B11", "B12"); //GeomaticsNew
            MergeCells(_filepath, "Bristol Stock report", "C11", "C12"); //GeomaticsUsed
            MergeCells(_filepath, "Bristol Stock report", "D11", "D12"); //GeomaticsUSBC
            MergeCells(_filepath, "Bristol Stock report", "B14", "B15"); //FinanceNew
            MergeCells(_filepath, "Bristol Stock report", "C14", "C15"); //FinanceUsed
            MergeCells(_filepath, "Bristol Stock report", "D14", "D15"); //FinanceUSBC
            MergeCells(_filepath, "Bristol Stock report", "B17", "B18"); //StandardNew
            MergeCells(_filepath, "Bristol Stock report", "C17", "C18"); //StandardUsed
        }

        public void CreateKentReport(DataTable dt)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            // https://learn.microsoft.com/en-us/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(_filepath, SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Kent Stock report"
                };

                sheets.Append(sheet);
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                UInt32 rowIndex = 1;

                AddCellString(sheetData, "A", rowIndex, "Client Name");
                AddCellString(sheetData, "D", rowIndex, "Kent Office");
                AddCellString(sheetData, "D", 2, "Data From: " + DateTime.Now.ToString("F"));

                AddCellString(sheetData, "A", 4, "Item");
                AddCellString(sheetData, "B", 4, "New Qty");
                AddCellString(sheetData, "C", 4, "Used Qty");
                AddCellString(sheetData, "D", 4, "Chargers required");
                AddCellString(sheetData, "A", 5, "Keyboard/mouse sets");

                foreach (DataRow d in dt.Rows)
                {
                    AddCellInt(sheetData, "B", 5, d["KeyboardNew"].ToString());
                    AddCellInt(sheetData, "C", 5, d["KeyboardUsed"].ToString());
                    AddCellString(sheetData, "A", 6, "Headsets");
                    AddCellInt(sheetData, "B", 6, d["HeadsetsNew"].ToString());
                    AddCellInt(sheetData, "C", 6, d["HeadsetsUsed"].ToString());
                    AddCellString(sheetData, "A", 7, "BAU loan laptops");
                    AddCellInt(sheetData, "B", 7, d["BAUNew"].ToString());
                    AddCellInt(sheetData, "C", 7, d["BAUUsed"].ToString());
                    AddCellString(sheetData, "D", 7, "(" + d["BAUChargers"].ToString() + ") Dell USB-C");
                    AddCellString(sheetData, "A", 8, "Stock laptops");
                    AddCellInt(sheetData, "B", 8, d["StockLaptopsNew"].ToString());
                    AddCellInt(sheetData, "C", 8, d["StockLaptopsUsed"].ToString());
                    AddCellString(sheetData, "D", 8, "(" + d["StockLaptopsChargers"].ToString() + ") Dell USB-C");
                    AddCellString(sheetData, "A", 9, "Monitors");
                    AddCellInt(sheetData, "B", 9, d["MonitorNew"].ToString());
                    AddCellInt(sheetData, "C", 9, d["MonitorUsed"].ToString());
                    AddCellString(sheetData, "A", 10, "Mobiles");
                    AddCellInt(sheetData, "B", 10, d["MobilePhoneNew"].ToString());
                    AddCellInt(sheetData, "C", 10, d["MobilePhoneUsed"].ToString());
                    AddCellInt(sheetData, "D", 10, d["MobilePhoneChargers"].ToString());
                    AddCellString(sheetData, "A", 11, "SIM cards");
                    AddCellInt(sheetData, "B", 11, d["MobileSim"].ToString());
                }

                workbookpart.Workbook.Save();
            }

            MergeCells(_filepath, "Kent Stock report", "B11", "C11"); //Mobile Sim
        }

        private void MergeCells(string docName, string sheetName, string cell1Name, string cell2Name)
        {
            ExcelFile.MergeTwoCells(docName, sheetName, cell1Name, cell2Name);
        }

        #region MergeTwoCells
        /// <summary>
        /// Given a document name, a worksheet name, and the names of two adjacent cells, merges the two cells.
        /// When two cells are merged, only the content from one cell is preserved:
        /// the upper-left cell for left-to-right languages or the upper-right cell for right-to-left languages.
        /// https://learn.microsoft.com/en-us/office/open-xml/how-to-merge-two-adjacent-cells-in-a-spreadsheet
        /// </summary>
        /// <param name="docName"></param>
        /// <param name="sheetName"></param>
        /// <param name="cell1Name"></param>
        /// <param name="cell2Name"></param>
        public static void MergeTwoCells(string docName, string sheetName, string cell1Name, string cell2Name)
        {
            // Open the document for editing.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
            {
                Worksheet worksheet = GetWorksheet(document, sheetName);
                if (worksheet == null || string.IsNullOrEmpty(cell1Name) || string.IsNullOrEmpty(cell2Name))
                {
                    return;
                }

                // Verify if the specified cells exist, and if they do not exist, create them.
                CreateSpreadsheetCellIfNotExist(worksheet, cell1Name);
                CreateSpreadsheetCellIfNotExist(worksheet, cell2Name);

                MergeCells mergeCells;
                if (worksheet.Elements<MergeCells>().Count() > 0)
                {
                    mergeCells = worksheet.Elements<MergeCells>().First();
                }
                else
                {
                    mergeCells = new MergeCells();

                    // Insert a MergeCells object into the specified position.
                    if (worksheet.Elements<CustomSheetView>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                    }
                    else if (worksheet.Elements<DataConsolidate>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
                    }
                    else if (worksheet.Elements<SortState>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
                    }
                    else if (worksheet.Elements<AutoFilter>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
                    }
                    else if (worksheet.Elements<Scenarios>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
                    }
                    else if (worksheet.Elements<ProtectedRanges>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
                    }
                    else if (worksheet.Elements<SheetProtection>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
                    }
                    else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
                    }
                    else
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                    }
                }

                // Create the merged cell and append it to the MergeCells collection.
                MergeCell mergeCell = new MergeCell() { Reference = new StringValue(cell1Name + ":" + cell2Name) };
                mergeCells.Append(mergeCell);

                worksheet.Save();
            }
        }
        #endregion

        #region CreateSpreadsheetCellIfNotExist
        /// <summary>
        /// Given a Worksheet and a cell name, verifies that the specified cell exists.
        /// If it does not exist, creates a new cell.
        /// </summary>
        private static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName)
        {
            string columnName = GetColumnName(cellName);
            uint rowIndex = GetRowIndex(cellName);

            IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r.RowIndex.Value == rowIndex);

            // If the Worksheet does not contain the specified row, create the specified row.
            // Create the specified cell in that row, and insert the row into the Worksheet.
            if (rows.Count() == 0)
            {
                Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
                Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                row.Append(cell);
                worksheet.Descendants<SheetData>().First().Append(row);
                worksheet.Save();
            }
            else
            {
                Row row = rows.First();

                IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

                // If the row does not contain the specified cell, create the specified cell.
                if (cells.Count() == 0)
                {
                    Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                    row.Append(cell);
                    worksheet.Save();
                }
            }
        }
        #endregion

        #region GetWorksheet
        /// <summary>
        /// Given a SpreadsheetDocument and a worksheet name, get the specified worksheet.
        /// </summary>
        private static Worksheet GetWorksheet(SpreadsheetDocument document, string worksheetName)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
            if (sheets.Count() == 0)
                return null;
            else
                return worksheetPart.Worksheet;
        }
        #endregion

        #region GetColumnName
        /// <summary>
        /// Given a cell name, parses the specified cell to get the column name.
        /// </summary>
        private static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }
        #endregion

        #region GetWorksheet
        /// <summary>
        /// Given a cell name, parses the specified cell to get the row index.
        /// </summary>
        private static uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }
        #endregion
        
    }
}
