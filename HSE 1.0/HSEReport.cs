using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Globalization;

namespace HSE_1._0
{
    class HSEReport
    {
        Application excelApp = new Microsoft.Office.Interop.Excel.Application();

        class Global
        {
            public static int filesPicked;
            public static int deliveries;
            public static int pickUps;
            public static int pcccBoxes;
            public static int pcccCabinets;
            //public static int pickUpsBalinaslow;
            public static int boxesDelivered;
            public static int[,] numberOfTrips = {
                    {0, 0},
                    {1, 0},
                    {2, 0},
                    {3, 0},
                    {4, 0},
                    {5, 0},
                    {6, 0},
                    {7, 0}
                };

            public static Dictionary<string, int> locations = new Dictionary<string, int>()
                {
                    { "Shantalla", 0 },
                    { "Athenry", 1 },
                    { "Tuam", 2 },
                    { "Loughrea", 3 },
                    { "Doughiska", 4 },
                    { "Mountbellew", 5 },
                    { "Ballinasloe", 6 },
                    { "Mervue", 7 }
                };

            public static List<string> pickList = new List<string>();

            public static string[] endOfWeek;
            public static string[] boxesPicked;
        }

        static void ShipmentsReceits(ref _Worksheet excelSheet)
        {
            Range excelRange = excelSheet.UsedRange;
            /// Remove bg-color, add borders ////
            excelRange.Interior.ColorIndex = 0;
            excelRange.Borders.LineStyle = XlLineStyle.xlContinuous;

            excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[1, 13]].Font.Bold = true;
            excelSheet.Columns["A:A"].ColumnWidth = 11;
            excelSheet.Columns["B:D"].ColumnWidth = 7;
            excelSheet.Columns["E:E"].ColumnWidth = 40;
            excelSheet.Columns["F:F"].ColumnWidth = 20;
            excelSheet.Columns["G:G"].ColumnWidth = 40;
            excelSheet.Columns["H:H"].ColumnWidth = 12;
            excelSheet.Columns["I:I"].ColumnWidth = 25;
            excelSheet.Columns["J:L"].ColumnWidth = 11;
            excelSheet.Columns["M:M"].ColumnWidth = 4;
            excelSheet.Columns["O:O"].ColumnWidth = 13;

            // Gets the Calendar instance associated with a CultureInfo.
            CultureInfo myCI = new CultureInfo("en-US");
            Calendar myCal = myCI.Calendar;

            // Gets the DTFI properties required by GetWeekOfYear.
            CalendarWeekRule myCWR = myCI.DateTimeFormat.CalendarWeekRule;
            DayOfWeek myFirstDOW = myCI.DateTimeFormat.FirstDayOfWeek;

            DateTime LastDay = new System.DateTime(DateTime.Now.Year, 12, 31);

            /// Below block gets week number of the first row ///
            string excelDate = excelSheet.Cells[2, 11].value.ToString();
            var tempDate = DateTime.Parse(excelDate);
            var tempWeek = myCal.GetWeekOfYear(tempDate, myCWR, myFirstDOW) - 1;

            // Color lines by week number
            string reversal;
            var reversalExists = false;
            for (int i = 2; i <= excelRange.Rows.Count; i++)
            {

                /// Below block gets week number of each next "row" ///
                excelDate = excelSheet.Cells[i, 11].value.ToString();
                tempDate = DateTime.Parse(excelDate);
                var tempWeek2 = myCal.GetWeekOfYear(tempDate, myCWR, myFirstDOW) - 1;

                /// compares temporary week (first row) to each next rows week///
                if (tempWeek == tempWeek2)
                {
                    excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 19;
                }
                else
                {
                    excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 20;
                    tempWeek = tempWeek2 + 1;
                }
                /// Check if reversal done on SAP - colors in red////
                reversal = excelSheet.Cells[i, 2].Value;
                reversal.ToString();
                if ((reversal == "102") || (reversal == "602"))
                {
                    excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 46;
                    reversalExists = true;
                }

                if (reversalExists)
                {
                    excelSheet.Cells[excelRange.Rows.Count + 2, 5].Value = "reversal done on SAP";
                    excelSheet.Cells[excelRange.Rows.Count + 2, 5].Interior.ColorIndex = 46;
                }
            }

            Form1 sendMsg = new Form1();

            var tempNum = excelSheet.Cells[2, 1].value;
            int numberOfDeliveries = 0;
            int numberOfPicks = 0;

            List<string> list = new List<string>();

            int numberOfFilesPicked = 0;
            int numberOfBoxesPicked = 0;
            int rowCount = excelRange.Rows.Count;
            for (int i = 2; i <= rowCount; i++)
            {
                //var tempNumNew = ;
                /// compares temporary receipt/ shipment to each next rows///
                if (tempNum != excelSheet.Cells[i, 1].value)
                {
                    // Get Document header text
                    var cellValue = excelSheet.Cells[i - 1, 9].value.ToUpper();
                    // Document header text check if contains location(array)
                    foreach (string key in Global.locations.Keys)
                    {
                        if (cellValue.Contains(key.ToUpper()))
                        {
                            Global.numberOfTrips[Global.locations[key], 1]++;
                            //sendMsg.sendMessage(key + "  "+ locations[key]+ "   If working = "+ numberOfTrips[locations[key], 1]);
                        }
                    }

                    if (cellValue.Contains("FILE"))
                    {
                        int findStr = cellValue.IndexOf("FILE");
                        string number = cellValue.Substring(0, findStr);
                        //sendMsg.sendMessage(number);
                        numberOfFilesPicked += Int32.Parse(number);
                    }

                    if (cellValue.Contains("PCCC"))
                    {
                        int findStr = cellValue.IndexOf("B_PCCC");
                        string number = cellValue.Substring(0, findStr);
                        //sendMsg.sendMessage(number);
                        numberOfBoxesPicked += Int32.Parse(number);
                        //sendMsg.sendMessage("Comentarij - " + cellValue + "  Blok 1 - " + numberOfBoxesPicked.ToString());
                    }

                    if (cellValue.Contains("B_COL"))
                    {
                        int findStr = cellValue.IndexOf("B_COL");
                        string number = cellValue.Substring(0, findStr);
                        //sendMsg.sendMessage(number);
                        Global.pickList.Add(number);
                        var tempVariable = excelSheet.Cells[i - 1, 11].value.ToString(" dd/MM/yyyy");
                        Global.pickList.Add(tempVariable.ToString());
                    }
                    /// IF FALSE - swaps temporary number (receipt/shipment number) and adds blank row
                    if (excelSheet.Cells[i-1,2].value2.Contains("601")) {
                        numberOfDeliveries++;
                    }

                    if (excelSheet.Cells[i-1, 2].value2.Contains("101"))
                    {
                        numberOfPicks++;   
                    }
                    tempNum = excelSheet.Cells[i, 1].value;
                    Range line = (Range)excelSheet.Rows[i];
                    line.Insert();
                    excelSheet.Range[excelSheet.Cells[i, 1], excelSheet.Cells[i, 13]].Interior.ColorIndex = 15;
                    /// ads one extra row to row count (rowNum) due to blank row added ///
                    rowCount += 1;
                }

                // // Document header text check if contains location(array) - last row in excel
                if (i == rowCount)
                {
                    var cellValue = excelSheet.Cells[i, 9].value.ToUpper();
                    foreach (string key in Global.locations.Keys)
                    {
                        if (cellValue.Contains(key.ToUpper()))
                        {
                            Global.numberOfTrips[Global.locations[key], 1]++;
                            //sendMsg.sendMessage(key + "   " + cellValue);
                        }
                    }

                    if (cellValue.Contains("FILE"))
                    {
                        int findStr = cellValue.IndexOf("FILE");
                        string number = cellValue.Substring(0, findStr);
                        //sendMsg.sendMessage(number);
                        numberOfFilesPicked += Int32.Parse(number);
                    }

                    if (cellValue.Contains("PCCC"))
                    {
                        int findStr = cellValue.IndexOf("B_PCCC");
                        string number = cellValue.Substring(0, findStr);
                        //sendMsg.sendMessage(number);
                        numberOfBoxesPicked += Int32.Parse(number);
                        //sendMsg.sendMessage(excelBook.Name + "Comentarij - " + cellValue + "  Blok 2 - " + numberOfBoxesPicked.ToString());
                    }

                    if (cellValue.Contains("B_COL"))
                    {
                        int findStr = cellValue.IndexOf("B_COL");
                        string number = cellValue.Substring(0, findStr);
                        //sendMsg.sendMessage(number);
                        Global.pickList.Add(number);
                        var tempVariable = excelSheet.Cells[i - 1, 11].value.ToString(" dd/MM/yyyy");
                        Global.pickList.Add(tempVariable.ToString());
                    }

                    if (excelSheet.Cells[i, 2].value2.Contains("601"))
                    {
                        numberOfDeliveries++;
                        Global.deliveries = numberOfDeliveries;
                    }

                    if (excelSheet.Cells[i, 2].value2.Contains("101"))
                    {
                        numberOfPicks++;
                        Global.pickUps = numberOfPicks;
                    }
                }
            }

            /*Global.boxesPicked = list.ToArray();*/

            Global.filesPicked = numberOfFilesPicked;
            Global.boxesDelivered = numberOfBoxesPicked;

            

            //sendMsg.sendMessage(numberOfDeliveries.ToString());
            //sendMsg.sendMessage(numberOfPicks.ToString());
        }



        static void Stock(ref _Worksheet excelSheet) {

            Form1 sendMsg = new Form1();

            Range excelRange = excelSheet.UsedRange;

            var filesCount = 0;
            var cabinetCount = 0;
            int rowCount = excelRange.Rows.Count;
            for (int i = 2; i <= rowCount; i++)
            {
                //sendMsg.sendMessage(rowCount.ToString());
                // Get Document header text
                var cellValue = excelSheet.Cells[i, 1].value.ToUpper();

                if (cellValue.Contains("FILES") || cellValue.Contains("BOX")) {
                    filesCount++;
                }
                if (cellValue.Contains("CABINET"))
                {
                    //sendMsg.sendMessage(cellValue);
                    cabinetCount++;
                } 
                //no need to delete any
                /*else if (!cellValue.Contains("CABINET") && !cellValue.Contains("FILES")) {
                    /// IF FALSE - swaps temporary number (receipt/shipment number) and adds blank row
                    Range line = (Range)excelSheet.Rows[i];
                    line.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp); ;
                    rowCount--;
                    i--;
                    //sendMsg.sendMessage(cellValue);
                }*/

            }

            excelSheet.Cells[rowCount + 1, 9].Value2 = filesCount;
            excelSheet.Cells[rowCount + 2, 9].Value2 = cabinetCount;
            Global.pcccBoxes = filesCount;
            Global.pcccCabinets = cabinetCount;

    }



        static int GetNumberOfWeeks(DateTime date)
        {
            //DateTime thisDay = DateTime.Today;

            int year = date.Year;
            int month = date.Month - 1;
            int count = 0;
            List<string> list = new List<string>();
            DateTime startDate = new DateTime(year, month, 1);
            DateTime endDate = startDate.AddMonths(1);
            while (startDate.DayOfWeek != DayOfWeek.Monday)
                startDate = startDate.AddDays(1);
            for (DateTime result = startDate; result < endDate; result = result.AddDays(7))
            {
                
                list.Add(result.AddDays(6).ToString(" dd/MM/yyyy"));
                count++;
            }
            Global.endOfWeek = list.ToArray();
            return count;
        }


        public void openFile(string[] arr, string openBalanceValue)
        {
            Form1 sendMsg = new Form1();
            shipments shipmentsReport = new shipments();

            Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            // New excel workbook into which we will save our worksheets
            Workbook newWorkBook = excelApp.Workbooks.Add();

            /// Open files passed into arr
            for (int arri = 0; arri < arr.Length; arri++) {

                Workbook excelBook = excelApp.Workbooks.Open(arr[arri]);
                _Worksheet excelSheet = excelBook.Sheets[1];
                Range excelRange = excelSheet.UsedRange;

                //sendMsg.sendMessage(arr[arri]);

                string sheetCellValue = excelSheet.Cells[2, 2].value;
                if (sheetCellValue == "102" || sheetCellValue == "101")
                {
                    excelSheet.Name = "Receipt";
                    ShipmentsReceits(ref excelSheet);
                }
                else if (sheetCellValue == "602" || sheetCellValue == "601")
                {
                    excelSheet.Name = "Shipments";
                    ShipmentsReceits(ref excelSheet);
                }
                else if (sheetCellValue == "0002") {
                    excelSheet.Name = "Current stock Boxes & Cabinets";
                    Stock(ref excelSheet);

                }


                Worksheet newWorkSheet = newWorkBook.Worksheets.get_Item(arri + 1);
                excelSheet.Copy(newWorkSheet);
                //sendMsg.sendMessage(excelSheet.ToString());

                excelBook.Save();
                excelBook.Close(true);

            }

            ///////////////// PCCC Invoicing summary sheet ///////////////////////////
            Worksheet summarykSheet = newWorkBook.Worksheets.get_Item(4);
            summarykSheet.Name = "PCCC invoicing summary";

            summarykSheet.Columns["A:A"].ColumnWidth = 26;
            summarykSheet.Columns["B:B"].ColumnWidth = 10;
            summarykSheet.Columns["C:C"].ColumnWidth = 14;

            DateTime thisDay = DateTime.Today;

            summarykSheet.Cells[1,1].Value2 = "Storage Charge " + DateTime.Now.AddMonths(-1).ToString("MMMM") + " " + thisDay.Year;
            summarykSheet.Cells[1, 3].Value2 = GetNumberOfWeeks(thisDay) + " Week Month";

            summarykSheet.Cells[3, 1].Value2 = "Retrieval & Delivery of Files";
            summarykSheet.Cells[4, 2].Value2 = "Qty";

            summarykSheet.Cells[5, 1].Value2 = "No of Files Picked";
            summarykSheet.Cells[5, 2].Value2 = Global.filesPicked;

            summarykSheet.Cells[6, 1].Value2 = "No. Deliveries";
            summarykSheet.Cells[6, 2].Value2 = Global.deliveries;

            summarykSheet.Cells[7, 1].Value2 = "No. of Pick Ups";
            summarykSheet.Cells[7, 2].Value2 = Global.pickUps;

            int cell = 8;
            foreach (string key in Global.locations.Keys)
            {
                summarykSheet.Cells[cell, 1].Value2 = "Delivery from/to " + key;
                summarykSheet.Cells[cell, 2].Value2 = Global.numberOfTrips[Global.locations[key], 1];
                //sendMsg.sendMessage(key + "  "+ Global.locations[key]+ "   If working = "+ Global.numberOfTrips[Global.locations[key], 1]);
                cell++;
            }

            summarykSheet.Cells[17, 1].Value2 = "Total Extra Charges for Month";

            summarykSheet.Cells[19, 1].Value2 = "Boxes Picked";
            summarykSheet.Cells[19, 2].Value2 = 0;

            summarykSheet.Cells[20, 1].Value2 = "Cabs Picked";
            summarykSheet.Cells[20, 2].Value2 = 0;

            summarykSheet.Cells[22, 1].Value2 = "Flatpacked Boxes delivered";
            summarykSheet.Cells[22, 2].Value2 = Global.boxesDelivered;

            /////////////////////////// PCCC Storage /////////////////////////

            string[] textBoxValues = openBalanceValue.Split('/');

            Worksheet pcccStorage = newWorkBook.Worksheets.get_Item(5);
            pcccStorage.Name = "PCCC Storage";

            pcccStorage.Columns["A:C"].ColumnWidth = 14;
            pcccStorage.Columns["D:K"].ColumnWidth = 11;
            pcccStorage.Columns["J:J"].ColumnWidth = 14;

            // PCCC Storage Boxes part
            pcccStorage.Cells[1, 1].Value2 = "Boxes";
            pcccStorage.Cells[2, 1].Value2 = "Date";
            pcccStorage.Cells[2, 2].Value2 = "Customer";
            pcccStorage.Cells[2, 3].Value2 = "Opening Bal";

           

            pcccStorage.Cells[2, 4].Value2 = "IN";
            pcccStorage.Cells[2, 5].Value2 = "OUT";
            pcccStorage.Cells[2, 6].Value2 = "Closing Bal";
            pcccStorage.Cells[2, 8].Value2 = "Storage";
            pcccStorage.Cells[2, 10].Value2 = "W/Ending";

            // PCCC Storage Cabinets part
            pcccStorage.Cells[9, 1].Value2 = "Cabinets";
            pcccStorage.Cells[10, 1].Value2 = "Date";
            pcccStorage.Cells[10, 2].Value2 = "Customer";
            pcccStorage.Cells[10, 3].Value2 = "Opening Bal";
 
            

            pcccStorage.Cells[10, 4].Value2 = "IN";
            pcccStorage.Cells[10, 5].Value2 = "OUT";
            pcccStorage.Cells[10, 6].Value2 = "Closing Bal";
            pcccStorage.Cells[10, 8].Value2 = "Storage";
            pcccStorage.Cells[10, 10].Value2 = "W/Ending";

            //sendMsg.sendMessage(textBoxValues.Length.ToString() + "  text box value length");

            if (textBoxValues.Length > 1)
            {
                    pcccStorage.Cells[3, 3].Value2 = textBoxValues[0]; // passed open balance from form
                    pcccStorage.Cells[11, 3].Value2 = textBoxValues[1];
            }
            

            // Add week number and week end date to both Boxes and Cabinets
            for (int i = 0; i < Global.endOfWeek.Length; i++) {
                // PCCC Boxes
                pcccStorage.Cells[3 + i, 1].Value2 = "Wk " + (1+i) + " " + DateTime.Now.AddMonths(-1).ToString("MMMM");
                pcccStorage.Cells[3 + i, 2].Value2 = "PCCC Boxes";
                pcccStorage.Cells[3 + i, 10].Value2 = Global.endOfWeek[i];
                //PCCC Cabinets
                pcccStorage.Cells[11 + i, 1].Value2 = "Wk " + (1 + i) + " " + DateTime.Now.AddMonths(-1).ToString("MMMM");
                pcccStorage.Cells[11 + i, 2].Value2 = "PCCC Cabinets";
                pcccStorage.Cells[11 + i, 10].Value2 = Global.endOfWeek[i];
            }

            string[] timePicked;
            timePicked = Global.pickList.ToArray();

            int counter = 0;
            for (int i = 0; i < timePicked.Length; i++)
            {

                if (i % 2 != 0)
                {

                    for (int y = counter; y < Global.endOfWeek.Length; y++)
                    {

                        if (DateTime.Parse(timePicked[i]) < DateTime.Parse(pcccStorage.Cells[3 + y, 10].Value2))
                        {
                            pcccStorage.Cells[3 + y, 4].Value2 = timePicked[i - 1];

                            counter = y;
                            y = Global.endOfWeek.Length;
                        } 
                    }
                }
            }

            // Boxes box to fill up
            int pcccSummaryBoxes = 0;
            int handlingBoxes = 0;
            for (int i = 0; i < Global.endOfWeek.Length; i++)
            {
                if (pcccStorage.Cells[3 + i, 4].Value2 == null) {

                    pcccStorage.Cells[3 + i, 4].Value2 = 0; // IN
                    pcccStorage.Cells[3 + i, 5].Value2 = 0; // OUT
                    pcccStorage.Cells[3 + i, 6].Value2 = Convert.ToInt32(pcccStorage.Cells[3 + i, 3].Value2) + Convert.ToInt32(pcccStorage.Cells[3 + i, 4].Value2); // Closing Balance
                    pcccStorage.Cells[3 + i, 8].Value2 = pcccStorage.Cells[3 + i, 6].Value2; // Storage
                    if (i < Global.endOfWeek.Length - 1)
                    {
                        pcccStorage.Cells[3 + (i + 1), 3].Value2 = pcccStorage.Cells[3 + i, 8].Value2;
                    }
                }
                else
                {
                    //sendMsg.sendMessage("eto ->" + pcccStorage.Cells[3 + i, 4].Value2.ToString());
                
                    pcccStorage.Cells[3 + i, 5].Value2 = 0; // OUT
                    pcccStorage.Cells[3 + i, 6].Value2 = Convert.ToInt32(pcccStorage.Cells[3 + i, 3].Value2) + Convert.ToInt32(pcccStorage.Cells[3 + i, 4].Value2); // Closing Balance
                    pcccStorage.Cells[3 + i, 8].Value2 = pcccStorage.Cells[3 + i, 6].Value2; // Storage
                    if (i < Global.endOfWeek.Length - 1) {
                        pcccStorage.Cells[3 + (i + 1), 3].Value2 = pcccStorage.Cells[3 + i, 8].Value2;
                    }
                }

                pcccSummaryBoxes += Convert.ToInt32(pcccStorage.Cells[3 + i, 6].Value2);
                handlingBoxes += Convert.ToInt32(pcccStorage.Cells[3 + i, 4].Value2);
            }
    
            int pcccSummaryCabinets = 0;
            // Cabinets to fill up
            for (int i = 0; i < Global.endOfWeek.Length; i++)
            {
                if (pcccStorage.Cells[3 + i, 4].Value2 == null)
                {
                    pcccStorage.Cells[11 + i, 4].Value2 = 0; // IN
                    pcccStorage.Cells[11 + i, 5].Value2 = 0; // OUT
                    pcccStorage.Cells[11 + i, 6].Value2 = Convert.ToInt32(pcccStorage.Cells[3 + i, 3].Value2) + Convert.ToInt32(pcccStorage.Cells[3 + i, 4].Value2); // Closing Balance
                    pcccStorage.Cells[11 + i, 8].Value2 = pcccStorage.Cells[3 + i, 6].Value2; // Storage
                    if (i < Global.endOfWeek.Length - 1)
                    {
                        pcccStorage.Cells[11 + (i + 1), 3].Value2 = pcccStorage.Cells[3 + i, 8].Value2;
                    }
                }
                else
                {
                    //sendMsg.sendMessage("eto ->" + pcccStorage.Cells[11 + i, 4].Value2.ToString());
                    pcccStorage.Cells[11 + i, 5].Value2 = 0; // OUT
                    pcccStorage.Cells[11 + i, 6].Value2 = Convert.ToInt32(pcccStorage.Cells[11 + i, 3].Value2) + Convert.ToInt32(pcccStorage.Cells[11 + i, 4].Value2); // Closing Balance
                    pcccStorage.Cells[11 + i, 8].Value2 = pcccStorage.Cells[11 + i, 6].Value2; // Storage
                    if (i < Global.endOfWeek.Length - 1)
                    {
                        pcccStorage.Cells[11 + (i + 1), 3].Value2 = pcccStorage.Cells[11 + i, 8].Value2;
                    }
                    pcccSummaryCabinets += Convert.ToInt32(pcccStorage.Cells[11 + i, 6].Value2);
                }

            }

            pcccStorage.Cells[2 + Global.endOfWeek.Length, 8].Value2 = Global.pcccBoxes; // current pccc box count
            pcccStorage.Cells[10 + Global.endOfWeek.Length, 8].Value2 = Global.pcccCabinets; // current pccc cabinet count

            /////////////////// Rates //////////////////////////
            Workbook excelRates = excelApp.Workbooks.Open(@"V:\Warehouses\Parkmore Warehouse\Reports\HSE_RATES.xlsx");
            _Worksheet excelRatesSheet = excelRates.Sheets[1];

            // Get new workBook sheet nr 6
            Worksheet pcccSummary = newWorkBook.Worksheets.get_Item(6);
            // Copy rates sheet into new Nr 6 sheet
            excelRatesSheet.Copy(pcccSummary);
            // close Rates workBook
            excelRates.Save();
            excelRates.Close(true);
            // Rename NewBook Nr 6 sheet
            pcccSummary = newWorkBook.Worksheets.get_Item("PCCC");
            pcccSummary.Name = "PCCC Summary " + DateTime.Now.AddMonths(-1).ToString("MMMM") + " " + thisDay.Year;

            pcccSummary.Cells[1, 1].Value2 = "PCCC Summary Invoice " + DateTime.Now.AddMonths(-1).ToString("MMMM") + " " + thisDay.Year;
            pcccSummary.Cells[21, 1].Value2 = "Total Invoice " + DateTime.Now.AddMonths(-1).ToString("MMMM") + " " + thisDay.Year;

            pcccSummary.Cells[4, 3].Value2 = pcccSummaryBoxes;
            pcccSummary.Cells[5, 3].Value2 = pcccSummaryCabinets;

            //locations 
            pcccSummary.Cells[9, 3].Value2 = Global.numberOfTrips[Global.locations["Shantalla"], 1] + Global.numberOfTrips[Global.locations["Doughiska"], 1] + Global.numberOfTrips[Global.locations["Mervue"], 1];
            //Global.numberOfTrips[Global.locations["Athenry"], 1] + Global.numberOfTrips[Global.locations["Tuam"], 1] + Global.numberOfTrips[Global.locations["Loughrea"], 1] + Global.numberOfTrips[Global.locations["Mountbellew"], 1] + Global.numberOfTrips[Global.locations["Ballinasloe"], 1];
            pcccSummary.Cells[10, 3].Value2 = Global.numberOfTrips[Global.locations["Athenry"], 1] + Global.numberOfTrips[Global.locations["Tuam"], 1] + Global.numberOfTrips[Global.locations["Loughrea"], 1] + Global.numberOfTrips[Global.locations["Mountbellew"], 1] + Global.numberOfTrips[Global.locations["Ballinasloe"], 1]; ;

            pcccSummary.Cells[14, 3].Value2 = Global.filesPicked;
            pcccSummary.Cells[15, 3].Value2 = handlingBoxes;

            /////////////////// END //////////////////////////

            newWorkBook.SaveAs(@"C:\Users\ssladmin\Desktop\Weekly rep\PCCC Monthly Invoice.xlsx");
            newWorkBook.Close(true);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
