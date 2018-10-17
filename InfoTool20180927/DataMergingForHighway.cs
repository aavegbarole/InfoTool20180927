using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using LinqToExcel;
using LinqToExcel.Attributes;
using Microsoft.Office.Interop.Excel;
using System.Threading;

namespace InfoTool20180927
{
    class DataMergingForHighway
    {
        List<string> routeNumberList = new List<string>();
        List<string> routeNameList = new List<string>();
        List<string> typeOfDataList = new List<string>();

        List<string> latitudeDataList = new List<string>();
        List<string> longitudeDataList = new List<string>();

        Form1 fObj = new Form1();

        public Tuple<string, string> ProcessingForMergingDataForHighway(string filePathWithName)
        {
            string[] filePathArr = filePathWithName.Split(new char[] { '\\' });
            string fileName = filePathArr[(filePathArr.Length - 1)];
            string folderPath = filePathWithName.Remove(filePathWithName.Length - fileName.Length);

            // Excel Object to read excel file
            var excel = new ExcelQueryFactory(filePathWithName);
            var worksheetNames = excel.GetWorksheetNames();
            // To find the list of sheets
            List<string> sheetNameList = new List<string>();
            foreach (var a in worksheetNames)
            {
                if (a.Any(char.IsDigit) && a.Contains("_"))
                {
                    sheetNameList.Add(a);
                }
            }

            //*******************************************************************************Actual Code Starts ****************************************************************    
            int endCounterBetweenDownAndUp = 0;
            int endCounterEndOfSheet = 0;
            for (int sheetNo = 0; sheetNo < sheetNameList.Count; sheetNo++)
            {
                char[] spaceSeparator = new char[] { '_' };
                string[] nameOfSheetWithRouteNumberAndRouteName = sheetNameList[sheetNo].ToString().Split(spaceSeparator, StringSplitOptions.None);
                string routeNumber = nameOfSheetWithRouteNumberAndRouteName[0];
                string routeName = nameOfSheetWithRouteNumberAndRouteName[1];
                var rowWiseData = from c in excel.Worksheet(sheetNameList[sheetNo])
                                  select new { c };
                foreach (var data in rowWiseData)
                {
                    //***** Adding Latitude data to list******************
                    //***** Adding type of data to list ***********************
                    if (!string.IsNullOrEmpty(data.c[2]) && !string.IsNullOrEmpty(data.c[3]))
                    {
                        latitudeDataList.Add(data.c[2]);
                        typeOfDataList.Add("D");
                        routeNumberList.Add(routeNumber);
                        routeNameList.Add(routeName);
                        longitudeDataList.Add(data.c[3]);
                    }
                    else
                    {
                        //Do Nothing

                    }
                }

                try
                {
                    //Data of sheet removed
                    latitudeDataList.RemoveAt(endCounterEndOfSheet);
                    longitudeDataList.RemoveAt(endCounterEndOfSheet);
                    typeOfDataList.RemoveAt(endCounterEndOfSheet);
                    routeNumberList.RemoveAt(endCounterEndOfSheet);
                    routeNameList.RemoveAt(endCounterEndOfSheet);
                    endCounterBetweenDownAndUp = latitudeDataList.Count;
                }
                catch (Exception)
                {
                    var result = MessageBox.Show("選択した入力ファイルの書式を確認してください。", "エラーメッセージ", MessageBoxButtons.RetryCancel);
                    switch (result)
                    {
                        case DialogResult.Retry:   // Retry button pressed
                            //System.Windows.Forms.Application.Restart();
                            System.Diagnostics.Process.Start(System.Windows.Forms.Application.ExecutablePath); // to start new instance of application
                            fObj.Close();  //to close the current instance
                            break;
                        case DialogResult.Cancel:    // Cancel button pressed
                            excel.Dispose();
                            GC.SuppressFinalize(this);
                            Environment.Exit(1);
                            break;
                        default:                 // Neither Retry nor Cancel pressed (just in case)
                            MessageBox.Show("もう一度お試しください");
                            break;
                    }
                }

                foreach (var data in rowWiseData)
                {
                    //***** Adding Latitude down data to list******************
                    //***** Adding type of data to list ***********************
                    if (!string.IsNullOrEmpty(data.c[7]) && !string.IsNullOrEmpty(data.c[8]))
                    {
                        latitudeDataList.Add(data.c[7]);
                        typeOfDataList.Add("U");
                        routeNumberList.Add(routeNumber);
                        routeNameList.Add(routeName);
                        longitudeDataList.Add(data.c[8]);
                    }
                    else
                    {
                        //Do Nothing
                    }
                }
                try
                {
                    //Data of sheet removed
                    latitudeDataList.RemoveAt(endCounterBetweenDownAndUp);
                    longitudeDataList.RemoveAt(endCounterBetweenDownAndUp);
                    typeOfDataList.RemoveAt(endCounterBetweenDownAndUp);
                    routeNumberList.RemoveAt(endCounterBetweenDownAndUp);
                    routeNameList.RemoveAt(endCounterBetweenDownAndUp);
                    endCounterEndOfSheet = longitudeDataList.Count;
                }
                catch (Exception)
                {
                    var result = MessageBox.Show("選択した入力ファイルの書式を確認してください。", "エラーメッセージ", MessageBoxButtons.RetryCancel);
                    switch (result)
                    {
                        case DialogResult.Retry:   // Retry button pressed
                            System.Diagnostics.Process.Start(System.Windows.Forms.Application.ExecutablePath); // to start new instance of application
                            fObj.Close();
                            break;
                        case DialogResult.Cancel:    // Cancel button pressed
                            excel.Dispose();
                            GC.SuppressFinalize(this);
                            Environment.Exit(1);
                            break;
                        default:                 // Neither Retry nor Cancel pressed (just in case)
                            MessageBox.Show("もう一度お試しください");
                            break;
                    }
                }
            }
            //*******************************************************************************Actual Code ends ****************************************************************   
            // Load Excel application
            Microsoft.Office.Interop.Excel.Application excelWrite = new Microsoft.Office.Interop.Excel.Application();
            // Create empty workbook
            excelWrite.Workbooks.Add();
            // Create Worksheet from active sheet
            Microsoft.Office.Interop.Excel._Worksheet workSheet = excelWrite.ActiveSheet;
            workSheet.Name = "Result";
            ((Range)workSheet.Cells[1, 1]).EntireColumn.ColumnWidth = 19;
            ((Range)workSheet.Cells[1, 1]).Font.Size = 12;
            ((Range)workSheet.Cells[1, 1]).Font.Bold = true;
            ((Range)workSheet.Cells[1, 2]).EntireColumn.ColumnWidth = 19;
            ((Range)workSheet.Cells[1, 2]).Font.Size = 12;
            ((Range)workSheet.Cells[1, 2]).Font.Bold = true;
            ((Range)workSheet.Cells[1, 3]).EntireColumn.ColumnWidth = 19;
            ((Range)workSheet.Cells[1, 3]).Font.Size = 12;
            ((Range)workSheet.Cells[1, 3]).Font.Bold = true;
            ((Range)workSheet.Cells[1, 4]).EntireColumn.ColumnWidth = 19;
            ((Range)workSheet.Cells[1, 4]).Font.Size = 12;
            ((Range)workSheet.Cells[1, 4]).Font.Bold = true;
            ((Range)workSheet.Cells[1, 4]).EntireColumn.NumberFormat = "0.00000000";
            ((Range)workSheet.Cells[1, 5]).EntireColumn.ColumnWidth = 19;
            ((Range)workSheet.Cells[1, 5]).Font.Size = 12;
            ((Range)workSheet.Cells[1, 5]).Font.Bold = true;
            ((Range)workSheet.Cells[1, 5]).EntireColumn.NumberFormat = "0.00000000";
            workSheet.Cells[1, "A"] = "路線番号";
            workSheet.Cells[1, "B"] = "路線名";
            workSheet.Cells[1, "C"] = "上下線";
            workSheet.Cells[1, "D"] = "緯度";
            workSheet.Cells[1, "E"] = "経度 ";
            for (int i = 0; i < routeNumberList.Count; i++)
            {
                workSheet.Cells[i + 2, "A"] = routeNumberList[i].ToString();
                workSheet.Cells[i + 2, "B"] = routeNameList[i].ToString();
                workSheet.Cells[i + 2, "C"] = typeOfDataList[i].ToString();
                workSheet.Cells[i + 2, "D"] = latitudeDataList[i].ToString();
                workSheet.Cells[i + 2, "E"] = longitudeDataList[i].ToString();
            }
            try
            {
                excelWrite.DisplayAlerts = false;
                workSheet.SaveAs(folderPath + "Result_ForHighway_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }

            catch (Exception)
            {
                var result = MessageBox.Show("Result_Highway.csvが既にシステム上で開いているかどうか確認してください。", "エラーメッセージ", MessageBoxButtons.RetryCancel);

                switch (result)
                {
                    case DialogResult.Retry:   // Retry button pressed
                        System.Diagnostics.Process.Start(System.Windows.Forms.Application.ExecutablePath); // to start new instance of application
                        fObj.Close();
                        break;
                    case DialogResult.Cancel:    // Cancel button pressed
                        excel.Dispose();
                        GC.SuppressFinalize(this);
                        Environment.Exit(1);
                        System.Windows.Forms.Application.Exit();
                        break;
                    default:                 // Neither Retry nor Cancel pressed (just in case)
                        MessageBox.Show("もう一度お試しください");
                        break;
                }
            }

            finally
            {
                ClearAllEndAll(excelWrite, workSheet);
            }
            return Tuple.Create(folderPath, folderPath + "Result_ForHighway_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
        }




        private void ClearAllEndAll(Microsoft.Office.Interop.Excel.Application excelObj, Microsoft.Office.Interop.Excel._Worksheet sheetObj)
        {
            // Quit Excel application
            excelObj.Quit();
            // Release COM objects (very important!)
            if (excelObj != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelObj);
            }
            if (sheetObj != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheetObj);
            }
            // Empty variables
            excelObj = null;
            sheetObj = null;
            // Force garbage collector cleaning
            GC.Collect();
        }
    }
}


