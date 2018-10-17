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
    class DataMergingForExpressWay
    {

        List<string> routeNameList = new List<string>();
        List<string> bothRouteNameList = new List<string>();

        List<string> postNameList = new List<string>();
        List<string> bothPostNameList = new List<string>();

        List<string> downDataList = new List<string>();
        List<string> upDataList = new List<string>();
        List<string> upAndDownDataList = new List<string>();

        List<string> latitudeDataList = new List<string>();
        List<string> bothLatitudeDataList = new List<string>();

        List<string> longitudeDataList = new List<string>();
        List<string> bothLongitudeDataList = new List<string>();

        Form1 fObj = new Form1();

        public Tuple<string, string> ProcessingForMergingDataForExpressway(string filePathWithName)
        {
            string[] filePathArr = filePathWithName.Split(new char[] { '\\' });
            string fileName = filePathArr[(filePathArr.Length - 1)];
            string folderPath = filePathWithName.Remove(filePathWithName.Length - fileName.Length);

            var excel = new ExcelQueryFactory(filePathWithName);

            var worksheetNames = excel.GetWorksheetNames();
            // To find the list of sheets
            List<string> sheetNameList = new List<string>();
            foreach (var a in worksheetNames)
            {
                sheetNameList.Add(a);
            }

            var rowWiseData = from r in excel.Worksheet(sheetNameList[0])
                              select new { r };

            foreach (var data in rowWiseData)
            {
                //***** Adding Route Name data to list******************
                //***** Adding Post data  to list******************
                //***** Adding latitude  data to list******************
                //***** Adding longitude  data to list******************
                if (!string.IsNullOrEmpty(data.r[5]) && !string.IsNullOrEmpty(data.r[36]) && !string.IsNullOrEmpty(data.r[37]) && !string.IsNullOrEmpty(data.r[38]))
                {
                    routeNameList.Add(data.r[5].ToString());
                    bothRouteNameList.Add(data.r[5].ToString());

                    postNameList.Add(data.r[36].ToString());
                    bothPostNameList.Add(data.r[36].ToString());

                    downDataList.Add("D");

                    latitudeDataList.Add(data.r[37].ToString());
                    bothLatitudeDataList.Add(data.r[37].ToString());

                    longitudeDataList.Add(data.r[38].ToString());
                    bothLongitudeDataList.Add(data.r[38].ToString());
                }
                else
                {

                    //Do Nothing
                }
            }

            foreach (var post in downDataList)
            {
                upDataList.Add("U");
            }

            routeNameList.Reverse();
            bothRouteNameList.AddRange(routeNameList);

            postNameList.Reverse();
            bothPostNameList.AddRange(postNameList);

            upAndDownDataList.AddRange(downDataList);
            upAndDownDataList.AddRange(upDataList);

            latitudeDataList.Reverse();
            bothLatitudeDataList.AddRange(latitudeDataList);

            longitudeDataList.Reverse();
            bothLongitudeDataList.AddRange(longitudeDataList);

            //*********************************************************************Reading and creation of data ends and writing starts******************************

            // Load Excel application
            Microsoft.Office.Interop.Excel.Application excelWrite = new Microsoft.Office.Interop.Excel.Application();
            excelWrite.DisplayAlerts = false;
            // Create empty workbook
            excelWrite.Workbooks.Add();


            // Create Worksheet from active sheet
            Microsoft.Office.Interop.Excel._Worksheet workSheet = excelWrite.ActiveSheet;
            workSheet.Name = "Result";

            ((Range)workSheet.Cells[1, 1]).EntireColumn.ColumnWidth = 19;
            ((Range)workSheet.Cells[1, 1]).Font.Size = 12;
            ((Range)workSheet.Cells[1, 1]).Font.Bold = true;
            ((Range)workSheet.Cells[1, 1]).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            ((Range)workSheet.Cells[1, 2]).EntireColumn.ColumnWidth = 19;
            ((Range)workSheet.Cells[1, 2]).EntireColumn.NumberFormat = "0.000";
            ((Range)workSheet.Cells[1, 2]).Font.Size = 12;
            ((Range)workSheet.Cells[1, 2]).Font.Bold = true;
            ((Range)workSheet.Cells[1, 3]).EntireColumn.ColumnWidth = 19;
            ((Range)workSheet.Cells[1, 3]).Font.Size = 12;
            ((Range)workSheet.Cells[1, 3]).Font.Bold = true;
            ((Range)workSheet.Cells[1, 4]).EntireColumn.ColumnWidth = 19;
            ((Range)workSheet.Cells[1, 4]).EntireColumn.NumberFormat = "0.000000000000";
            ((Range)workSheet.Cells[1, 4]).Font.Size = 12;
            ((Range)workSheet.Cells[1, 4]).Font.Bold = true;
            ((Range)workSheet.Cells[1, 5]).EntireColumn.ColumnWidth = 19;
            ((Range)workSheet.Cells[1, 5]).EntireColumn.NumberFormat = "0.000000000000";
            ((Range)workSheet.Cells[1, 5]).Font.Size = 12;
            ((Range)workSheet.Cells[1, 5]).Font.Bold = true;

            workSheet.Cells[1, "A"] = "路線名";
            workSheet.Cells[1, "B"] = "キロポスト";
            workSheet.Cells[1, "C"] = "上下線";
            workSheet.Cells[1, "D"] = "緯度";
            workSheet.Cells[1, "E"] = "経度 ";

            for (int i = 0; i < bothRouteNameList.Count; i++)
            {
                workSheet.Cells[i + 2, "A"] = bothRouteNameList[i].ToString();
                workSheet.Cells[i + 2, "B"] = bothPostNameList[i].ToString();
                workSheet.Cells[i + 2, "C"] = upAndDownDataList[i].ToString();
                workSheet.Cells[i + 2, "D"] = bothLatitudeDataList[i].ToString();
                workSheet.Cells[i + 2, "E"] = bothLongitudeDataList[i].ToString();
            }

            try
            {
                excelWrite.DisplayAlerts = false;
                workSheet.SaveAs(folderPath + "Result_ForExpressway_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");

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

            finally
            {
                ClearAllEndAll(excelWrite, workSheet);
            }
            return Tuple.Create(folderPath, folderPath + "Result_ForExpressway_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            //return <folderPath, (folderPath + "Result_ForExpressway_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv")>;
        }

        private void ClearAllEndAll(Microsoft.Office.Interop.Excel.Application excelObj, Microsoft.Office.Interop.Excel._Worksheet sheetObj)
        {
            // Quit Excel application
            excelObj.Quit();


            // Release COM objects (very important!)
            if (excelObj != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelObj);


            if (sheetObj != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheetObj);


            // Empty variables
            excelObj = null;
            sheetObj = null;


            // Force garbage collector cleaning
            GC.Collect();

        }


    }
}
