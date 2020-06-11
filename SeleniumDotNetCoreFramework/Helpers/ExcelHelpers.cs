using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using SeleniumDotNetCoreFramework.Base;
using excel = Microsoft.Office.Interop.Excel;

namespace SeleniumDotNetCoreFramework.Helpers
{
    public class ExcelHelpers
    {
        public static Microsoft.Office.Interop.Excel.Application xlApp;

        public static Microsoft.Office.Interop.Excel.Workbook xlWorkBook;

        public static Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
        public static Hashtable dataHash;
        public static string workBookName;



        public static void openExcel(string SheetName)
        {
            try
            {
                string workBookName = getScenarioName();
                xlApp = new Microsoft.Office.Interop.Excel.Application();

                //open workbook
                string sltnPath = DriverContext.getSolutionPath();
                string testDataPath = sltnPath + @"TestData\" + workBookName + ".xlsx";
                Thread.Sleep(2000);
                xlWorkBook = xlApp.Workbooks.Open(testDataPath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                //Access worksheet

                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets[SheetName];
            }
            catch (Exception e)
            {
                Logger.log("error::" + e.Message);
            }
        }
        public static void openExcelRunConfig(string workBookName, string SheetName)
        {
            try
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();

                //open workbook
                string sltnPath = DriverContext.getSolutionPath();
                string testDataPath = sltnPath + @"TestData\" + workBookName + ".xlsx";
                xlWorkBook = xlApp.Workbooks.Open(testDataPath, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                //Access worksheet

                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets[SheetName];
            }
            catch (Exception e)
            {
                Logger.log("error::" + e.Message);
            }
        }
        public static string readDatafromExcel(string SheetName, string columnValue)
        {

            openExcel(SheetName);
            string testCaseName = DriverContext.getTestCaseName();
            int rowNumber = xlWorkSheet.Columns.Find(testCaseName).Cells.Row;
            excel.Range range = xlWorkSheet.UsedRange;
            int columnNumber = xlWorkSheet.Columns.Find(columnValue).Cells.Column;
            //string testDataValue = xlWorkSheet.UsedRange[rowNumber, columnNumber].Value2.ToString();
            string testDataValue = Convert.ToString((range.Cells[rowNumber, columnNumber] as excel.Range).Value2);
            closeExcel();
            return testDataValue;



        }

        public static string readDatafromExcel(string SheetName, int rowValue, string columnValue)
        {

            openExcel(SheetName);
            string testCaseName = DriverContext.getTestCaseName();
            int rowNumber = xlWorkSheet.Columns.Find(rowValue).Cells.Row;

            int columnNumber = xlWorkSheet.Columns.Find(columnValue).Cells.Column;

            string testDataValue = xlWorkSheet.Range[rowNumber, columnNumber].Value2;
            closeExcel();
            return testDataValue;

        }
        /* for writing we need to make in open excel method parameter as false
        save(for writing)
        WriteData Method base don the Test Case Name
        */
        //
        public static void writeDataToExcel(string SheetName, string columnValue, string dataValue)
        {
            openExcel(SheetName);
            string testCaseName = DriverContext.getTestCaseName();

            int rowNumber = xlWorkSheet.Columns.Find(testCaseName).Cells.Row;

            int columnNumber = xlWorkSheet.Columns.Find(columnValue).Cells.Column;

            xlWorkSheet.Cells[rowNumber, columnNumber] = dataValue;

            closeExcel();


        }

        public static void storeDataHash()
        {
            try
            {
                openExcelRunConfig("RunConfigurationManager", "TestScripts");
                string testcaseID = DriverContext.getTestCaseName();

                dataHash = new Hashtable();
                int rowNumber = xlWorkSheet.Columns.Find(testcaseID).Cells.Row;
                int columnCount = xlWorkSheet.UsedRange.Columns.Count;
                for (int i = 1; i <= columnCount; i++)
                {

                    var testKey = xlWorkSheet.Range[HookBase.getColumnNameFromIndex(i) + 1].Value2;
                    var testValue = xlWorkSheet.Range[HookBase.getColumnNameFromIndex(i) + rowNumber].Value2;
                    if (testKey != null)
                    {
                        dataHash.Add(testKey, testValue);

                    }
                }
            }
            catch (Exception e)
            {
                Logger.log("error" + e.Message);
            }
            finally
            {
                closeExcel();
            }
        }


        public static string getData(string SheetName, string ColumnVariable)
        {
            string newSheetName = "";

            string env = ExcelHelpers.getParameter("Environment");
            if (env.Equals("UAT"))
            {
                newSheetName = SheetName;
            }
            else if (env.Equals("QA"))
            {
                newSheetName = SheetName + "_QA";
            }
            else if (env.Equals("PreProd"))
            {
                newSheetName = SheetName + "_PreProd";
            }
            else
            {
                newSheetName = SheetName;
            }
            string variable = readDatafromExcel(newSheetName, ColumnVariable);

            return variable;
        }

        public static void setData(string SheetName, string columnValue, string dataValue)
        {
            string newSheetName = "";
            string env = ExcelHelpers.getParameter("Environment");
            if (env.Equals("UAT"))
            {
                newSheetName = SheetName;
            }
            else if (env.Equals("QA"))
            {
                newSheetName = SheetName + "_QA";
            }
            else if (env.Equals("PreProd"))
            {
                newSheetName = SheetName + "_PreProd";
            }
            else
            {
                newSheetName = SheetName;
            }

            writeDataToExcel(newSheetName, columnValue, dataValue);

        }
        public static void closeExcel()
        {
            try
            {
                xlWorkBook.Save();
                xlWorkBook.Close();
                xlApp.Quit();
            }
            catch
            {


            }
        }


        public static string getScenarioName()
        {
            string scenario = "";

            scenario = (string)dataHash["Test Scenario"];

            return scenario;
        }

        public static string getParameter(string Key)
        {
            string parameter = (string)dataHash[Key];

            return parameter;
        }

        //Write into the Excel Data Sheet 

    }

}

