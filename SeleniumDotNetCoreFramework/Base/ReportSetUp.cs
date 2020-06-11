using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using OpenQA.Selenium;
using SeleniumDotNetCoreFramework.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace SeleniumDotNetCoreFramework.Base
{
    public class ReportSetUp
    {
        private static ReportSetUp reportSetUpInstance = null;
        public static string htmlreportPath;



        public ReportSetUp()
        {

        }
        public static ReportSetUp getInstance()
        {

            if (reportSetUpInstance == null)
            {

                reportSetUpInstance = new ReportSetUp();


            }
            return reportSetUpInstance;

        }



        public AventStack.ExtentReports.ExtentReports getReport()
        {
            string reportPath = ReportingHelpers.reportpath();
            string timeNow = DateTime.Now.ToLongTimeString().ToString().Replace(':', '_');
            //initialize the reportsetup
            htmlreportPath = reportPath + @"\" + "RunResult_" + timeNow + ".html";

            Logger.log("HTML Report Path" + htmlreportPath);
            var htmlReporter = new ExtentV3HtmlReporter(htmlreportPath);

            var extent = new AventStack.ExtentReports.ExtentReports();
            extent.AttachReporter(htmlReporter);
            // adding environment variables

            string machineName = Environment.MachineName;
            string systemOS = Environment.OSVersion.ToString();
            extent.AddSystemInfo("Machine Name", machineName);

            extent.AddSystemInfo("Operating System", systemOS);

            return extent;



        }
    }
}
