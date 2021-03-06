﻿using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using SeleniumDotNetCoreFramework.Helpers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumDotNetCoreFramework.Helpers
{
    public class ReportingHelpers
    {
        public static string timeStamp = SeleniumDotNetCoreFramework.Helpers.Reporting.timeStamp();
        public static string reportPath = "";
        public static string reportZipath = "";
        public static string DBPath = "";

        public static string reportpath()
        {
            string asmblyPath = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actPath = asmblyPath.Substring(0, asmblyPath.LastIndexOf("bin"));
            string sltnPath = new Uri(actPath).LocalPath;
            reportPath = sltnPath + @"TestResults\";
            string sub_directorypath = reportPath + timeStamp;
            if (!Directory.Exists(sub_directorypath))
            {
                Directory.CreateDirectory(sub_directorypath);
            }
            return sub_directorypath;
        }

        public static string reportZipPath()
        {

            string asmblyPath = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actPath = asmblyPath.Substring(0, asmblyPath.LastIndexOf("bin"));
            string sltnPath = new Uri(actPath).LocalPath;
            reportZipath = sltnPath + @"EmailReports\" + timeStamp;


            return reportZipath;

        }

        public static string SQLiteDBPath(string DBFilePath)
        {
            string asmblyPath = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string actPath = asmblyPath.Substring(0, asmblyPath.LastIndexOf("bin"));
            string sltnPath = new Uri(actPath).LocalPath;
            DBPath = sltnPath + DBFilePath;
            return DBPath;
        }

    }
}
