using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SQLite;
using SeleniumDotNetCoreFramework.Helpers;
using System.Data;
using Microsoft.Data.Sqlite;
using SeleniumDotNetCoreFramework.Config;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;


namespace SeleniumDotNetCoreFramework.Helpers
{
    public class SQLiteDBHelpers
    {
        public static SQLiteConnection sqliteConn;
        public static SQLiteCommand cmd = null;
        public static DataTable dt;
        public static string DBPath;
        public static string DatabaseFile = "";
        public static SQLiteConnection DBConnect()
        {
            DatabaseFile = ReportingHelpers.SQLiteDBPath(Settings.AppConnectionString);

            string DataSource = string.Format("Data Source={0}", DatabaseFile);


            try
            {
                sqliteConn = new SQLiteConnection(DataSource);
                sqliteConn.Open();
                return sqliteConn;

            }

            catch (Exception e)
            {
                Logger.log("ERROR::" + e.Message);

            }
            return null;
        }
        public static void DBClose()
        {
            try
            {
                sqliteConn.Close();
            }
            catch (Exception e)
            {
                Logger.log("ERROR::" + e.Message);

            }

        }
        public static DataTable ExecuteQuery(string queryString)
        {
            DataSet dataset;
            DBConnect();
            try
            {  //Checking the state of the connection
                if (sqliteConn == null || ((sqliteConn != null && (sqliteConn.State == ConnectionState.Closed ||
                  sqliteConn.State == ConnectionState.Broken))))
                    sqliteConn.Open();
                SQLiteDataAdapter Adaptor = new SQLiteDataAdapter();
                Adaptor.SelectCommand = new SQLiteCommand(queryString, sqliteConn);
                Adaptor.SelectCommand.CommandType = CommandType.Text;

                dataset = new DataSet();
                Adaptor.Fill(dataset, "table");
                sqliteConn.Close();
                return dataset.Tables["table"];
            }
            catch (Exception e)
            {
                dt = null;
                sqliteConn.Close();
                Logger.log("ERROR::" + e.Message);
                return null;
            }
            finally
            {
                sqliteConn.Close();
                dt = null;
            }

        }

    }
}
