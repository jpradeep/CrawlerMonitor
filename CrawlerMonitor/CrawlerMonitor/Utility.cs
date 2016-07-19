using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Functions;
using System.Data;
using System.Configuration;
using System.IO;
using Functions;
using System.Diagnostics;
using System.Globalization;

namespace CrawlerMonitor
{
    class Utility
    {
        public static GenericDatabase database = null;

        #region GetDBConnection
        /// <summary>
        /// GetDBConnection
        /// </summary>
        /// <returns></returns>
        private static GenericDatabase GetDBConnection()
        {
            GenericDatabase database = null;
            try
            {
                database = new GenericDatabase(ConfigurationManager.AppSettings.Get(Constants.CONNECTION_STRING), ConfigurationManager.AppSettings.Get(Constants.DATABASE_NAME), GenericDatabase.DatabaseType.SqlServer);
            }
            catch (Exception Ex)
            {
                //Utility.SendErrorMail("", Ex.Message);
                Logger.LogMessage(Ex.Message, LogMessageLevel.B_Error);
            }
            return database;
        }
        #endregion

        #region CrawlerMonitoring
        /// <summary>
        /// Crawler Monitoring Daily
        /// </summary>
        /// <returns></returns>
        public static bool CrawlerMonitoring(string inp)
        {

            string localFileName = FileOperation.GetRootPath("TEMP");
            string xmlFileName = string.Empty;
            List<string> files = new List<string>();
            DataTable resultTable = new DataTable();
            DataSet dsDaily = new DataSet();
            dsDaily = GenerateReport(inp);

            localFileName = FileOperation.GetRootPath("TEMP");
            xmlFileName = "CrawlerMonitoring_" + inp + ".csv";
            localFileName = Text.TextBeforeLastTag(localFileName, "/").Replace("/", "\\") + "\\" + xmlFileName;
            files.Add(localFileName);

            if (dsDaily.Tables.Count > 0)
                GenerateCSVFile(dsDaily, files, "CrawlerMonitoringDaily");
            else
                Logger.LogMessage("No Records Found ", LogMessageLevel.B_Error);

            sendEmail(inp, files);

            return false;
        }
        #endregion

        #region getStatus
        /// <summary>
        /// getStatus
        /// </summary>
        /// <param name="status"></param>
        /// <returns></returns>
        public static bool getStatus(string status)
        {
            string qry = string.Empty;
            string today = DateTime.Now.ToShortDateString();
            string currTime = DateTime.Now.ToShortTimeString();
            string siteName = string.Empty;
            database = Utility.GetDBConnection();

            siteName = ConfigurationManager.AppSettings.Get("DAILY_STATUS");
            List<string> siteNameList = new List<string>(siteName.Split(','));
            DataTable crawlerStatus = new DataTable();
            string dbName = string.Empty;
            crawlerStatus.Columns.Add("Site Name");
            crawlerStatus.Columns.Add("Completed Count");
            crawlerStatus.Columns.Add("Running Count");
            crawlerStatus.Columns.Add("Pending Count");
            crawlerStatus.Columns.Add("Status");

            foreach (string sitename in siteNameList)
            {
                qry = string.Empty;
                //Empty sitename check 
                if (sitename == string.Empty)
                    continue;
                // Read the DB and Table name from Config file
                dbName = ConfigurationManager.AppSettings.Get(sitename.ToUpper().Trim());
                try
                {
                    //Verify the crawler launched status
                    qry = "SELECT top 1 time FROM [configsetting].[dbo].[ProgramSchedule](nolock) where sitename = '" + sitename.ToUpper().Trim() + "' order by time asc";
                    string schedulerTime = database.GetScalar(qry).ToString();

                    if (Convert.ToDateTime(currTime) < Convert.ToDateTime(schedulerTime))
                    {
                        crawlerStatus.Rows.Add(sitename.ToUpper().Trim(), "-", "-", "-", "Not Started");
                        continue;
                    }
                    if (!string.IsNullOrEmpty(dbName))
                    {
                        DataTable dt = new DataTable();
                        if (sitename.ToUpper().Trim() == "FBO_DOCDOWNLOAD")
                        {
                            string dbName_FBO_5 = "[WebCrawlerOutput_FBODOC_FBO5].[dbo].[FBO_DocDownload_Dispatcher]";
                            string sitename_FBO_5 = "FBO_5_DocDownload";
                            dt = statusValidations(dbName_FBO_5, sitename_FBO_5);
                            crawlerStatus.Merge(dt);
                        }

                        dt = statusValidations(dbName, sitename);
                        crawlerStatus.Merge(dt);
                        continue;
                    }
                    else
                    {
                        qry = string.Empty;
                        if (sitename.Contains("Fed_USASpend_4_RawCSV"))
                            qry = "SELECT COUNT(*) FROM [Utility].[dbo].[PopXMLSQSMessages] (nolock) where sourcename = 'Fed_USASpending_4' and cast(CreatedDate as datetime) >= '" + today + "'";
                        else
                            qry = "SELECT COUNT(*) FROM [Utility].[dbo].[PopXMLSQSMessages] (nolock) where sourcename = '" + sitename.ToUpper().Trim() + "' and cast(CreatedDate as datetime) >= '" + today + "'";

                        int xmlCnt = Convert.ToInt32(database.GetScalarValue(qry));
                        if (xmlCnt > 0)
                            crawlerStatus.Rows.Add(sitename.ToUpper().Trim(), "-", "-", "-", "Completed");
                        else
                            crawlerStatus.Rows.Add(sitename.ToUpper().Trim(), "-", "-", "-", "Not Completed");
                        continue;
                    }
                }
                catch (Exception e)
                {
                    Logger.LogMessage("STATUS Generation failed. Because of " + e.Message.ToString(), LogMessageLevel.B_Error);
                }

            }
            sendStatusMail(crawlerStatus, status, "<h4>Crawler Status Report for (TOP) Priority Sources</h4>", "Crawler Status Report");
            return true;
        }
        #endregion

        #region sendStatusMail
        /// <summary>
        /// sendStatusMail
        /// </summary>
        /// <param name="dt"></param>
        public static void sendStatusMail(DataTable dt, string input, string heading, string subject)
        {
            string message = string.Empty;
            message += heading;
            message += "<table  style='border-collapse:collapse;font-family:verdana,helvetica,arial,sans-serif;border:1px solid #c3c3c3'>";
            message += "<tr>";
            try
            {
                foreach (DataColumn column in dt.Columns)
                {
                    message += "<th style='white-space:nowrap;background-color:#555555;color:#fff;border:1px solid #c3c3c3;padding:3px;vertical-align:top;text-align:center;font-size:11px'>";
                    message += column.ColumnName.ToUpper();
                }
                message += "</tr>";
                int rowCount = 0;

                foreach (DataRow rows in dt.Rows)
                {
                    message += "<tr>";
                    rowCount++;
                    for (int i = 0; i < dt.Columns.Count; ++i)
                    {
                        if (rowCount % 2 == 0)
                        {
                            message += "<td style='vertical-align:top;text-align:left;font-size:11px;background-color:#DDDDDD;padding:3px;border:1px solid #c3c3c3'>";
                        }
                        else
                        {
                            message += "<td style='vertical-align:top;text-align:left;font-size:11px;background-color:#FFB581;padding:3px;border:1px solid #c3c3c3'>";
                        }
                        message += rows[i].ToString();
                    }
                    message += "</tr>";
                }
                message += "</table><br>";
                string emailSubject = subject + " - " + DateTime.Now.ToString();
                String SMTP = ConfigurationManager.AppSettings.Get("SMTP");
                String From = ConfigurationManager.AppSettings.Get("FromEmail");
                String UserName = ConfigurationManager.AppSettings.Get("EmailUserName");
                String Password = ConfigurationManager.AppSettings.Get("EmailPassword");
                String To = ConfigurationManager.AppSettings.Get("ToEmail");
                if (input.Contains("STATUS"))
                    To = ConfigurationManager.AppSettings.Get("DailystatusEmail");
                else if (input.Contains("DAILY_FEDERAL_REPORT"))
                    To = ConfigurationManager.AppSettings.Get("DailyFederalReportEmail");
                List<string> ToMailID = new List<string>(To.Split(';'));
                CrawlerMonitor.Web.SendHtmlEmail(From, message, ToMailID, emailSubject, UserName, Password, SMTP);
            }
            catch (Exception Ex)
            {
                Logger.LogMessage(Ex.Message, LogMessageLevel.B_Error);
            }
        }
        #endregion

        #region statusValidations
        /// <summary>
        /// statusValidations
        /// </summary>
        /// <returns></returns>
        public static DataTable statusValidations(string dbName, string siteName)
        {
            string completedCnt = string.Empty;
            string pendingCnt = string.Empty;
            string runningCnt = string.Empty;
            DataTable dt = new DataTable();
            dt.Columns.Add("Site Name");
            dt.Columns.Add("Completed Count");
            dt.Columns.Add("Running Count");
            dt.Columns.Add("Pending Count");
            dt.Columns.Add("Status");
            completedCnt = getCompletedStatusCount(dbName).ToString();
            pendingCnt = getPendingStatusCount(dbName).ToString();
            runningCnt = getRunningStatusCount(dbName).ToString();

            if (pendingCnt == "0" && runningCnt == "0" && completedCnt != "0")
                dt.Rows.Add(siteName.ToUpper().Trim(), completedCnt, pendingCnt, runningCnt, "Completed");
            else
                dt.Rows.Add(siteName.ToUpper().Trim(), completedCnt, pendingCnt, runningCnt, "Running");

            return dt;
        }
        #endregion

        #region getCompletedStatusCount
        /// <summary>
        /// getCompletedStatusCount
        /// </summary>
        /// <param name="dbName"></param>
        /// <returns></returns>
        public static int getCompletedStatusCount(string dbName)
        {
            //DATABASE Connection esstablished
            database = Utility.GetDBConnection();
            int cnt = 0;
            try
            {
                string qry = string.Empty;
                qry = "SELECT count(*) FROM " + dbName + "(nolock) where status = 1";
                cnt = Convert.ToInt32(database.GetScalarValue(qry));
            }
            catch (Exception e)
            {
                Logger.LogMessage("Failed to Generate the Completed status " + e.Message.ToString(), LogMessageLevel.B_Error);
            }
            return cnt;
        }
        #endregion

        #region getPendingStatusCount
        /// <summary>
        /// getPendingStatusCount
        /// </summary>
        /// <param name="dbName"></param>
        /// <returns></returns>
        public static int getPendingStatusCount(string dbName)
        {
            //DATABASE Connection esstablished
            database = Utility.GetDBConnection();
            int cnt = 0;
            string qry = string.Empty;
            try
            {
                qry = "SELECT count(*) FROM " + dbName + "(nolock) where status = 0";
                cnt = Convert.ToInt32(database.GetScalarValue(qry));
            }
            catch (Exception e)
            {
                Logger.LogMessage("Failed to Generate the Pending status " + e.Message.ToString(), LogMessageLevel.B_Error);
            }
            return cnt;
        }
        #endregion

        #region getRunningStatusCount
        /// <summary>
        /// getRunningStatusCount
        /// </summary>
        /// <param name="dbName"></param>
        /// <returns></returns>
        public static int getRunningStatusCount(string dbName)
        {
            //DATABASE Connection esstablished
            database = Utility.GetDBConnection();
            int cnt = 0;
            string qry = string.Empty;
            try
            {
                qry = "SELECT count(*) FROM " + dbName + "(nolock) where status != 1 and status != 0";
                cnt = Convert.ToInt32(database.GetScalarValue(qry));
            }
            catch (Exception e)
            {
                Logger.LogMessage("Failed to Generate the Running status " + e.Message.ToString(), LogMessageLevel.B_Error);
            }
            return cnt;
        }
        #endregion

        #region GenerateReport
        /// <summary>
        /// GenerateReport
        /// </summary>
        /// <param name="inp"></param>
        /// <returns></returns>
        private static DataSet GenerateReport(string inp)
        {
            string qry = string.Empty;
            DataSet dsDaily = new DataSet("CrawlerMonitor");
            //DB Connection
            database = Utility.GetDBConnection();
            DataTable programSchedule = new DataTable();
            string today = DateTime.Now.ToShortDateString();
            string prvDate = DateTime.Now.ToShortDateString();

            if (inp == "DAILY" || inp == "ALL")
            {
                DataTable Daily = new DataTable();
                prvDate = DateTime.Now.AddDays(-2).ToShortDateString();
                //DateTime.ParseExact(prvDate, "yyyyMMdd", CultureInfo.InvariantCulture);
                qry = "SELECT distinct a.SourceID, a.sitename, a.Priority, cast(isnull(a.lastCrawl,(SELECT MIN(CAST(createddate as date)) ";
                qry += "FROM [Utility].[dbo].[PopXMLSQSMessages])) as varchar) as lastCrawl, UPPER(a.schedule) as Schedule from (SELECT ps.SourceID, ps.SiteName, ps.schedule, MAX(CAST(ss.CreatedDate as date)) as lastCrawl, ps.Priority ";
                qry += "FROM [configsetting].[dbo].[ProgramSchedule] as ps (nolock) LEFT JOIN [Utility].[dbo].[PopXMLSQSMessages] as ss (nolock) ON ps.sitename = ss.sourcename ";
                qry += "where ps.Enabled = 1 and ps.SourceID != 'NULL' and ps.SiteName not in (SELECT distinct sourcename FROM [Utility].[dbo].[PopXMLSQSMessages] (nolock) ";
                qry += "where CAST(createddate AS DATE) between '" + DateTime.Parse(prvDate).ToString("yyyyMMdd") + "' and '" + DateTime.Parse(today).ToString("yyyyMMdd") + "') ";
                qry += "AND Schedule = 'Daily' and ps.SourceID is not null group by ps.SiteName, ps.SourceID, ps.Schedule, ps.Priority) as a where a.SiteName not in (SELECT sitename ";
                qry += "FROM [webcrawlerconfiguration].[dbo].[ScanStatus] (nolock) where CAST(ExtractionStartTime as DATE) between '" + DateTime.Parse(prvDate).ToString("yyyyMMdd") + "' and '" + DateTime.Parse(today).ToString("yyyyMMdd") + "' and CrawlerExporterStatus = 'completed' ";
                qry += "and SiteName in ( SELECT distinct SiteName FROM [configsetting].[dbo].[ProgramSchedule] (nolock) where Enabled = 1 and Schedule = 'Daily' and SourceID != 'NULL' and (sitename not like '%_First' AND sitename not like '%_Link')) ";
                qry += "group by SiteName) and (sitename not like '%_First' AND sitename not like '%_Link') order by a.Priority, a.SiteName";

                database.GetDataTable(qry, Daily);
                dsDaily.Tables.Add(Daily);
            }

            if (inp == "WEEKLY" || inp == "ALL")
            {
                qry = string.Empty;
                DataTable Weekly = new DataTable();
                prvDate = DateTime.Now.AddDays(-7).ToShortDateString();
                qry = "SELECT distinct a.SourceID, a.sitename, a.Priority, cast(isnull(a.lastCrawl,(SELECT MIN(CAST(createddate as date)) ";
                qry += "FROM [Utility].[dbo].[PopXMLSQSMessages])) as varchar) as lastCrawl, UPPER(a.schedule) as Schedule from (SELECT ps.SourceID, ps.SiteName, ps.schedule, MAX(CAST(ss.CreatedDate as date)) as lastCrawl, ps.Priority ";
                qry += "FROM [configsetting].[dbo].[ProgramSchedule] as ps (nolock) LEFT JOIN [Utility].[dbo].[PopXMLSQSMessages] as ss (nolock) ON ps.sitename = ss.sourcename ";
                qry += "where ps.Enabled = 1 and ps.SourceID != 'NULL' and ps.SiteName not in (SELECT distinct sourcename FROM [Utility].[dbo].[PopXMLSQSMessages] (nolock) ";
                qry += "where CAST(createddate AS DATE) between '" + DateTime.Parse(prvDate).ToString("yyyyMMdd") + "' and '" + DateTime.Parse(today).ToString("yyyyMMdd") + "') ";
                qry += "AND Schedule like '%Weekly%' and ps.SourceID is not null group by ps.SiteName, ps.SourceID, ps.Schedule, ps.Priority) as a where a.SiteName not in (SELECT sitename ";
                qry += "FROM [webcrawlerconfiguration].[dbo].[ScanStatus] (nolock) where CAST(ExtractionStartTime as DATE) between '" + DateTime.Parse(prvDate).ToString("yyyyMMdd") + "' and '" + DateTime.Parse(today).ToString("yyyyMMdd") + "' and CrawlerExporterStatus = 'completed' ";
                qry += "and SiteName in ( SELECT distinct SiteName FROM [configsetting].[dbo].[ProgramSchedule] (nolock) where Enabled = 1 and Schedule like '%Weekly%' and SourceID != 'NULL' and (sitename not like '%_First' AND sitename not like '%_Link')) ";
                qry += "group by SiteName) and (sitename not like '%_First' AND sitename not like '%_Link') order by a.Priority, a.SiteName";

                database.GetDataTable(qry, Weekly);
                dsDaily.Tables.Add(Weekly);
            }

            if (inp == "MONTHLY" || inp == "ALL")
            {
                qry = string.Empty;
                DataTable Monthly = new DataTable();
                prvDate = DateTime.Now.AddDays(-30).ToShortDateString();
                qry = "SELECT distinct a.SourceID, a.sitename, a.Priority, cast(isnull(a.lastCrawl,(SELECT MIN(CAST(createddate as date)) ";
                qry += "FROM [Utility].[dbo].[PopXMLSQSMessages])) as varchar) as lastCrawl, UPPER(a.schedule) as Schedule from (SELECT ps.SourceID, ps.SiteName, ps.schedule, MAX(CAST(ss.CreatedDate as date)) as lastCrawl, ps.Priority ";
                qry += "FROM [configsetting].[dbo].[ProgramSchedule] as ps (nolock) LEFT JOIN [Utility].[dbo].[PopXMLSQSMessages] as ss (nolock) ON ps.sitename = ss.sourcename ";
                qry += "where ps.Enabled = 1 and ps.SourceID != 'NULL' and ps.SiteName not in (SELECT distinct sourcename FROM [Utility].[dbo].[PopXMLSQSMessages] (nolock) ";
                qry += "where CAST(createddate AS DATE) between '" + DateTime.Parse(prvDate).ToString("yyyyMMdd") + "' and '" + DateTime.Parse(today).ToString("yyyyMMdd") + "') ";
                qry += "AND Schedule like '%monthly%' and ps.SourceID is not null group by ps.SiteName, ps.SourceID, ps.Schedule, ps.Priority) as a where a.SiteName not in (SELECT sitename ";
                qry += "FROM [webcrawlerconfiguration].[dbo].[ScanStatus] (nolock) where CAST(ExtractionStartTime as DATE) between '" + DateTime.Parse(prvDate).ToString("yyyyMMdd") + "' and '" + DateTime.Parse(today).ToString("yyyyMMdd") + "' and CrawlerExporterStatus = 'completed' ";
                qry += "and SiteName in ( SELECT distinct SiteName FROM [configsetting].[dbo].[ProgramSchedule] (nolock) where Enabled = 1 and Schedule like '%monthly%' and SourceID != 'NULL' and (sitename not like '%_First' AND sitename not like '%_Link')) ";
                qry += "group by SiteName) and (sitename not like '%_First' AND sitename not like '%_Link') order by a.Priority, a.SiteName";

                database.GetDataTable(qry, Monthly);
                dsDaily.Tables.Add(Monthly);
            }

            if (inp == "QUARTELY" || inp == "ALL")
            {
                qry = string.Empty;
                DataTable Yearly = new DataTable();
                prvDate = DateTime.Now.AddDays(-90).ToShortDateString();
                qry = "SELECT distinct a.SourceID, a.sitename, a.Priority, cast(isnull(a.lastCrawl,(SELECT MIN(CAST(createddate as date)) ";
                qry += "FROM [Utility].[dbo].[PopXMLSQSMessages])) as varchar) as lastCrawl, UPPER(a.schedule) as Schedule from (SELECT ps.SourceID, ps.SiteName, ps.schedule, MAX(CAST(ss.CreatedDate as date)) as lastCrawl, ps.Priority ";
                qry += "FROM [configsetting].[dbo].[ProgramSchedule] as ps (nolock) LEFT JOIN [Utility].[dbo].[PopXMLSQSMessages] as ss (nolock) ON ps.sitename = ss.sourcename ";
                qry += "where ps.Enabled = 1 and ps.SourceID != 'NULL' and ps.SiteName not in (SELECT distinct sourcename FROM [Utility].[dbo].[PopXMLSQSMessages] (nolock) ";
                qry += "where CAST(createddate AS DATE) between '" + DateTime.Parse(prvDate).ToString("yyyyMMdd") + "' and '" + DateTime.Parse(today).ToString("yyyyMMdd") + "') ";
                qry += "AND Schedule like '%Yearly%' and ps.SourceID is not null group by ps.SiteName, ps.SourceID, ps.Schedule, ps.Priority) as a where a.SiteName not in (SELECT sitename ";
                qry += "FROM [webcrawlerconfiguration].[dbo].[ScanStatus] (nolock) where CAST(ExtractionStartTime as DATE) between '" + DateTime.Parse(prvDate).ToString("yyyyMMdd") + "' and '" + DateTime.Parse(today).ToString("yyyyMMdd") + "' and CrawlerExporterStatus = 'completed' ";
                qry += "and SiteName in ( SELECT distinct SiteName FROM [configsetting].[dbo].[ProgramSchedule] (nolock) where Enabled = 1 and Schedule like '%Yearly%' and SourceID != 'NULL' and (sitename not like '%_First' AND sitename not like '%_Link')) ";
                qry += "group by SiteName) and (sitename not like '%_First' AND sitename not like '%_Link') order by a.Priority, a.SiteName";

                database.GetDataTable(qry, Yearly);
                dsDaily.Tables.Add(Yearly);
            }

            DataTable dtAll = new DataTable();

            if (dsDaily.Tables.Count > 1)
            {
                for (int i = 0; i < dsDaily.Tables.Count; i++)
                {
                    dtAll.Merge(dsDaily.Tables[i]);
                }
                dsDaily = new DataSet();

                DataView dataview = dtAll.DefaultView;
                dataview.Sort = "Priority";
                DataTable dt = dataview.ToTable();
                dsDaily.Tables.Add(dt);
            }

            //Get the list of crawlers which are not able to develop it
            if (dsDaily.Tables.Count > 0)
            {
                qry = string.Empty;
                DataTable CrawlerFailure = new DataTable();
                qry = "SELECT SiteName FROM [Utility].[dbo].[CrawlerFailure] where Enabled = 1";
                database.GetDataTable(qry, CrawlerFailure);
                DataTable dt = dsDaily.Tables[0];
                DataTable dsm = getLinq(dt, CrawlerFailure);
                dsDaily = new DataSet();
                dsDaily.Tables.Add(dsm);
            }
            return dsDaily;
        }
        #endregion

        #region getLinq
        /// <summary>
        /// getLinq
        /// </summary>
        /// <param name="dt1"></param>
        /// <param name="dt2"></param>
        /// <returns></returns>
        private static DataTable getLinq(DataTable dt1, DataTable dt2)
        {
            var qry1 = dt1.AsEnumerable().Select(a => new { SiteName = a["SourceName"].ToString() });
            var qry2 = dt2.AsEnumerable().Select(b => new { SiteName = b["SourceName"].ToString() });

            var exceptAB = qry1.Except(qry2);

            DataTable dtMisMatch = (from a in dt1.AsEnumerable()
                                    join ab in exceptAB on a["SourceName"].ToString() equals ab.SiteName
                                    select a).CopyToDataTable();
            return dtMisMatch;
        }
        #endregion

        #region Generate CSV File
        /// <summary>
        /// GenerateCSVFile
        /// </summary>
        /// <param name="dtReports"></param>
        /// <param name="lst_folder_path"></param>
        /// <param name="reportType"></param>
        private static void GenerateCSVFile(DataSet dtReports, List<string> lst_folder_path, string reportType)
        {
            for (int _tblCnt = 0; _tblCnt < dtReports.Tables.Count; _tblCnt++)
            {
                try
                {
                    StreamWriter sw = new StreamWriter(lst_folder_path[_tblCnt]);
                    bool check = false;
                    foreach (DataColumn dColumn in dtReports.Tables[_tblCnt].Columns)
                    {
                        if (check)
                            sw.Write(",");
                        sw.Write(dColumn.ColumnName);
                        check = true;
                    }
                    sw.Write(sw.NewLine);
                    foreach (DataRow drRow in dtReports.Tables[_tblCnt].Rows)
                    {
                        for (int i = 0; i < dtReports.Tables[_tblCnt].Columns.Count; i++)
                        {
                            if (i != 0) sw.Write(",");
                            if (reportType == "DailyReport")
                            {
                                if (i == 3)
                                {
                                    string priority = drRow[i].ToString();
                                    if (priority == "1") priority = "Critical";
                                    else if (priority == "2") priority = "High";
                                    else if (priority == "3") priority = "Medium";
                                    else priority = "Low";
                                    sw.Write("\"" + priority.Replace("\"", "\"\"") + "\"");
                                }
                                else
                                    sw.Write("\"" + drRow[i].ToString().Replace("\"", "\"\"") + "\"");
                            }
                            else
                                sw.Write("\"" + drRow[i].ToString().Replace("\"", "\"\"") + "\"");
                        }
                        sw.Write(sw.NewLine);
                    }
                    sw.Flush();
                    sw.Close();
                }
                catch (Exception ex)
                {
                    Logger.LogMessage("Failed to create report" + ex.Message.ToString(), LogMessageLevel.B_Error);
                }
            }
        }
        #endregion

        #region sendEmail for XML FAULT
        /// <summary>
        /// sendEmail
        /// </summary>
        /// <param name="inp"></param>
        /// <param name="files"></param>
        public static void sendEmail(string inp, List<string> files)
        {
            string emailBody = string.Empty;
            emailBody += "Hi All<BR><BR>";
            emailBody += "Please find the attached report of Failure in XML generation for " + inp + " Crawlers. <BR>";

            emailBody += "<BR><B>Thanks & Regards,</B><BR>";
            emailBody += "Crawler Monitoring Team";
            string emailSubject = "Failure in XML generation report for " + inp + " Crawlers";
            try
            {
                String SMTP = ConfigurationManager.AppSettings.Get("SMTP");
                String From = ConfigurationManager.AppSettings.Get("FromEmail");
                String To = ConfigurationManager.AppSettings.Get("ToEmail");
                String UserName = ConfigurationManager.AppSettings.Get("EmailUserName");
                String Password = ConfigurationManager.AppSettings.Get("EmailPassword");
                List<string> ToMailID = new List<string>(To.Split(';'));
                CrawlerMonitor.Web.SendHtmlEmailWithAttachment(From, emailBody, ToMailID, emailSubject, UserName, Password, SMTP, files);
                foreach (string file in files)
                {
                    if (File.Exists(file)) System.IO.File.Delete(file);
                }
            }
            catch (Exception Ex)
            {
                Logger.LogMessage(Ex.Message, LogMessageLevel.B_Error);
            }
        }
        #endregion

        #region dailyFederalCralwerReport
        /// <summary>
        /// dailyFederalCralwerReport
        /// </summary>
        public static void dailyFederalCralwerReport(string inp)
        {
            string qry = string.Empty;
            string today = DateTime.Now.ToShortDateString();
            string currTime = DateTime.Now.ToShortTimeString();
            string siteName = string.Empty;
            string dbName = string.Empty;
            database = Utility.GetDBConnection();
            DataTable crawlerReport = new DataTable();
            crawlerReport.Columns.Add("Site Name");
            crawlerReport.Columns.Add("Crawled Records");
            crawlerReport.Columns.Add("Completion Time");
            crawlerReport.Columns.Add("Status");
            bool statusCheck = false;
            string messageMail = string.Empty;
            siteName = ConfigurationManager.AppSettings.Get("DAILY_FEDERAL_REPORT");
            List<string> siteNameList = new List<string>(siteName.Split(','));
            string message = siteName.Replace(',', '&');
            //Process.Start("cmd.exe", "/K " + "C:/Crawler/DataCrawler.exe");
            try
            {
                foreach (string sitename in siteNameList)
                {
                    statusCheck = false;
                    qry = string.Empty;
                    //Empty sitename check 
                    if (sitename == string.Empty)
                        continue;
                    // Read the DB and Table name from Config file
                    dbName = ConfigurationManager.AppSettings.Get(sitename.ToUpper().Trim());

                    //Verify the crawler launched status
                    qry = "SELECT top 1 time FROM [configsetting].[dbo].[ProgramSchedule](nolock) where sitename = '" + sitename.ToUpper().Trim() + "' order by time asc";
                    string schedulerTime = database.GetScalar(qry).ToString();

                    if (Convert.ToDateTime(currTime) < Convert.ToDateTime(schedulerTime))
                    {
                        crawlerReport.Rows.Add(sitename.ToUpper().Trim(), "-", "-", "Not started");
                        messageMail = "Crawler is not started";
                        continue;
                    }
                    if (!string.IsNullOrEmpty(dbName))
                    {
                        DataTable dt = new DataTable();
                        if (getCompletedStatusCount(dbName).ToString() != "0" && getPendingStatusCount(dbName).ToString() == "0" && getRunningStatusCount(dbName).ToString() == "0")
                        {
                            qry = string.Empty;
                            qry = "SELECT * FROM [webcrawlerconfiguration].[dbo].[ScanStatus] (nolock) where sitename = '" + sitename.ToUpper().Trim() + "' and cast(extractionstarttime as datetime) >= '" + today + "'";

                            DataTable dtScan = new DataTable();
                            database.GetDataTable(qry, dtScan);
                            //Verify all the instances status
                            if (dtScan.Rows.Count == dtScan.Select("CrawlerExporterStatus = 'Completed'").Length)
                            {
                                string status = dtScan.Compute("SUM(SentToAtribusCount)", "").ToString() + "Records Sent to Atribus. <br>" + dtScan.Compute("SUM(Total_No_Records)", "").ToString() + " Cleaned records are exported to S3.";
                                crawlerReport.Rows.Add(sitename.ToUpper().Trim(), Convert.ToInt32(dtScan.Compute("SUM(SentToAtribusCount)", "")) + Convert.ToInt32(dtScan.Compute("SUM(Total_No_Records)", "")), Convert.ToDateTime(dtScan.Compute("Max(ExtractionEndTime)", "")), status);
                                statusCheck = true;
                            }
                            else
                            {
                                crawlerReport.Rows.Add(sitename.ToUpper().Trim(), "-", "-", "Crawler is not exported.");
                                messageMail = "Crawler is not exported";
                                continue;
                            }
                        }
                        else
                        {
                            getStatus("DAILY_STATUS");
                            return;
                        }
                    }
                    else
                    {
                        crawlerReport.Rows.Add(sitename.ToUpper().Trim(), "-", "-", "Misconfigured");
                        messageMail = "Site Name is Misconfigured";
                        continue;
                    }
                }
                if (statusCheck)
                    sendStatusMail(crawlerReport, inp, "<h4>" + message + " Crawlers Daily Report on " + DateTime.Now.ToShortDateString() + "</h4>", message + " Crawlers Daily Report ");
                else
                    sendFailureReportsMail(crawlerReport, messageMail);
            }
            catch (Exception e)
            {
                Logger.LogMessage("" + e.Message.ToString(), LogMessageLevel.B_Error);
            }
        }
        #endregion

        #region sendFailureReportsMail
        /// <summary>
        /// sendFailureReportsMail
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="msg"></param>
        public static void sendFailureReportsMail(DataTable dt, string msg)
        {
            string message = string.Empty;
            message += "<h4> Report Generation is failed because of " + msg + ". </h4>";
            message += "<table  style='border-collapse:collapse;font-family:verdana,helvetica,arial,sans-serif;border:1px solid #c3c3c3'>";
            message += "<tr>";
            try
            {
                foreach (DataColumn column in dt.Columns)
                {
                    message += "<th style='white-space:nowrap;background-color:#555555;color:#fff;border:1px solid #c3c3c3;padding:3px;vertical-align:top;text-align:center;font-size:11px'>";
                    message += column.ColumnName.ToUpper();
                }
                message += "</tr>";
                int rowCount = 0;

                foreach (DataRow rows in dt.Rows)
                {
                    message += "<tr>";
                    rowCount++;
                    for (int i = 0; i < dt.Columns.Count; ++i)
                    {
                        if (rowCount % 2 == 0)
                        {
                            message += "<td style='vertical-align:top;text-align:left;font-size:11px;background-color:#DDDDDD;padding:3px;border:1px solid #c3c3c3'>";
                        }
                        else
                        {
                            message += "<td style='vertical-align:top;text-align:left;font-size:11px;background-color:#FFB581;padding:3px;border:1px solid #c3c3c3'>";
                        }
                        message += rows[i].ToString();
                    }
                    message += "</tr>";
                }
                message += "</table><br>";
                string emailSubject = msg + " - " + DateTime.Now.ToString();
                String SMTP = ConfigurationManager.AppSettings.Get("SMTP");
                String From = ConfigurationManager.AppSettings.Get("FromEmail");
                String UserName = ConfigurationManager.AppSettings.Get("EmailUserName");
                String Password = ConfigurationManager.AppSettings.Get("EmailPassword");
                String To = ConfigurationManager.AppSettings.Get("DailyFederalReportFailureEmail");
                List<string> ToMailID = new List<string>(To.Split(';'));
                CrawlerMonitor.Web.SendHtmlEmail(From, message, ToMailID, emailSubject, UserName, Password, SMTP);
            }
            catch (Exception e)
            {
                Logger.LogMessage("Report Generation is failed " + e.Message.ToString(), LogMessageLevel.B_Error);
            }
        }
        #endregion

        #region docDownloadFederalReport
        /// <summary>
        /// docDownloadFederalReport
        /// </summary>
        /// <returns></returns>
        public static bool docDownloadFederalReport(string input)
        {
            database = Utility.GetDBConnection();
            DataTable crawlerReport = new DataTable();
            string siteName = ConfigurationManager.AppSettings.Get(input);

            Process[] processlist = Process.GetProcesses();

            foreach (Process theprocess in processlist)
            {
                Console.WriteLine("Process: {0} ID: {1}", theprocess.ProcessName, theprocess.Id);
            }

            return true;
        }
        #endregion

        #region faultReport
        /// <summary>
        /// faultReport
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static bool faultReport(string input)
        {
            string qry = string.Empty;
            DataSet dsDaily = new DataSet("CrawlerMonitor");
            //DB Connection
            database = Utility.GetDBConnection();
            var allCrawlers = new List<KeyValuePair<string, string>>();
            allCrawlers.Add(new KeyValuePair<string, string>("DAILY", "3"));
            allCrawlers.Add(new KeyValuePair<string, string>("WEEKLY", "7"));
            allCrawlers.Add(new KeyValuePair<string, string>("MONTHLY", "30"));
            allCrawlers.Add(new KeyValuePair<string, string>("QUARTELY", "90"));
            DataTable programSchedule = new DataTable();
            string today = DateTime.Now.ToShortDateString();
            qry = "SELECT * FROM [configsetting].[dbo].[ProgramSchedule] where Enabled = 1";
            database.GetDataTable(qry, programSchedule);
            try
            {

                qry = string.Empty;
                DataTable dt = new DataTable();
                DataTable DailySQS = new DataTable();
                
                foreach (var pair in allCrawlers)
                {
                    if (pair.Key.ToUpper() == input.ToUpper() || input.ToUpper() == "ALL")
                    {
                        DataTable daily = new DataTable(pair.Key.ToUpper());
                        daily.Columns.Add("SourceID");
                        daily.Columns.Add("SourceName");
                        daily.Columns.Add("Priority");
                        daily.Columns.Add("Last_Crawl_Date");
                        daily.Columns.Add("Schedule");
                        dt = getFailedSources(pair.Key.ToUpper(), pair.Value);
                        qry = "SELECT distinct cast(sourceID as varchar) as SourceID, cast(sourcename as varchar) as SourceName, max(convert(varchar(20), createddate,1)) as Last_Crawl_Date FROM [Utility].[dbo].[PopXMLSQSMessages] group by sourceid, sourcename";
                        database.GetDataTable(qry, DailySQS);
                        foreach (DataRow row in dt.Rows)
                        {
                            //DataRow[] drsqs = DailySQS.Select("SourceID='" +row[0].ToString()+"' AND SourceName='" +row[1].ToString() + "'");
                            DataRow drsqs = DailySQS.AsEnumerable().FirstOrDefault(r => r.Field<string>("SourceID") == row[0].ToString() && r.Field<string>("SourceName") == row[1].ToString());
                            DataRow drps = programSchedule.AsEnumerable().FirstOrDefault(r => r.Field<string>("sitename") == row[1].ToString());
                            string lastCrawlDate = drsqs == null ? DateTime.Now.AddDays(-150).ToShortDateString() : drsqs[2].ToString();
                            daily.Rows.Add(row[0].ToString(), row[1].ToString(), drps[15].ToString().ToUpper(), lastCrawlDate, drps[9].ToString().ToUpper());
                        }
                        dsDaily.Tables.Add(daily);
                    }
                    else
                        continue;
                }
                
                DataTable dtAll = new DataTable();
                if (dsDaily.Tables.Count > 1)
                {
                    for (int i = 0; i < dsDaily.Tables.Count; i++)
                    {
                        dtAll.Merge(dsDaily.Tables[i]);
                    }
                    dsDaily = new DataSet();

                    DataView dataview = dtAll.DefaultView;
                    dataview.Sort = "Priority";
                    DataTable dtView = dataview.ToTable();
                    dsDaily.Tables.Add(dtView);
                }

                //Get the list of crawlers which are not able to develop it
                if (dsDaily.Tables.Count > 0)
                {
                    qry = string.Empty;
                    DataTable CrawlerFailure = new DataTable();
                    qry = "SELECT SiteName as SourceName FROM [Utility].[dbo].[CrawlerFailure] where Enabled = 1";
                    database.GetDataTable(qry, CrawlerFailure);
                    DataTable dtChk = dsDaily.Tables[0];
                    DataTable dsm = getLinq(dtChk, CrawlerFailure);
                    dsDaily = new DataSet();
                    dsDaily.Tables.Add(dsm);
                }
                string localFileName = FileOperation.GetRootPath("TEMP");
                string xmlFileName = string.Empty;
                List<string> files = new List<string>();
                
                localFileName = FileOperation.GetRootPath("TEMP");
                xmlFileName = "CrawlerMonitoring_" + input + ".csv";
                localFileName = Text.TextBeforeLastTag(localFileName, "/").Replace("/", "\\") + "\\" + xmlFileName;
                files.Add(localFileName);

                if (dsDaily.Tables.Count > 0)
                    GenerateCSVFile(dsDaily, files, "CrawlerMonitoringDaily");
                else
                    Logger.LogMessage("No Records Found ", LogMessageLevel.B_Error);

                sendEmail(input, files);


            }
            catch (Exception e)
            {
                Logger.LogMessage("" + e.Message.ToString(), LogMessageLevel.B_Error);
            }
            return false;
        }
        #endregion

        #region pscount
        /// <summary>
        /// getFailedSources
        /// </summary>
        /// <param name="frequency"></param>
        /// <param name="sourceID"></param>
        /// <returns></returns>
        public static DataTable getFailedSources(string input, string frequency)
        {
            string qry = string.Empty;
            DataTable dt = new DataTable();
            database = Utility.GetDBConnection();
            qry = "select distinct sourceid, sitename FROM [configsetting].[dbo].[ProgramSchedule] where Enabled = 1 and Schedule like '%" + input + "%' and SourceID is not null and sourceID != 'NULL' and sitename not like '%_First' AND sitename not like '%_Link' ";
            qry += "AND cast(SourceID as varchar) not in (select distinct cast(sourceID as varchar) from [Utility].[dbo].[PopXMLSQSMessages] ";
            qry += "where convert(VARCHAR(20),createddate,1) >= convert(varchar(20),DATEADD(d,-" + frequency + ",GETDATE()), 1)) and sitename not in (select distinct sourcename ";
            qry += "from [Utility].[dbo].[PopXMLSQSMessages] where convert(VARCHAR(20),createddate,1) >= convert(varchar(20),DATEADD(d,-" + frequency + ",GETDATE()), 1)) order by SiteName";
            database.GetDataTable(qry, dt);
            return dt;
        }
        #endregion

        #region getSitenamefromPS
        /// <summary>
        /// getSitenamefromPS
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static DataTable getSitenamefromPS(string input)
        {
            database = Utility.GetDBConnection();
            DataTable dt = new DataTable();
            string qry = string.Empty;
            qry = "SELECT distinct SourceID, SiteName FROM [configsetting].[dbo].[ProgramSchedule] where Enabled = 1 and ";
            qry += "Schedule like '%" + input + "%' and SourceID is not null and sourceID != 'NULL' and sitename not like '%_First' AND sitename not like '%_Link' order by SiteName";
            database.GetDataTable(qry, dt);

            return dt;
        }
        #endregion

    }
}
