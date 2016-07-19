using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CrawlerMonitor
{
    class Program
    {
        static void Main(string[] args)
        {
            string typeSelection = string.Empty;

            if (args.Length == 0)
            {
                Console.WriteLine("Please pass parameter like Daily-Weekly-Monthly-Quaterly....");
                Console.ReadKey();
            }
            foreach (string parameter in args)
            {

                if (parameter.ToUpper().Contains("FAULT_DAILY"))
                {
                    Utility.faultReport("DAILY");
                }
                else if (parameter.ToUpper().Contains("FAULT_WEEKLY"))
                {
                    Utility.faultReport("WEEKLY");
                }
                else if (parameter.ToUpper().Contains("FAULT_MONTHLY"))
                {
                    Utility.faultReport("MONTHLY");
                }
                else if (parameter.ToUpper().Contains("FAULT_QUARTELY"))
                {
                    Utility.faultReport("QUARTELY");
                }
                else if (parameter.ToUpper().Contains("FAULT_ALL"))
                {
                    Utility.faultReport("ALL");
                }
                //typeSelection = string.Empty;
                else if (parameter.ToUpper().Contains("DAILY_STATUS"))
                {
                    Utility.getStatus("DAILY_STATUS");
                }
                else if (parameter.ToUpper().Contains("DAILY_FEDERAL_REPORT"))
                {
                    Utility.dailyFederalCralwerReport("DAILY_FEDERAL_REPORT");
                }
                else if (parameter.ToUpper().Contains("DAILY_FEDERAL_DOCDOWNLOAD_REPORT"))
                {
                    Utility.docDownloadFederalReport("DAILY_FEDERAL_DOCDOWNLOAD_REPORT");
                }
            }
        }
    }
}
