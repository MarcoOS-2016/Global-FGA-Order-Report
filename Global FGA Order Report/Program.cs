using System;
using System.IO;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Global_FGA_Order_Report
{
    public class Program
    {
        static void Main(string[] args)
        {
            GlobalFGAOrderReport.CleanCookie();

            string defaulturl = "http://auspwfooedweb02.aus.amer.dell.com:8080/ODM/Default.aspx";
            CookieContainer cookies = new CookieContainer();
            CookieCollection cookie = GlobalFGAOrderReport.FetchCookie(defaulturl);
            cookies.Add(cookie);

            string posturl = ConfigFileUtility.GetValue("Post_URL");
            string outputfolder = ConfigFileUtility.GetValue("Output_Folder");
            string outputfilename = Path.Combine(outputfolder, string.Format("Global_FGA_Order_Report_{0}.xlsx", DateTime.Now.ToString("yyyyMMdd_HHmmss")));

            MiscUtility.LogHistory("Start to fetch Global FGA Order Report from the FDL Website...");
            Console.WriteLine(string.Format("[{0}] - Start to fetch Global FGA Order Report from the FDL Website...", DateTime.Now.ToString()));

            string htmltext = GlobalFGAOrderReport.PostByWebRequest(posturl, cookies);
            
            Console.WriteLine(string.Format("[{0}] - Done!", DateTime.Now.ToString()));
            MiscUtility.LogHistory("Done!");

            //string htmlfilename = "Global FGA Order Report.html";
            //FileUtility.SaveFile(htmlfilename, htmldoc);
            //string htmltext = FileUtility.LoadTextFile(htmlfilename);

            MiscUtility.LogHistory(string.Format("Start to save report into the excel file - {0}...", outputfilename));
            Console.WriteLine(string.Format("[{0}] - Start to save report into the excel file - {1}...", DateTime.Now.ToString(), outputfilename));
            
            GlobalFGAOrderReport.ExportTableToExcel(outputfilename, htmltext);
            
            Console.WriteLine(string.Format("[{0}] - Done!", DateTime.Now.ToString()));
            MiscUtility.LogHistory("Done!");
        }
    }
}
