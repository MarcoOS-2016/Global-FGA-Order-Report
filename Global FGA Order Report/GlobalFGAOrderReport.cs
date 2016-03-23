using System;
using System.IO;
using System.Net;
using System.Web;
using System.Xml;
using System.Linq;
using System.Text;
using System.Data;
using HtmlAgilityPack;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Global_FGA_Order_Report
{
    public class GlobalFGAOrderReport
    {
        public static string PostByWebRequest(string posturl, CookieContainer cookies)
        {
            try
            {
                HttpWebRequest request = WebRequest.Create(posturl) as HttpWebRequest;
                request.Method = "GET";
                request.KeepAlive = false;

                //Get the response
                HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                System.IO.Stream responseStream = response.GetResponseStream();
                System.IO.StreamReader reader = new System.IO.StreamReader(responseStream, Encoding.UTF8);
                string srcString = reader.ReadToEnd();

                //Get the ViewState
                string viewStateFlag = "id=\"__VIEWSTATE\" value=\"";
                int i = srcString.IndexOf(viewStateFlag) + viewStateFlag.Length;
                int j = srcString.IndexOf("\"", i);
                string viewState = srcString.Substring(i, j - i);

                //Get the ViewState
                string EventValidationFlag = "id=\"__EVENTVALIDATION\" value=\"";
                i = srcString.IndexOf(EventValidationFlag) + EventValidationFlag.Length;
                j = srcString.IndexOf("\"", i);
                string eventValidation = srcString.Substring(i, j - i);

                viewState = Uri.EscapeDataString(viewState);
                eventValidation = Uri.EscapeDataString(eventValidation);

                string formatString = "__EVENTTARGET=btnShowReport&__EVENTARGUMENT=&__VIEWSTATE={0}&__EVENTVALIDATION={1}";
                string postString = string.Format(formatString, viewState, eventValidation);
                //string postString = "__EVENTTARGET=btnShowReport&__EVENTARGUMENT=&__VIEWSTATE=%2FwEPDwUKMTEyMzQxODg3N2QYAQUeX19Db250cm9sc1JlcXVpcmVQb3N0QmFja0tleV9fFgEFDGNoZWNrSXNFeGNlbD3kkz4%2FiQWP9XmohgEuveoQQN3n&__EVENTVALIDATION=%2FwEWAwLHrp7gCAL3lYZCAsPJ8dUFDJaAy2Ce0w7CzX3ppztvp8NtkEY%3D";
                //Change to byte[]
                byte[] postData = Encoding.ASCII.GetBytes(postString);

                //Compose the new request
                request = WebRequest.Create(posturl) as HttpWebRequest;
                request.Method = "POST";
                request.KeepAlive = false;
                request.Proxy = null;
                request.ContentType = "application/x-www-form-urlencoded";
                request.CookieContainer = cookies;
                request.ContentLength = postData.Length;
                request.Timeout = 1000 * 10000;

                System.IO.Stream outputStream = request.GetRequestStream();
                outputStream.Write(postData, 0, postData.Length);
                outputStream.Close();

                //Get the new response
                response = request.GetResponse() as HttpWebResponse;
                responseStream = response.GetResponseStream();
                reader = new System.IO.StreamReader(responseStream);
                srcString = reader.ReadToEnd();
                return srcString;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: /n/r" + ex.Message);
                throw ex;
            }
        }

        public static CookieCollection FetchCookie(string posturl)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(posturl);
            request.Timeout = 9999999;
            request.CookieContainer = new CookieContainer();

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            return response.Cookies;
        }

        public static bool FileDelete(string path)
        {
            //first set the File\'s ReadOnly to 0
            //if EXP, restore its Attributes
            System.IO.FileInfo file = new System.IO.FileInfo(path);
            System.IO.FileAttributes att = 0;
            bool attModified = false;
            try
            {
                //### ATT_GETnSET
                att = file.Attributes;
                file.Attributes &= (~System.IO.FileAttributes.ReadOnly);
                attModified = true;
                file.Delete();
            }
            catch (Exception e)
            {
                if (attModified)
                    file.Attributes = att;
                return false;
            }
            return true;
        }

        public static void CleanCookie()
        {
            string[] theFiles = System.IO.Directory.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Cookies), "*", System.IO.SearchOption.AllDirectories);
            foreach (string s in theFiles)
                FileDelete(s);

            RunCmd("RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2");
        }

        static void RunCmd(string cmd)
        {
            System.Diagnostics.Process.Start("cmd.exe", "/c" + cmd);
        }

        #region ----------- Get GlobalFGAOrderData Report from Website -----------
        //public static List<string> GetGlobalFGAOrderData(string posturl, string postdata, CookieContainer cookie)
        public static List<string> GetGlobalFGAOrderData(string posturl, string postdata)
        {
            List<string> responselist = new List<string>();

            HttpWebRequest request = WebRequest.Create(posturl) as HttpWebRequest;
            HttpWebResponse response = null;
            ASCIIEncoding encoding = new ASCIIEncoding();

            byte[] b = encoding.GetBytes(postdata);
            request.UserAgent = "Mozilla/4.0";
            request.Method = "POST";
            request.Timeout = 9999999;

            //request.CookieContainer = cookie;
            request.ContentLength = b.Length;

            using (Stream stream = request.GetRequestStream())
            {
                stream.Write(b, 0, b.Length);
            }

            try
            {
                using (response = request.GetResponse() as HttpWebResponse)
                {
                    using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                    {
                        //if (response.Cookies.Count > 0)
                        //    cookie.Add(response.Cookies);
                        //responselist.Add(cookie);
                        responselist.Add(reader.ReadToEnd());
                    }
                }
            }
            catch (WebException ex)
            {
                WebResponse wr = ex.Response;
                using (Stream st = wr.GetResponseStream())
                {
                    using (StreamReader sr = new StreamReader(st, System.Text.Encoding.Default))
                    {
                        responselist.Add(sr.ReadToEnd());
                    }
                }
            }
            catch (Exception ex)
            {
                responselist.Add("Exception: /n/r" + ex.Message);
            }

            return responselist;
        }
        #endregion

        #region ---------- Get Cookie from local ----------
        public static void GetCookie(string url)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Timeout = 9999999;
            request.CookieContainer = new CookieContainer();

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();            

            // Print the properties of each cookie.
            foreach (Cookie cook in response.Cookies)
            {
                Console.WriteLine("Cookie:");
                FileUtility.SaveFile("Cookie.txt", "Cookie:");

                Console.WriteLine("{0} = {1}", cook.Name, cook.Value);
                FileUtility.SaveFile("Cookie.txt", string.Format("{0} = {1}", cook.Name, cook.Value));
                
                Console.WriteLine("Domain: {0}", cook.Domain);
                FileUtility.SaveFile("Cookie.txt", string.Format("Domain: {0}", cook.Domain));

                Console.WriteLine("Path: {0}", cook.Path);
                FileUtility.SaveFile("Cookie.txt", string.Format("Path: {0}", cook.Path));

                Console.WriteLine("Port: {0}", cook.Port);
                FileUtility.SaveFile("Cookie.txt", string.Format("Port: {0}", cook.Port));

                Console.WriteLine("Secure: {0}", cook.Secure);
                FileUtility.SaveFile("Cookie.txt", string.Format("Secure: {0}", cook.Secure));

                Console.WriteLine("When issued: {0}", cook.TimeStamp);
                FileUtility.SaveFile("Cookie.txt", string.Format("When issued: {0}", cook.TimeStamp));

                Console.WriteLine("Expires: {0} (expired? {1})", cook.Expires, cook.Expired);
                FileUtility.SaveFile("Cookie.txt", string.Format("Expires: {0} (expired? {1})", cook.Expires, cook.Expired));

                Console.WriteLine("Don't save: {0}", cook.Discard);
                FileUtility.SaveFile("Cookie.txt", string.Format("Don't save: {0}", cook.Discard));

                Console.WriteLine("Comment: {0}", cook.Comment);
                FileUtility.SaveFile("Cookie.txt", string.Format("Comment: {0}", cook.Comment));

                Console.WriteLine("Uri for comments: {0}", cook.CommentUri);
                FileUtility.SaveFile("Cookie.txt", string.Format("Uri for comments: {0}", cook.CommentUri));

                Console.WriteLine("Version: RFC {0}", cook.Version == 1 ? "2109" : "2965");
                FileUtility.SaveFile("Cookie.txt", string.Format("Version: RFC {0}", cook.Version == 1 ? "2109" : "2965"));
                
                // Show the string representation of the cookie.
                Console.WriteLine("String: {0}", cook.ToString());
                FileUtility.SaveFile("Cookie.txt", string.Format("String: {0}", cook.ToString()));
            }
        }
        #endregion

        public static void ExportTableToExcel(string fullfilename, string htmltext)
        {
            try
            {
                List<string> columnnamelist = new List<string>();
                List<string> celldatalist = new List<string>();

                HtmlDocument htmldoc = new HtmlDocument();
                htmldoc.LoadHtml(htmltext);

                HtmlNodeCollection keynodes = htmldoc.DocumentNode.SelectNodes("//table[@id='GridView1']/tr/th");
                if (keynodes.Count > 0)
                {
                    foreach (HtmlNode keynode in keynodes)
                    {
                        columnnamelist.Add(keynode.InnerText);
                    }
                }

                keynodes = htmldoc.DocumentNode.SelectNodes("//table[@id='GridView1']/tr/td");
                if (keynodes.Count > 0)
                {
                    foreach (HtmlNode keynode in keynodes)
                    {
                        if (keynode.InnerText.Equals("&nbsp;"))
                            celldatalist.Add("");
                        else
                            celldatalist.Add(keynode.InnerText);
                    }
                }

                DataTable datatable = CreateDataTable(columnnamelist, celldatalist);
                ExcelFileUtility.SaveExcelFile(fullfilename, datatable);
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }
        }

        public static DataTable CreateDataTable(List<string> columnnamelist, List<string> celldata)
        {
            int count = 0;
            DataTable datatable = new DataTable();
            DataRow dr = datatable.NewRow();

            try
            {
                for (int index = 0; index < columnnamelist.Count; index++)
                {
                    DataColumn dc = new DataColumn();
                    dc.ColumnName = columnnamelist[index];
                    dc.DataType = System.Type.GetType("System.String");
                    datatable.Columns.Add(dc);
                }

                for (int indey = 0; indey < celldata.Count; indey++)
                {
                    dr[count] = celldata[indey];
                    count++;

                    if (count >= columnnamelist.Count)
                    {
                        datatable.Rows.Add(dr);
                        count = 0;
                        dr = datatable.NewRow();
                    }
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }

            return datatable;
        }

        public static DataTable FetchFGAPOReport(string filename)
        {
            DataTable datatable = new DataTable();
            string newfgapofilename = ExcelFileUtility.SaveAsStandardFileFormat(filename);
            string sheetname = GetSheetName(newfgapofilename);

            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(newfgapofilename))
                {
                    string columnnames = ConfigFileUtility.GetValue("Column_Names");
                    datatable = dao.ReadExcelFile(sheetname, columnnames).Tables[0];
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }

            return datatable;
        }

        public static string GetSheetName(string fullfilename)
        {
            DataTable sheetnameinfile = null;
            DataTable datatable = new DataTable();
            List<string> sheetnamelist = new List<string>();

            try
            {
                using (ExcelAccessDAO dao = new ExcelAccessDAO(fullfilename))
                {
                    sheetnameinfile = dao.GetExcelSheetName();
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("{0}, {1}", ex.Message, ex.StackTrace));
                throw;
            }

            string sheetname = string.Empty;
            for (int indey = 0; indey < sheetnameinfile.Rows.Count; indey++)
            {
                if (sheetnameinfile.Rows[indey]["TABLE_NAME"].ToString().ToUpper().Contains("FGA"))   //Filter virtual sheet
                    sheetname = sheetnameinfile.Rows[indey]["TABLE_NAME"].ToString();
            }

            return sheetname;
        }
    }
}
