using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Global_FGA_Order_Report
{
    public class MiscUtility
    {
        public static void LogHistory(string text)
        {
            string logfilename = "History.log";
            FileUtility.SaveFile(logfilename, string.Format("[{0}] - {1}", DateTime.Now.ToString(), text));
        }
    }
}
