using System;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Global_FGA_Order_Report
{
    public class ExcelAccessDAO : ExcelFileAccess
    {
        public ExcelAccessDAO()
        {
        }

        public ExcelAccessDAO(string filename)
            : base(filename)
        {
            try
            {
                if (base.connection.State == System.Data.ConnectionState.Closed)
                {
                    base.connection.Open();
                }
            }
            catch
            {
                throw;
            }
        }

        public ExcelAccessDAO(string filename, bool isfield)
            : base(filename, isfield)
        {
            try
            {
                if (base.connection.State == System.Data.ConnectionState.Closed)
                {
                    base.connection.Open();
                }
            }
            catch
            {
                throw;
            }
        }

        public DataTable GetExcelSheetName()
        {
            DataTable schematable = base.connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new Object[] { null, null, null, "TABLE" });
            return schematable;
        }

        public DataSet ReadExcelFile(string sheetname)
        {
            string sqlString = string.Empty;

            if (sheetname.Contains("$"))
                sqlString = String.Format("SELECT * FROM [{0}];", sheetname);
            else
                sqlString = String.Format("SELECT * FROM [{0}$];", sheetname);

            return this.ExecuteQuery(sqlString);
        }

        public DataSet ReadExcelFile(string sheetname, string fieldname)
        {
            string sqlString = string.Empty;

            if (sheetname.Contains("$"))
                sqlString = String.Format("SELECT {0} FROM [{1}];", fieldname, sheetname);
            else
                sqlString = String.Format("SELECT {0} FROM [{1}$];", fieldname, sheetname);

            return this.ExecuteQuery(sqlString);
        }
    }
}
