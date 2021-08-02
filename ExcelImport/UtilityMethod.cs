using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace ExcelImport
{
    class UtilityMethod
    {
        public string DataTableToJSONWithJSONNet(System.Data.DataTable table)
        {
            string JSONString = string.Empty;
            JSONString = JsonConvert.SerializeObject(table);
            return JSONString;
        }

        public System.Data.DataTable ReadExcel(string fileName, string fileExt, bool IswithQuery)
        {
            string conn = string.Empty;
            System.Data.DataTable dtexcel = new System.Data.DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt;
                    //OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select Body,IIF(Body1 IS NULL,'hiren',Body1) AS temp,IIF(Body IS NULL,'IMNULL',MID(Body,LEN(Body) - 3,4)) As bodylen from [Sheet1$]", con); //here we read data from sheet1
                    if (IswithQuery)
                    {
                        oleAdpt = new OleDbDataAdapter("select News_ID, DatePublished, Title, Alias, link, MoreInfo, posted, body, IIF((LEN(body) = 0 AND LEN(link) > 0 AND MID(link,LEN(link) - 3,4) = '.pdf'), 'This archived news release was scanned from a paper copy that may show damage or excessive wear. Some text may be difficult or impossible to read. If you require assistance with the content of this release, please contact us.',body) As newBody, posted_by, Program, Category from [Sheet1$]", con); //here we read data from sheet1
                    }
                    else
                    {
                        oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con);
                    }
                    //IIF(body is not null,body, IIF(MID(link,LEN(link) - 3,4) = '.pdf','IM PDF','NO')) As temp,

                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch(Exception ex) {  }
                
            }
            return dtexcel;
        }

        public System.Data.DataTable ReadExcelSheet(string path)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;

            System.Data.DataTable dtOuptpuExcel = new System.Data.DataTable();

            string str, coltitle;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            //xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\Pragma Infotech\Downloads\Hiren Download\FWS Website Main.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            //xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\Pragma Infotech\Downloads\Hiren Download\ReadExcelToGrid\ReadExcelToGrid\App_Data\Sample.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

            xlWorkBook = xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            for (cCnt = 1; cCnt <= cl; cCnt++)
            {
                coltitle = (string)(range.Cells[1, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                dtOuptpuExcel.Columns.Add(coltitle, typeof(string));
            }

            for (rCnt = 2; rCnt <= rw; rCnt++)
            {
                DataRow drCurrentRow = dtOuptpuExcel.NewRow();
                string[] bodywithOtherDet = new string[5];
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    coltitle = (string)(range.Cells[1, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                    str = (string)(range.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                    drCurrentRow[coltitle] = str;
                }

                dtOuptpuExcel.Rows.Add(drCurrentRow);
            }

            //xlWorkBook.Close(true, null, null);
            xlWorkBook.Close(0);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            return dtOuptpuExcel;
        }

        public string GetURLFilename(string hreflink)
        {
            Uri uri = new Uri(hreflink);

            string filename = System.IO.Path.GetFileName(uri.LocalPath);

            return filename;
        }
    }
}
