using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using HtmlAgilityPack;
using Newtonsoft.Json;

namespace ExcelImport
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            txtPath.Text = @"C:\Excel\Sample.xlsx";
        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {


                progressBar1.Visible = true;
                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                Microsoft.Office.Interop.Excel.Range range;

                string str;
                int rCnt;
                int cCnt;
                int rw = 0;
                int cl = 0;

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                //xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\Pragma Infotech\Downloads\Hiren Download\FWS Website Main.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                //xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\Pragma Infotech\Downloads\Hiren Download\ReadExcelToGrid\ReadExcelToGrid\App_Data\Sample.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);


                if (String.IsNullOrEmpty(txtPath.Text.ToString()))
                {
                    MessageBox.Show("Enter Excel File Path.");
                    return;
                }

                xlWorkBook = xlApp.Workbooks.Open(txtPath.Text.ToString(), 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                range = xlWorkSheet.UsedRange;
                rw = range.Rows.Count;
                cl = range.Columns.Count;
                DataTable dtOuptpuExcel = new DataTable();
                dtOuptpuExcel.Columns.Add("Body", typeof(string));
                dtOuptpuExcel.Columns.Add("Is Valid HTML", typeof(string));
                dtOuptpuExcel.Columns.Add("Is Valid URL", typeof(string));
                //dtOuptpuExcel.Columns.Add("URL type", typeof(string));
                dtOuptpuExcel.Columns.Add("Is fws.gov", typeof(string));
                dtOuptpuExcel.Columns.Add("Is MailTo", typeof(string));

                progressBar1.Maximum = rw;
                progressBar1.Step = 1;
                List<string> urls = new List<string>();
                for (rCnt = 2; rCnt <= rw; rCnt++)
                {
                    progressBar1.PerformStep();
                    for (cCnt = 1; cCnt <= cl; cCnt++)
                    {
                        str = (string)(range.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2;

                        #region HtmlAgilityPack
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        DataRow drCurrentRow = dtOuptpuExcel.NewRow();
                        drCurrentRow["Body"] = str;

                        if (string.IsNullOrEmpty(str))
                        {
                            drCurrentRow["Is Valid HTML"] = "N/A";
                            drCurrentRow["Is Valid URL"] = "N/A";
                            //drCurrentRow["URL type"] = "N/A";
                            dtOuptpuExcel.Rows.Add(drCurrentRow);
                            continue;
                        }

                        doc.LoadHtml(str);

                        if (doc.ParseErrors.Count() > 0)
                        {
                            //Invalid HTML
                            drCurrentRow["Is Valid HTML"] = "Invalid";
                            drCurrentRow["Is Valid URL"] = "N/A";
                            //drCurrentRow["URL type"] = "N/A";
                        }
                        else
                        {
                            //Valid
                            drCurrentRow["Is Valid HTML"] = "Valid";

                            if (doc.DocumentNode.Descendants("a").Count() > 0)
                            {

                                IEnumerable<HtmlNode> links = doc.DocumentNode.Descendants("a")
                                                           .Where(x => x.Attributes["href"] != null
                                                            && x.Attributes["href"].Value != null);
                                //string urlOutput = "";
                                bool isInValidURL = false;
                                bool result = false;
                                string urltype = string.Empty;
                                foreach (var item in links)
                                {
                                    string hrefValue = item.Attributes["href"].Value;
                                    
                                    if (hrefValue.IndexOf("</a>") > 0)
                                    {
                                        hrefValue = hrefValue.Substring(0, hrefValue.IndexOf("</a>"));
                                    }
                                    else if (hrefValue.IndexOf("</A>") > 0)
                                    {
                                        hrefValue = hrefValue.Substring(0, hrefValue.IndexOf("</A>"));
                                    }


                                    urls.Add(hrefValue);



                                    Uri uriResult;
                                    result = Uri.TryCreate(hrefValue, UriKind.Absolute, out uriResult)
                                        && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);


                                    if (!result)
                                    {
                                        isInValidURL = true;
                                    }
                                    else
                                    {
                                        if (hrefValue.Contains("www.fws.gov"))
                                        {
                                            drCurrentRow["Is fws.gov"] = "Yes";
                                        }
                                    }

                                    if (hrefValue.Contains("mailto:"))
                                    {
                                        drCurrentRow["Is MailTo"] = "Yes";
                                    }
                                }

                                drCurrentRow["Is Valid URL"] = isInValidURL ? "InValid" : (result ? "Valid" : "InValid");

                            }
                        }

                        dtOuptpuExcel.Rows.Add(drCurrentRow);

                        #endregion

                    }
                }


                GenerateCountReport(urls);

                //string jsonoutput = DataTableToJSONWithJSONNet(dtOuptpuExcel);
                //WriteDataTableToExcel(dtOuptpuExcel, "Report", @"C:\Users\Pragma Infotech\Desktop\Hiren\Report.xlsx");
                GenerateExcel(dtOuptpuExcel, @"C:\Excel\Report.xlsx");


                //xlWorkBook.Close(true, null, null);
                xlWorkBook.Close(0);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);


                MessageBox.Show("Report generated successfully.");

                progressBar1.Visible = false;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }


        public void GenerateCountReport(List<string> urls)
        {

            try
            {


                var urlsGroups = urls.GroupBy(i => i)
                            .Select(grp => new
                            {
                                URL = grp.Key,
                                TotalCount = grp.Count()
                            })
                            .ToArray();

                DataTable dtOuptputExcel = new DataTable();
                dtOuptputExcel.Columns.Add("URL", typeof(string));
                dtOuptputExcel.Columns.Add("Total Count", typeof(string));
                dtOuptputExcel.Columns.Add("HTTP type", typeof(string));
                dtOuptputExcel.Columns.Add("URL_wo_Scheme", typeof(string));
                dtOuptputExcel.Columns.Add("Extension", typeof(string));
                dtOuptputExcel.Columns.Add("DocType", typeof(string));
                foreach (var item in urlsGroups)
                {
                    Uri uriResult;
                    bool result = Uri.TryCreate(item.URL.ToString(), UriKind.Absolute, out uriResult)
                        && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);



                    DataRow dr = dtOuptputExcel.NewRow();
                    dr["URL"] = item.URL.ToString();
                    dr["Total Count"] = item.TotalCount.ToString();
                    if (result)
                    {
                        dr["HTTP type"] = uriResult.Scheme.ToString();
                        SetHTTPType(dtOuptputExcel, uriResult, dr);
                        SetDocType(item.URL.ToString(), dtOuptputExcel, dr);
                    }


                    dtOuptputExcel.Rows.Add(dr);
                }
                dtOuptputExcel.Columns.Remove("URL_wo_Scheme");
                dtOuptputExcel.Columns.Remove("Extension");
                //string jsonoutput = DataTableToJSONWithJSONNet(dtOuptputExcel);
                GenerateExcel(dtOuptputExcel, @"C:\Excel\CountReport.xlsx");
                //WriteDataTableToExcel(dtOuptputExcel, "Count Report", @"C:\Users\Pragma Infotech\Desktop\Hiren\CountReport.xlsx");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void SetHTTPType(DataTable dtOuptputExcel, Uri uriResult, DataRow dr)
        {
            #region find similar URL without http/https
            string uriWithoutScheme = uriResult.Host + uriResult.PathAndQuery + uriResult.Fragment;
            dr["URL_wo_Scheme"] = uriWithoutScheme;


            DataRow[] drArrSameURL = dtOuptputExcel.Select("URL_wo_Scheme = '" + uriWithoutScheme.Replace("'", "''") + "'");
            if (drArrSameURL.Length > 0)
            {
                foreach (DataRow dataRow in drArrSameURL)
                {
                    if (uriResult.Scheme.ToString() != dataRow["HTTP type"].ToString())
                    {
                        dtOuptputExcel.Rows[dtOuptputExcel.Rows.IndexOf(dataRow)]["HTTP type"] = "BOTH";
                        dr["HTTP type"] = "BOTH";//Change current row value also...
                    }
                }
            }
            #endregion
        }

        private static void SetDocType(string url, DataTable dtOuptputExcel, DataRow dr)
        {
            string pattern = @"\.\w{3,4}($|\?)";
            RegexOptions options = RegexOptions.Multiline;
            List<string> docextensions = new List<string> { ".doc", ".docx", ".pdf", ".jpeg", ".txt", ".bmp", ".png", ".mp3", ".mp4", ".ppt", ".mov" };
            List<string> pageextensions = new List<string> { ".aspx", ".html", ".php", ".htm", ".jsp" };

            foreach (Match m in Regex.Matches(url, pattern, options))
            {
                string extension = m.Value;
                dr["Extension"] = extension;
                if (docextensions.Contains(extension.Replace("?", "")))
                {
                    dr["DocType"] = "Document";
                }
                else if (pageextensions.Contains(extension.Replace("?", "")))
                {
                    dr["DocType"] = "Page";
                }
            }
        }
        public static void GenerateExcel(DataTable dataTable, string path)
        {

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(dataTable);

            // create a excel app along side with workbook and worksheet and give a name to it  
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = excelWorkBook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;
            foreach (DataTable table in dataSet.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name  
                Microsoft.Office.Interop.Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

                // add all the columns  
                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }

                // add all the rows  
                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
            }

            excelWorkBook.SaveAs(path); // -> this will do the custom  
            excelWorkBook.Close(0);
            excelApp.Quit();
        }


        public string DataTableToJSONWithJSONNet(DataTable table)
        {
            string JSONString = string.Empty;
            JSONString = JsonConvert.SerializeObject(table);
            return JSONString;
        }
    }
}

