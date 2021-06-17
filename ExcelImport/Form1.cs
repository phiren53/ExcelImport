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
                dtOuptpuExcel.Columns.Add("URL type", typeof(string));

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
                            drCurrentRow["URL type"] = "N/A";
                            dtOuptpuExcel.Rows.Add(drCurrentRow);
                            continue;
                        }

                        doc.LoadHtml(str);

                        if (doc.ParseErrors.Count() > 0)
                        {
                            //Invalid HTML
                            drCurrentRow["Is Valid HTML"] = "Invalid";
                            drCurrentRow["Is Valid URL"] = "N/A";
                            drCurrentRow["URL type"] = "N/A";
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

                                    urls.Add(hrefValue);

                                    Uri uriResult;
                                    result = Uri.TryCreate(hrefValue, UriKind.Absolute, out uriResult)
                                        && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);

                                    if (!result)
                                    {
                                        isInValidURL = true;

                                        //// to define URL type...
                                        //if (string.IsNullOrEmpty(urltype))
                                        //{
                                        //    urltype = uriResult.Scheme;
                                        //}
                                        //else
                                        //{
                                        //    if (urltype != uriResult.Scheme)
                                        //        urltype = "Both";
                                        //}
                                    }

                                    //urlOutput = urlOutput +"\n"+ (hrefValue + " - " + result);
                                    //MessageBox.Show(hrefValue);
                                    //MessageBox.Show(str, hrefValue + " - "  + result);
                                }

                                drCurrentRow["Is Valid URL"] = isInValidURL ? "InValid" : (result ? "Valid" : "InValid");

                                //MessageBox.Show(urlOutput);


                            }
                        }

                        dtOuptpuExcel.Rows.Add(drCurrentRow);



                        #endregion

                    }
                }


                GenerateCountReport(urls);

                string jsonoutput = DataTableToJSONWithJSONNet(dtOuptpuExcel);
                //WriteDataTableToExcel(dtOuptpuExcel, "Report", @"C:\Users\Pragma Infotech\Desktop\Hiren\Report.xlsx");
                GenerateExcel(dtOuptpuExcel, @"C:\Excel\Report.xlsx");


                xlWorkBook.Close(true, null, null);
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

        public bool WriteDataTableToExcel(System.Data.DataTable dataTable, string worksheetName, string saveAsLocation)
        {

            //// Bind table data to Stream Writer to export data to respective folder
            //StreamWriter wr = new StreamWriter(saveAsLocation);
            //// Write Columns to excel file
            //for (int i = 0; i < dataTable.Columns.Count; i++)
            //{
            //    wr.Write(dataTable.Columns[i].ToString().ToUpper() + "\t");
            //}
            //wr.WriteLine();
            ////write rows to excel file
            //for (int i = 0; i < (dataTable.Rows.Count); i++)
            //{
            //    for (int j = 0; j < dataTable.Columns.Count; j++)
            //    {
            //        if (dataTable.Rows[i][j] != null)
            //        {
            //            wr.Write(Convert.ToString(dataTable.Rows[i][j]) + "\t");
            //        }
            //        else
            //        {
            //            wr.Write("\t");
            //        }
            //    }
            //    wr.WriteLine();
            //}
            //wr.Close();
            //return true;


            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet;

            try
            {
                //  get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;

                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Add(Type.Missing);

                // Workk sheet
                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
                excelSheet.Name = worksheetName;

                // loop through each row and add values to our sheet
                int rowcount = 1;

                foreach (DataRow datarow in dataTable.Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= dataTable.Columns.Count; i++)
                    {
                        // on the first iteration we add the column headers
                        if (rowcount == 3)
                        {
                            excelSheet.Cells[2, i] = dataTable.Columns[i - 1].ColumnName;
                        }
                        // Filling the excel file 
                        excelSheet.Cells[rowcount, i] = datarow[i - 1].ToString();
                    }
                }

                //now save the workbook and exit Excel
                excelworkBook.SaveAs(saveAsLocation); ;
                excelworkBook.Close();
                excel.Quit();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            finally
            {
                excelSheet = null;
                excelworkBook = null;
            }
        }


        public void GenerateCountReport(List<string> urls)
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
                }
                dtOuptputExcel.Rows.Add(dr);
            }
            string jsonoutput = DataTableToJSONWithJSONNet(dtOuptputExcel);
            GenerateExcel(dtOuptputExcel, @"C:\Excel\CountReport.xlsx");
            //WriteDataTableToExcel(dtOuptputExcel, "Count Report", @"C:\Users\Pragma Infotech\Desktop\Hiren\CountReport.xlsx");
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
            excelWorkBook.Close();
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

