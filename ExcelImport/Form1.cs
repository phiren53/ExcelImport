using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
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

                List<string> urls = new List<string>();
                DataTable dtOuptpuExcel = new DataTable();
                var urllistwithBody = new List<KeyValuePair<string, string[]>>();

                ReadExcelWithURLExtraction(out dtOuptpuExcel, out urls, out urllistwithBody);

                GenerateCountReport(urls);

                lblStatus.Text = "Main Report reletad data generation is in progress...";
                //string jsonoutput = DataTableToJSONWithJSONNet(dtOuptpuExcel);
                //WriteDataTableToExcel(dtOuptpuExcel, "Report", @"C:\Users\Pragma Infotech\Desktop\Hiren\Report.xlsx");
                lblStatus.Text = "Generating main report excel...";
                GenerateExcel(dtOuptpuExcel, @"C:\Excel\Report.xlsx");

                lblStatus.Visible = false;

                MessageBox.Show("Report generated successfully.");

                progressBar1.Visible = false;

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        private void ReadExcelWithURLExtraction(out DataTable dtOuptpuExcel, out List<string> urls, out List<KeyValuePair<string, string[]>> urllistwithBody)
        {
            progressBar1.Visible = true;
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;
            urllistwithBody = new List<KeyValuePair<string, string[]>>();
            urls = new List<string>();

            dtOuptpuExcel = new DataTable();
            dtOuptpuExcel.Columns.Add("News_ID", typeof(string));
            dtOuptpuExcel.Columns.Add("DatePublished", typeof(string));
            dtOuptpuExcel.Columns.Add("Title", typeof(string));
            dtOuptpuExcel.Columns.Add("link", typeof(string));
            dtOuptpuExcel.Columns.Add("Body", typeof(string));
            dtOuptpuExcel.Columns.Add("Is Valid HTML", typeof(string));
            dtOuptpuExcel.Columns.Add("Is Valid URL", typeof(string));
            dtOuptpuExcel.Columns.Add("Is fws.gov", typeof(string));
            dtOuptpuExcel.Columns.Add("Is MailTo", typeof(string));

            string str, coltitle;
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


            progressBar1.Maximum = rw;
            progressBar1.Step = 1;

            lblStatus.Visible = true;
            lblStatus.Text = "Data extraction process is in progress...";
            for (rCnt = 2; rCnt <= rw; rCnt++)
            {
                progressBar1.PerformStep();
                DataRow drCurrentRow = dtOuptpuExcel.NewRow();
                string[] bodywithOtherDet = new string[5];
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    coltitle = (string)(range.Cells[1, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                    str = (string)(range.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2;

                    #region HtmlAgilityPack
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();

                    drCurrentRow[coltitle] = str;
                    bodywithOtherDet[cCnt - 1] = str;
                    if (string.IsNullOrEmpty(str) || coltitle.ToUpper() != "BODY")
                    {
                        drCurrentRow["Is Valid HTML"] = "N/A";
                        drCurrentRow["Is Valid URL"] = "N/A";
                        //drCurrentRow["URL type"] = "N/A";
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
                            int stBracketIndex, endBracketIndex, angularBracketIndex;
                            foreach (var item in links)
                            {
                                string hrefValue = item.Attributes["href"].Value;
                                stBracketIndex = hrefValue.IndexOf("<");
                                endBracketIndex = hrefValue.IndexOf(">");

                                if (stBracketIndex > 0 || endBracketIndex > 0)
                                {
                                    if (stBracketIndex > 0 && endBracketIndex > 0)
                                    {
                                        angularBracketIndex = (stBracketIndex > endBracketIndex) ? endBracketIndex : stBracketIndex;
                                    }
                                    else
                                    {
                                        angularBracketIndex = (stBracketIndex > 0) ? stBracketIndex : endBracketIndex;
                                    }

                                    hrefValue = hrefValue.Substring(0, angularBracketIndex);
                                }

                                urls.Add(hrefValue);

                                urllistwithBody.Add(new KeyValuePair<string, string[]>(hrefValue, bodywithOtherDet));

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

                    // dtOuptpuExcel.Rows.Add(drCurrentRow);

                    #endregion

                }

                dtOuptpuExcel.Rows.Add(drCurrentRow);
            }

            //xlWorkBook.Close(true, null, null);
            xlWorkBook.Close(0);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        public void GenerateCountReport(List<string> urls)
        {

            try
            {

                lblStatus.Text = "Count Report reletad data generation is in progress...";
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
                lblStatus.Text = "Generating Count report excel...";
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
            List<string> docextensions = new List<string> { ".doc", ".docx", ".pdf", ".jpeg", ".jpg", ".txt", ".bmp", ".png", ".mp3", ".mp4", ".ppt", ".mov" };
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
            try
            {


                DataSet dataSet = new DataSet();
                dataSet.Tables.Add(dataTable);

                // create a excel app along side with workbook and worksheet and give a name to it  
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();
                //Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = excelWorkBook.Sheets[1];
                //Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;
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
            catch (Exception ex)
            {

                throw ex;
            }
        }

        private void btnGetJson_Click(object sender, EventArgs e)
        {
            string currentURL = string.Empty;
            try
            {
                DataTable dtExcelData, dtParentChildData = new DataTable();

                dtParentChildData.Columns.Add("URL_Data", typeof(string));
                dtParentChildData.Columns.Add("ID", typeof(int));
                dtParentChildData.Columns.Add("PARENT_ID", typeof(int));
                dtParentChildData.Columns.Add("ORG_PATH", typeof(string));
                dtParentChildData.Columns.Add("URL_COUNT", typeof(string));
                dtParentChildData.Columns.Add("News_ID", typeof(string));
                dtParentChildData.Columns.Add("DatePublished", typeof(string));
                dtParentChildData.Columns.Add("Title", typeof(string));
                dtParentChildData.Columns.Add("link", typeof(string));
                dtParentChildData.Columns.Add("ORG_URL", typeof(string));

                List<string> urls = new List<string>();
                var urllistwithBody = new List<KeyValuePair<string, string[]>>();

                ReadExcelWithURLExtraction(out dtExcelData, out urls, out urllistwithBody);
                string urlDetail = string.Empty, orgpath = string.Empty;
                int idSequence = 0;

                #region to generate json with Body
                //var outp = JsonConvert.SerializeObject(urllistwithBody);

                // Write that JSON to txt file,  
                //System.IO.File.WriteAllText(@"C:\Users\Pragma Infotech\Desktop\Hiren\ExcelImport\ExcelImport\bin\Debug\AppData\" + "output.json", outp);
                #endregion

                var urlsGroups = urls.GroupBy(i => i)
                            .Select(grp => new
                            {
                                URL = grp.Key,
                                TotalCount = grp.Count()
                            })
                            .ToArray();

                List<URLDetail> urldetailList = new List<URLDetail>();
                foreach (var item in urlsGroups)
                {
                    urlDetail = item.URL;
                    try
                    {

                        currentURL = urlDetail;
                        //IList<KeyValuePair<string, string>> listofBody = urllistwithBody.Where(v => v.Key == urlDetail.ToString()).ToList();
                        List<string[]> bodyValues = (from v in urllistwithBody where v.Key == urlDetail.ToString() select v.Value).ToList();
                        if (currentURL.Trim() == "")
                        {
                            continue;
                        }

                        if (urlDetail.Length >= 7 && urlDetail.Substring(0, 7) == "http://")
                        {
                            urlDetail = urlDetail.Remove(0, 7);
                        }
                        else if (urlDetail.Length >= 8 && urlDetail.Substring(0, 8) == "https://")
                        {
                            urlDetail = urlDetail.Remove(0, 8);
                        }

                        string[] strArrURLPaths = urlDetail.Split("/");

                        foreach (var path in strArrURLPaths)
                        {
                            if (string.IsNullOrEmpty(path)) continue;

                            string newPath = path.Replace("'", "");

                            //Find Parent id...
                            string parentPath = string.Join('/', strArrURLPaths.Take(Array.IndexOf(strArrURLPaths, newPath)));
                            orgpath = parentPath + (!string.IsNullOrEmpty(parentPath) ? "/" : "") + newPath;
                            DataRow[] drArrPath = dtParentChildData.Select("ORG_PATH = '" + orgpath.Replace("'", "") + "'");
                            if (drArrPath.Length == 0)
                            {
                                DataRow drNewRow = dtParentChildData.NewRow();
                                //generated new id for path if not exists...
                                idSequence++;

                                drNewRow["URL_Data"] = path;
                                drNewRow["ID"] = idSequence;
                                drNewRow["ORG_PATH"] = orgpath;
                                drNewRow["ORG_URL"] = currentURL;



                                if (!string.IsNullOrEmpty(parentPath))
                                {
                                    DataRow[] drparentRow = dtParentChildData.Select("ORG_PATH = '" + parentPath.Replace("'", "") + "'");
                                    if (drparentRow.Length > 0)
                                    {
                                        drNewRow["PARENT_ID"] = int.Parse(drparentRow[0]["ID"].ToString());
                                    }
                                }
                                else
                                {
                                    drNewRow["PARENT_ID"] = 0;
                                }
                                drNewRow["URL_COUNT"] = item.TotalCount;
                                dtParentChildData.Rows.Add(drNewRow);

                                #region Add in the object
                                urldetailList.Add(new URLDetail()
                                {
                                    URL_Data = path,
                                    ID = idSequence,
                                    ORG_PATH = orgpath,
                                    ORG_URL = currentURL,
                                    PARENT_ID = int.Parse(drNewRow["PARENT_ID"].ToString()),
                                    URL_COUNT = item.TotalCount,
                                    BodyDetail = bodyValues
                                });

                                #endregion

                            }
                            else
                            {
                                //If path exists in tree then nothing to do...
                            }

                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                }
                string output = JsonConvert.SerializeObject(urldetailList);
                //string jsonTreeViewData = DataTableToJSONWithJSONNet(dtParentChildData);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        private void btnNewJson_Click(object sender, EventArgs e)
        {
            string currentURL = string.Empty;
            int? _intNull = null;
            try
            {
                DataTable dtExcelData, dtParentChildData = new DataTable();

                dtParentChildData.Columns.Add("URL_Data", typeof(string));
                dtParentChildData.Columns.Add("ID", typeof(int));
                dtParentChildData.Columns.Add("PARENT_ID", typeof(int));
                dtParentChildData.Columns.Add("ORG_PATH", typeof(string));
                dtParentChildData.Columns.Add("URL_COUNT", typeof(string));
                dtParentChildData.Columns.Add("News_ID", typeof(string));
                dtParentChildData.Columns.Add("DatePublished", typeof(string));
                dtParentChildData.Columns.Add("Title", typeof(string));
                dtParentChildData.Columns.Add("link", typeof(string));
                dtParentChildData.Columns.Add("ORG_URL", typeof(string));

                List<string> urls = new List<string>();
                var urllistwithBody = new List<KeyValuePair<string, string[]>>();

                //ReadExcelData(out dtExcelData, out urls, out urllistwithBody);
                lblStatus.Text = "Datatable for Excel Prepared...";
                string urlDetail = string.Empty, orgpath = string.Empty;
                int idSequence = 0;

                #region to generate json with Body
                //var outp = JsonConvert.SerializeObject(urls);

                // Write that JSON to txt file,  
                //System.IO.File.WriteAllText(@"C:\Users\Pragma Infotech\Desktop\Hiren\ExcelImport\ExcelImport\bin\Debug\AppData\" + "urls.json", outp);
                var urlsjson = System.IO.File.ReadAllText(@"C:\Users\Pragma Infotech\Desktop\Hiren\ExcelImport\ExcelImport\bin\Debug\AppData\" + "urls.json");
                var urllistwithBodyjon = System.IO.File.ReadAllText(@"C:\Users\Pragma Infotech\Desktop\Hiren\ExcelImport\ExcelImport\bin\Debug\AppData\" + "urllistwithBody.json");

                urls = JsonConvert.DeserializeObject<List<string>>(urlsjson);
                urllistwithBody = JsonConvert.DeserializeObject<List<KeyValuePair<string, string[]>>>(urllistwithBodyjon);
                #endregion

                var urlsGroups = urls.GroupBy(i => i)
                            .Select(grp => new
                            {
                                URL = grp.Key,
                                TotalCount = grp.Count()
                            })
                            .ToArray();

                List<URLDetail> urldetailList = new List<URLDetail>();
                lblStatus.Text = "Generating Parent-Child Data from URLs...";
                foreach (var item in urlsGroups)
                {
                    urlDetail = item.URL;
                    try
                    {

                        currentURL = urlDetail;
                        if (currentURL == "http://www.pub.whitehouse.gov/uri-res/I2R?urn:pdi://oma.eop.gov")
                        {


                        }
                        //IList<KeyValuePair<string, string>> listofBody = urllistwithBody.Where(v => v.Key == urlDetail.ToString()).ToList();
                        List<string[]> bodyValues = (from v in urllistwithBody where v.Key == urlDetail.ToString() select v.Value).ToList();
                        if (currentURL.Trim() == "")
                        {
                            continue;
                        }

                        if (urlDetail.Length >= 7 && urlDetail.Substring(0, 7) == "http://")
                        {
                            urlDetail = urlDetail.Remove(0, 7);
                        }
                        else if (urlDetail.Length >= 8 && urlDetail.Substring(0, 8) == "https://")
                        {
                            urlDetail = urlDetail.Remove(0, 8);
                        }

                        string[] strArrURLPaths = urlDetail.Split("/");

                        foreach (var path in strArrURLPaths)
                        {
                            if (string.IsNullOrEmpty(path)) continue;

                            string newPath = path.Replace("'", "");

                            //Find Parent id...
                            string parentPath = string.Join('/', strArrURLPaths.Take(Array.IndexOf(strArrURLPaths, newPath)));
                            orgpath = parentPath + (!string.IsNullOrEmpty(parentPath) ? "/" : "") + newPath;
                            DataRow[] drArrPath = dtParentChildData.Select("ORG_PATH = '" + orgpath.Replace("'", "") + "'");
                            if (drArrPath.Length == 0)
                            {
                                DataRow drNewRow = dtParentChildData.NewRow();
                                //generated new id for path if not exists...
                                idSequence++;

                                drNewRow["URL_Data"] = path;
                                drNewRow["ID"] = idSequence;
                                drNewRow["ORG_PATH"] = orgpath;
                                drNewRow["ORG_URL"] = currentURL;



                                if (!string.IsNullOrEmpty(parentPath))
                                {
                                    DataRow[] drparentRow = dtParentChildData.Select("ORG_PATH = '" + parentPath.Replace("'", "") + "'");
                                    if (drparentRow.Length > 0)
                                    {
                                        drNewRow["PARENT_ID"] = int.Parse(drparentRow[0]["ID"].ToString());
                                    }
                                    else
                                    {
                                    }
                                }
                                else
                                {
                                    drNewRow["PARENT_ID"] = 0;
                                }
                                drNewRow["URL_COUNT"] = item.TotalCount;
                                dtParentChildData.Rows.Add(drNewRow);

                                #region Add in the object
                                foreach (var bodyval in bodyValues)
                                {
                                    urldetailList.Add(new URLDetail()
                                    {
                                        URL_Data = path,
                                        ID = idSequence,
                                        ORG_PATH = orgpath,
                                        ORG_URL = currentURL,
                                        PARENT_ID = string.IsNullOrEmpty(drNewRow["PARENT_ID"].ToString()) ? _intNull : int.Parse(drNewRow["PARENT_ID"].ToString()),
                                        URL_COUNT = item.TotalCount,
                                        Body = bodyval
                                    });

                                }


                                #endregion

                            }
                            else
                            {
                                //If path exists in tree then nothing to do...
                            }

                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                }
                lblStatus.Text = "Generating Json...";
                string output = JsonConvert.SerializeObject(urldetailList);
                System.IO.File.WriteAllText(@"C:\Users\Pragma Infotech\Desktop\Hiren\ExcelImport\ExcelImport\bin\Debug\AppData\" + "urldetailList.json", output);
                //string jsonTreeViewData = DataTableToJSONWithJSONNet(dtParentChildData);
                lblStatus.Text = "JSON Generated...";
            }
            catch (Exception ex)
            {
                string s = currentURL;
                throw ex;
            }
        }

        private void ReadExcelData(out DataTable dtOuptpuExcel, out List<string> urls, out List<KeyValuePair<string, string[]>> urllistwithBody)
        {
            try
            {


                progressBar1.Visible = true;
                Microsoft.Office.Interop.Excel.Application xlApp;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                Microsoft.Office.Interop.Excel.Range range;
                urllistwithBody = new List<KeyValuePair<string, string[]>>();
                urls = new List<string>();

                dtOuptpuExcel = new DataTable();
                dtOuptpuExcel.Columns.Add("News_ID", typeof(string));
                dtOuptpuExcel.Columns.Add("DatePublished", typeof(string));
                dtOuptpuExcel.Columns.Add("Title", typeof(string));
                dtOuptpuExcel.Columns.Add("link", typeof(string));
                dtOuptpuExcel.Columns.Add("Body", typeof(string));
                //dtOuptpuExcel.Columns.Add("Is Valid HTML", typeof(string));
                //dtOuptpuExcel.Columns.Add("Is Valid URL", typeof(string));
                //dtOuptpuExcel.Columns.Add("Is fws.gov", typeof(string));
                //dtOuptpuExcel.Columns.Add("Is MailTo", typeof(string));

                string str, coltitle;
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


                progressBar1.Maximum = rw;
                progressBar1.Step = 1;

                lblStatus.Visible = true;
                lblStatus.Text = "Data extraction process is in progress...";
                for (rCnt = 2; rCnt <= rw; rCnt++)
                {
                    progressBar1.PerformStep();
                    DataRow drCurrentRow = dtOuptpuExcel.NewRow();
                    string[] bodywithOtherDet = new string[5];
                    for (cCnt = 1; cCnt <= cl; cCnt++)
                    {
                        coltitle = (string)(range.Cells[1, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                        str = (string)(range.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2;

                        #region HtmlAgilityPack
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();

                        drCurrentRow[coltitle] = str;
                        bodywithOtherDet[cCnt - 1] = str;
                        if (string.IsNullOrEmpty(str) || coltitle.ToUpper() != "BODY")
                        {
                            //drCurrentRow["Is Valid HTML"] = "N/A";
                            //drCurrentRow["Is Valid URL"] = "N/A";
                            //drCurrentRow["URL type"] = "N/A";
                            continue;
                        }

                        doc.LoadHtml(str);

                        if (doc.ParseErrors.Count() > 0)
                        {
                            //Invalid HTML
                            //drCurrentRow["Is Valid HTML"] = "Invalid";
                            //drCurrentRow["Is Valid URL"] = "N/A";
                            //drCurrentRow["URL type"] = "N/A";
                        }
                        else
                        {
                            //Valid
                            //drCurrentRow["Is Valid HTML"] = "Valid";

                            if (doc.DocumentNode.Descendants("a").Count() > 0)
                            {

                                IEnumerable<HtmlNode> links = doc.DocumentNode.Descendants("a")
                                                           .Where(x => x.Attributes["href"] != null
                                                            && x.Attributes["href"].Value != null);

                                //bool isInValidURL = false;
                                //bool result = false;
                                string urltype = string.Empty;
                                int stBracketIndex, endBracketIndex, angularBracketIndex;
                                foreach (var item in links)
                                {
                                    string hrefValue = item.Attributes["href"].Value;
                                    stBracketIndex = hrefValue.IndexOf("<");
                                    endBracketIndex = hrefValue.IndexOf(">");

                                    if (stBracketIndex > 0 || endBracketIndex > 0)
                                    {
                                        if (stBracketIndex > 0 && endBracketIndex > 0)
                                        {
                                            angularBracketIndex = (stBracketIndex > endBracketIndex) ? endBracketIndex : stBracketIndex;
                                        }
                                        else
                                        {
                                            angularBracketIndex = (stBracketIndex > 0) ? stBracketIndex : endBracketIndex;
                                        }

                                        hrefValue = hrefValue.Substring(0, angularBracketIndex);
                                    }

                                    urls.Add(hrefValue);

                                    urllistwithBody.Add(new KeyValuePair<string, string[]>(hrefValue, bodywithOtherDet));

                                    //Uri uriResult;
                                    //result = Uri.TryCreate(hrefValue, UriKind.Absolute, out uriResult)
                                    //    && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);


                                    //if (!result)
                                    //{
                                    //    isInValidURL = true;
                                    //}
                                    //else
                                    //{
                                    //    if (hrefValue.Contains("www.fws.gov"))
                                    //    {
                                    //        drCurrentRow["Is fws.gov"] = "Yes";
                                    //    }
                                    //}

                                    //if (hrefValue.Contains("mailto:"))
                                    //{
                                    //    drCurrentRow["Is MailTo"] = "Yes";
                                    //}
                                }

                                //drCurrentRow["Is Valid URL"] = isInValidURL ? "InValid" : (result ? "Valid" : "InValid");

                            }
                        }

                        // dtOuptpuExcel.Rows.Add(drCurrentRow);

                        #endregion

                    }

                    dtOuptpuExcel.Rows.Add(drCurrentRow);
                }
                progressBar1.Maximum = rw;
                //xlWorkBook.Close(true, null, null);
                xlWorkBook.Close(0);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void btnURLReplace_Click(object sender, EventArgs e)
        {
            try
            {
                UtilityMethod utilityMethod = new UtilityMethod();
                string path = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                DataTable dtMappingData = utilityMethod.ReadExcel(path + "\\News Test Mappings.xlsx", "xlsx", false);

                #region Find and Replace in JSON
                var urlsjson = System.IO.File.ReadAllText(@"C:\Users\Pragma Infotech\Desktop\Hiren\ExcelImport\ExcelImport\bin\Debug\AppData\" + "urldetailList.json");

                List<URLDetail> urlList = JsonConvert.DeserializeObject<List<URLDetail>>(urlsjson);
                DataTable dtUrlData = utilityMethod.ReadExcel(path + "\\News_29072021.xlsx", "xlsx", true);
                //dtUrlData.Columns.Add("New Body", typeof(string));

                foreach (DataRow mappingItem in dtMappingData.Rows)
                {


                    var matchedURLs = from url in urlList
                                      where url.Body[4].ToLower().Contains(mappingItem["URL"].ToString().ToLower())
                                      select url;

                    foreach (var item in matchedURLs)
                    {
                        //Console.WriteLine($"{item.Name,-15}{item.Score}");
                        item.Body[4] = item.Body[4].ToLower().Replace(mappingItem["URL"].ToString(), mappingItem["Replacement URL"].ToString());
                        item.Body[3] = item.Body[3].ToLower().Replace(mappingItem["URL"].ToString(), mappingItem["Replacement URL"].ToString());
                        item.ORG_PATH = item.ORG_PATH.ToLower().Replace(mappingItem["URL"].ToString(), mappingItem["Replacement URL"].ToString());
                        item.ORG_URL = item.ORG_URL.ToLower().Replace(mappingItem["URL"].ToString(), mappingItem["Replacement URL"].ToString());
                        item.URL_Data = item.ORG_URL.ToLower().Replace(mappingItem["URL"].ToString(), mappingItem["Replacement URL"].ToString());
                    }

                    //To update Link in Body object
                    matchedURLs = from url in urlList
                                  where url.Body[3].ToLower().Contains(mappingItem["URL"].ToString().ToLower())
                                  select url;

                    foreach (var item in matchedURLs)
                    {
                        //Console.WriteLine($"{item.Name,-15}{item.Score}");
                        item.Body[4] = item.Body[4].ToLower().Replace(mappingItem["URL"].ToString(), mappingItem["Replacement URL"].ToString());
                        item.Body[3] = item.Body[3].ToLower().Replace(mappingItem["URL"].ToString(), mappingItem["Replacement URL"].ToString());
                        item.ORG_PATH = item.ORG_PATH.ToLower().Replace(mappingItem["URL"].ToString(), mappingItem["Replacement URL"].ToString());
                        item.ORG_URL = item.ORG_URL.ToLower().Replace(mappingItem["URL"].ToString(), mappingItem["Replacement URL"].ToString());
                        item.URL_Data = item.ORG_URL.ToLower().Replace(mappingItem["URL"].ToString(), mappingItem["Replacement URL"].ToString());
                    }

                    #region Add Replaced Body value in the Excel

                    var drArrmatchedURLs = from drurl in dtUrlData.AsEnumerable()
                                           where drurl["newbody"].ToString().Contains(mappingItem["URL"].ToString().ToLower())
                                           select drurl;

                    foreach (var drURL in drArrmatchedURLs)
                    {
                        dtUrlData.Rows[dtUrlData.Rows.IndexOf(drURL)]["newbody"] = drURL["newbody"].ToString().Replace(mappingItem["URL"].ToString(), mappingItem["Replacement URL"].ToString());
                    }
                    #endregion

                }
                string output = JsonConvert.SerializeObject(urlList);
                System.IO.File.WriteAllText(@"C:\Users\Pragma Infotech\Desktop\Hiren\ExcelImport\ExcelImport\bin\Debug\AppData\" + "newreplacedurl.json", output);
                //GenerateExcel(dtUrlData, path + "\\NewBodyWithReplacedURL.xlsx");
                dtUrlData.Columns["newbody"].ColumnName = "body";
                string output1 = utilityMethod.DataTableToJSONWithJSONNet(dtUrlData);
                
                System.IO.File.WriteAllText(@"C:\Users\Pragma Infotech\Desktop\Hiren\ExcelImport\ExcelImport\bin\Debug\AppData\" + "newreplacedurlwithdefaultText.json", output1);
                #endregion

                MessageBox.Show("Process completed.");

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void btnJSON_Click(object sender, EventArgs e)
        {
            try
            {

                UtilityMethod utilityMethod = new UtilityMethod();
                DataTable dtExcelData = utilityMethod.ReadExcel(txtPath.Text, "xlsx", true);
                string exceljson = utilityMethod.DataTableToJSONWithJSONNet(dtExcelData);
                string path = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                System.IO.File.WriteAllText(path + "\\pdfbasedata.json", exceljson);


            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void btnDownloadFile_Click(object sender, EventArgs e)
        {
            try
            {
                UtilityMethod utilityMethod = new UtilityMethod();
                DataTable dtExcelData = utilityMethod.ReadExcel(txtPath.Text, "xlsx", false);//Take Count Report File to process
                DataRow[] dataRows = dtExcelData.Select("DocType = 'Document'");
                StringBuilder sb = new StringBuilder();
                progressBar1.Visible = true;
                progressBar1.Maximum = dataRows.Length;
                progressBar1.Step = 1;
                foreach (DataRow item in dataRows)
                {
                    progressBar1.PerformStep();
                    // A web URL with a file response
                    string myWebUrlFile = item["URL"].ToString();

                    if (myWebUrlFile.Contains("fws.gov"))
                    {
                        //string[] pathname = myWebUrlFile.Split("/");
                        string filename = utilityMethod.GetURLFilename(myWebUrlFile);// pathname[pathname.Length - 1].ToString();
                                                                                     
                        string myLocalFilePath = "C:/Excel/DownloadFiles/" + filename;// Local path where the file will be saved

                        if (!filename.Contains(".pdf")) continue;

                        try
                        {
                            using (var client = new WebClient())
                            {
                                client.DownloadFile(myWebUrlFile, myLocalFilePath);
                            }
                        }
                        catch (WebException we)
                        {
                            sb.AppendLine(myWebUrlFile);
                            
                        }
                    }

                }

                using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:/Excel/DownloadFiles/ErrorPaths.txt"))
                {
                    file.WriteLine(sb.ToString()); // "sb" is the StringBuilder
                }
                progressBar1.Visible = false;
                MessageBox.Show("Downloading completed.");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

