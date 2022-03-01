Source File : C:\Excel\Excel - Process 2 09-12-2021\News Export 12-07-2021 For Report 1522022.xls


1. Generate Report
Required column in Excel (News_ID,	DatePublished,	Title,	link,	body)
Output : C:\Excel\Report.xlsx
------------------------------------------------

2. Generate URLDetail List JSON
It will extract excel as per above format
and it will generate 2 json file 
   - urls.json   (List of URL found from the body)
   - urllistwithBody.json  (list of URL with respected body)
------------------------------------------------

3. Parent Child URL List JSON
It will take generate Parent- child data from the previous json (urls.json & urllistwithBody.json)
Output File: C:\Users\Pragma Infotech\Desktop\Hiren\ExcelImport\ExcelImport\bin\Debug\AppData\" + "urldetailList.json
------------------------------------------------

4. Find - Replace
It will replace URL base on mapping information provided in source excel.

Source File : C:\\Excel\\Excel - Process 2 09-12-2021\\URL Mapping 1.0.xlsx
              C:\Users\Pragma Infotech\Desktop\Hiren\ExcelImport\ExcelImport\bin\Debug\AppData\" + "urldetailList.json
              C:\\Excel\\Excel - Process 2 09-12-2021\\News Export 12-07-2021 For Replace URL1.xlsx

Output File : C:\Users\Pragma Infotech\Desktop\Hiren\ExcelImport\ExcelImport\bin\Debug\AppData\" + "newreplacedurl.json
              C:\Users\Pragma Infotech\Desktop\Hiren\ExcelImport\ExcelImport\bin\Debug\AppData\" + "newreplacedurlwithdefaultText.json
------------------------------------------------
5. 


------------------------------------------------
6. Generate JSON - WithAndWithout PressRelease
To generate JSON With and Without PressRelease category along with Teaser and Region (Lookup in the excel)
For Region Excel : News Regions Export 12-07-2021.xls
- Add ""Teaser" Column in excel
- In excel sheet name should be "Sheet1"

Input File : C:\Excel\Excel - Process 2 09-12-2021\News Export 12-07-2021 For Replace URL1.xlsx
Output File : C:\\Users\\Pragma Infotech\\Desktop\\Hiren\\ExcelImport\\ExcelImport\\bin\\Debug\\AppData\\PressReleasedata.json
              C:\\Users\\Pragma Infotech\\Desktop\\Hiren\\ExcelImport\\ExcelImport\\bin\\Debug\\AppData\\WithoutPressReleasedata.json

              