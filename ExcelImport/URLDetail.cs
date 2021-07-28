using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelImport
{
    class URLDetail
    {
        public string URL_Data { get; set; }
        public int ID { get; set; }
        public string ORG_PATH { get; set; }
        public string ORG_URL { get; set; }
        public int? PARENT_ID { get; set; }
        public int URL_COUNT { get; set; }
        public List<string[]> BodyDetail { get; set; }
        public string[] Body { get; set; }
        /*
         drNewRow["URL_Data"] = path;
                                drNewRow["ID"] = idSequence;
                                drNewRow["ORG_PATH"] = orgpath;
                                drNewRow["ORG_URL"] = currentURL; 
         */

    }
}
