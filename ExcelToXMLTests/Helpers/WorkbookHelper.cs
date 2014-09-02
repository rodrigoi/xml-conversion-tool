using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;

namespace ExcelToXMLTests.Helpers {
    public sealed class WorkbookHelper {
        private WorkbookHelper() { }

        public static XLWorkbook CreateBasicWorkBook() {
            var workbook = new XLWorkbook();

            workbook.Worksheets.Add("Omniture");
            var omniture = workbook.Worksheet("Omniture");

            AppendOmnitureHeaders(omniture);
            AppendOmnitureStandardPageView(omniture);

            workbook.Worksheets.Add("floodlight");

            return workbook;
        }

        public static void AppendOmnitureHeaders(IXLWorksheet omniture) { 
            var headers = new string[] {
                "Id", 
                "Parent Id", 
                "Site", 
                "Site Section", 
                "Subsection", 
                "Active State", 
                "Object", 
                "Object Description", 
                "Call To Action", 
                "pagename", 
                "event", 
                "prop1", 
                "prop2", 
                "prop3", 
                "prop4", 
                "prop5", 
                "prop6", 
                "prop7", 
                "prop8", 
                "prop9", 
                "prop10", 
                "prop11", 
                "prop12", 
                "prop13", 
                "prop14", 
                "prop15", 
                "prop16", 
                "prop17", 
                "prop18", 
                "prop19", 
                "prop20", 
                "prop21", 
                "prop22", 
                "prop23", 
                "prop24", 
                "prop25", 
                "prop26", 
                "prop27"
            };
            AppendArray(omniture, headers);
        }
        public static void AppendOmnitureStandardPageView(IXLWorksheet omniture) {
            var values = new string[] {
                "t-1",
                "",
                "DUNLOP",
                "HOMEPAGE",
                "HOMEPAGE",
                "INDEX",
                "index",
                "",
                "",
                "DUNLOP | HOMEPAGE | HOMEPAGE | INDEX",
                "",
                "HOMEPAGE",
                "HOMEPAGE",
                "HOMEPAGE | HOMEPAGE",
                "INDEX",
                "HOMEPAGE | INDEX",
                "HOMEPAGE | HOMEPAGE  | INDEX",
                "INDEX",
                "HOMEPAGE | INDEX",
                "HOMEPAGE | INDEX | INDEX",
                "HOMEPAGE | HOMEPAGE | HOMEPAGE | INDEX"
            };
        }
        private static void AppendArray(IXLWorksheet worksheet, string[] headers) {
            for (int index = 0; index < headers.Length; index++) {
                worksheet.Row(1).Cell(index + 1).Value = headers[index];
            }
        }
    }
}