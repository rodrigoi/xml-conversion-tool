using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Xml.Linq;
using ClosedXML.Excel;
using ExcelToXML.Model;
using ExcelToXML.Services.Omniture;
using System.IO;

namespace ExcelToXML.Interpreters.Excel {
    [SourceFileExtension(".xlsx")]
    public class ExcelInterpreter : IInterpreter {
        public string FileTimeStamp { get; set; }

        public void CreateTrackingXml(string inputFileName, string trackingOutputFileName, string metadataOutputFileName) {
            Trace.WriteLine("opening source workbook... be patient :)");

            var workbook = new ClosedXML.Excel.XLWorkbook(inputFileName);
            Trace.WriteLine("omniture source workbook opened");

            var worksheet = workbook.Worksheets.Where(x => x.Name.ToLower() == "omniture").FirstOrDefault();

            XElement definitions = new XElement("definitions");

            int maxTrackingId = GetMaxTrackingId(definitions);
            var lastRowUsed = worksheet.LastRowUsed().RowNumber();
            var lastColumnUsed = worksheet.LastColumnUsed().ColumnNumber();

            String trackingId = "";

            for (int rowIndex = 2; rowIndex <= lastRowUsed; rowIndex++) {
                var row = worksheet.Row(rowIndex);
                var element = definitions.Elements("tracking").Where(x => string.Equals(
                    x.Attribute("id").Value, 
                    row.Cell(1).GetValue<string>(), 
                    StringComparison.InvariantCultureIgnoreCase)
                ).SingleOrDefault();

                var tracking = new XElement("tracking");

                var omniture = new Omniture(row);
                omniture.Skip = BuildSkipList(row, lastColumnUsed);

                dynamic propBuilder = new PropBuilder(omniture);
                DetectPatternFor("pageName", omniture, row.Cell(10).GetValue<string>(), propBuilder.pageName);
                DetectPatternFor("prop1", omniture, row.Cell(12).GetValue<string>(), propBuilder.prop1);
                DetectPatternFor("prop2", omniture, row.Cell(13).GetValue<string>(), propBuilder.prop2);
                DetectPatternFor("prop3", omniture, row.Cell(14).GetValue<string>(), propBuilder.prop3);
                DetectPatternFor("prop4", omniture, row.Cell(15).GetValue<string>(), propBuilder.prop4);
                DetectPatternFor("prop5", omniture, row.Cell(16).GetValue<string>(), propBuilder.prop5);
                DetectPatternFor("prop6", omniture, row.Cell(17).GetValue<string>(), propBuilder.prop6);
                DetectPatternFor("prop7", omniture, row.Cell(18).GetValue<string>(), propBuilder.prop7);
                DetectPatternFor("prop8", omniture, row.Cell(19).GetValue<string>(), propBuilder.prop8);
                DetectPatternFor("prop9", omniture, row.Cell(20).GetValue<string>(), propBuilder.prop9);
                DetectPatternFor("prop10", omniture, row.Cell(21).GetValue<string>(), propBuilder.prop10);
                DetectPatternFor("prop11", omniture, row.Cell(22).GetValue<string>(), propBuilder.prop11);
                DetectPatternFor("prop12", omniture, row.Cell(23).GetValue<string>(), propBuilder.prop12);
                DetectPatternFor("prop13", omniture, row.Cell(24).GetValue<string>(), propBuilder.prop13);
                DetectPatternFor("prop14", omniture, row.Cell(25).GetValue<string>(), propBuilder.prop14);
                DetectPatternFor("prop15", omniture, row.Cell(26).GetValue<string>(), propBuilder.prop15);
                DetectPatternFor("prop16", omniture, row.Cell(27).GetValue<string>(), propBuilder.prop16);
                DetectPatternFor("prop17", omniture, row.Cell(28).GetValue<string>(), propBuilder.prop17);
                DetectPatternFor("prop18", omniture, row.Cell(29).GetValue<string>(), propBuilder.prop18);
                DetectPatternFor("prop19", omniture, row.Cell(30).GetValue<string>(), propBuilder.prop19);
                DetectPatternFor("prop20", omniture, row.Cell(31).GetValue<string>(), propBuilder.prop20);
                DetectPatternFor("prop21", omniture, row.Cell(32).GetValue<string>(), propBuilder.prop21);
                DetectPatternFor("prop22", omniture, row.Cell(33).GetValue<string>(), propBuilder.prop22);
                DetectPatternFor("prop23", omniture, row.Cell(34).GetValue<string>(), propBuilder.prop23);
                DetectPatternFor("prop24", omniture, row.Cell(35).GetValue<string>(), propBuilder.prop24);
                DetectPatternFor("prop25", omniture, row.Cell(36).GetValue<string>(), propBuilder.prop25);

                if (!string.IsNullOrEmpty(omniture.ParentId)) {
                    var parentElementRow = FindRowById(omniture.ParentId, worksheet);
                    if (parentElementRow != null) {
                        var omnitureParent = new Omniture(parentElementRow);
                        var properties = typeof(Omniture).GetProperties();
                        properties
                            .Where(p => p.GetValue(omniture, null).GetType() == typeof(string) &&
                                string.Equals(
                                    p.GetValue(omniture, null).ToString(), 
                                    p.GetValue(omnitureParent, null).ToString(), 
                                    StringComparison.InvariantCultureIgnoreCase)
                                ).ToList()
                                .ForEach(p => p.SetValue(omniture, null, null));
                    } else {
                        omniture.ParentId = string.Empty;
                    }
                }

                trackingId = row.Cell(1).GetValue<string>();
                if (string.IsNullOrEmpty(trackingId)) {
                    maxTrackingId++;
                    trackingId = string.Format("t-{0}", maxTrackingId);
                    row.Cell(1).Value = string.Format("t-{0}", maxTrackingId);
                }

                if (!string.Equals(trackingId, "t-0", StringComparison.InvariantCultureIgnoreCase) && !omniture.isEmpty) {
                    tracking.Add((XElement)omniture);
                }

                if (tracking != null && !string.IsNullOrEmpty(trackingId)) {
                    tracking.SetAttributeValue("id", trackingId);
                    definitions.Add(tracking);
                }
            }
            

            var floodlightWorksheet = workbook.Worksheets.Where(x => x.Name.ToLower() == "floodlight").FirstOrDefault();
            lastRowUsed = floodlightWorksheet.LastRowUsed().RowNumber();
            lastColumnUsed = floodlightWorksheet.LastColumnUsed().ColumnNumber();

            trackingId = "";
            for (int rowIndex = 3; rowIndex <= lastRowUsed; rowIndex++) {
                var row = floodlightWorksheet.Row(rowIndex);

                var element = definitions.Elements("tracking").Where(x => string.Equals(
                    x.Attribute("id").Value,
                    row.Cell(1).GetValue<string>(),
                    StringComparison.InvariantCultureIgnoreCase)
                ).SingleOrDefault();

                if (element == null) {
                    trackingId = row.Cell(1).GetValue<string>();
                    if (string.IsNullOrEmpty(trackingId)) {
                        maxTrackingId++;
                        trackingId = string.Format("t-{0}", maxTrackingId);
                        row.Cell(1).Value = string.Format("t-{0}", maxTrackingId);
                    }
                    element = new XElement("tracking");
                    element.SetAttributeValue("id", trackingId);
                }

                var floodlight = new Floodlight(row);

                var floodlightElement = element.Elements("floodlight").Where(x => x.Elements("category").Where(c => c.Attribute("varyby") != null).Count() > 0).SingleOrDefault();
                if (floodlightElement != null && floodlight.Category.Items != null && floodlight.Category.Items.Count > 0) {
                    floodlightElement.Element("category").Add(
                        floodlight.Category.Items.Select<KeyValuePair<string, string>, XElement>(x => {
                            XElement valueElement = new XElement("value", x.Value);
                            valueElement.SetAttributeValue("for", x.Key);
                            return valueElement;
                        })
                    );
                } else {
                    element.Add((XElement)floodlight);
                }
            }

            floodlightWorksheet
                .Rows()
                .Where(x => x.Cell(1).GetValue<string>().StartsWith("t-"))
                .Select(x => {
                    var floodlight = new Floodlight(x);

                    return floodlight;
                })
                .ToList()
                .ForEach(x => { });

            Trace.WriteLine("saving tracking xml file");
            definitions.Save(trackingOutputFileName);
            Trace.WriteLine("saved");

            XElement metadata = new XElement("metadata"); ;

            var metadatatWorksheet = workbook.Worksheets.Where(x => x.Name.ToLower() == "meta").FirstOrDefault();
            lastRowUsed = floodlightWorksheet.LastRowUsed().RowNumber();

            metadata.Add(metadatatWorksheet
                .Rows()
                .Where(x => x.Cell(1).GetValue<string>().StartsWith("/"))
                .Select(x => {
                    var url = new XElement("url");
                    var id = x.Cell(1).GetValue<string>();

                    id = id.EndsWith("/") ? id : id + "/";

                    url.SetAttributeValue("id", id);
                    url.Add(new XElement("title", x.Cell(3).GetValue<string>()));
                    url.Add(new XElement("description", x.Cell(4).GetValue<string>()));
                    url.Add(new XElement("keywords", x.Cell(5).GetValue<string>()));
                    url.Add(new XElement("pageViewTracking", x.Cell(2).GetValue<string>()));

                    return url;
                }));

            Trace.WriteLine("saving metadata xml file");
            metadata.Save(metadataOutputFileName);
            Trace.WriteLine("saved");

            Trace.WriteLine("saving omniture workbook changes (ids and stuff)...");
            Trace.WriteLine("this may take a while, I'm not excel you know...");

            var fileName = String.Format("{0}-{1}.xlsx", Path.GetFileNameWithoutExtension(inputFileName), FileTimeStamp);

            workbook.SaveAs(Path.Combine(Path.GetDirectoryName(inputFileName), fileName));
            Trace.WriteLine("saved");
        }
        private static IXLRow FindRowById(string id, IXLWorksheet worksheet) { 
            var lastRowUsed = worksheet.LastRowUsed().RowNumber();
            for (int rowIndex = 0; rowIndex <= lastRowUsed; rowIndex++) {
                if (worksheet.Row(rowIndex).Cell(1).GetValue<string>().ToLower() == id.ToLower()) {
                    return worksheet.Row(rowIndex);
                }
            }
            return null;
        }
        private static int GetMaxTrackingId(XElement trackingXml) {
            var trackingElements = trackingXml.Elements("tracking");
            if (trackingElements != null && trackingElements.Count() > 0) {
                int maxTrackingId = trackingElements.Max(x => { int result; int.TryParse(x.Attribute("id").Value.Substring(2), out result); return result; });
                return maxTrackingId;
            }
            return 0;
        }
        private static void CompareOmnitureProperties(ClosedXML.Excel.IXLRow row, int cell, XElement omniture, string property) {
            var propertyElement = omniture.Elements(property).SingleOrDefault();
            var propertyValue = propertyElement != null ? propertyElement.Value : string.Empty;
            if (propertyValue.ToLower() != row.Cell(cell).GetValue<string>().ToLower()) {
                if (propertyElement == null) {
                    propertyElement = new XElement(property);
                    omniture.Add(propertyElement);
                }
                propertyElement.SetValue(row.Cell(cell).GetValue<string>().ToLower());
            }
        }
        private static string BuildSkipList(IXLRow row, int lastColumnUsed) {
            StringBuilder skipList = new StringBuilder();
            for (int cellIndex = 12; cellIndex <= lastColumnUsed - 2; cellIndex++) {
                if (row.Cell(cellIndex).GetValue<string>() == string.Empty) {
                    if (skipList.Length > 0) {
                        skipList.Append(",");
                    }
                    skipList.Append(cellIndex - 11);
                }
            }
            return skipList.ToString();
        }
        private static void DetectPatternFor(string propertyName, Omniture omniture, string sourceValue, string calculatedValue) {
            if (!string.IsNullOrEmpty(sourceValue) && calculatedValue != null) {

                Dictionary<string, string> javascriptReplacements = new Dictionary<string, string>() { 
                    {"Domain", "domain"},
                    {"SiteSection", "ss"},
                    {"SubSection","sub"},
                    {"ActiveState","a"},
                    {"ObjectValue","sel"},
                    {"ObjectDescription","od"},
                    {"CallToAction", "cta"}
                };

                if (sourceValue.ToLower() != calculatedValue) {
                    var pattern = new StringBuilder();
                    var propertyParts = sourceValue.Split('|');

                    var properties = typeof(Omniture).GetProperties();

                    for (int i = 0; i < propertyParts.Length; i++) {

                        if (pattern.Length > 0) {
                            pattern.Append(" | ");
                        }

                        var objectProperty = properties.Where(
                            x => x.GetValue(omniture, null).ToString().ToLower() == propertyParts[i].ToLower().Trim())
                            .FirstOrDefault();

                        if (objectProperty != null && javascriptReplacements.ContainsKey(objectProperty.Name)) {
                            pattern
                                .Append("{")
                                .Append(javascriptReplacements[objectProperty.Name])
                                .Append("}");
                        } else {
                            pattern.Append(propertyParts[i].ToLower().Trim());
                        }
                    }
                    omniture.Exceptions.Add(propertyName, pattern.ToString().Trim(new char[] { '|', ' ' }).Replace(" | | ", " | "));
                }
            }
        }
    }
}