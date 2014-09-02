using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Xml.Linq;

namespace ExcelToXML.Model {
    public class Floodlight {
        public string Source { get; set; }
        public string Type { get; set; }
        public FloodlightProperty Category { get; set; }
        public string Session { get; set; }

        public Floodlight() {
            this.Source = string.Empty;
            this.Type = string.Empty;
            this.Session = string.Empty;
        }
        public Floodlight(IXLRow row) {
            this.Source = row.Cell(2).GetValue<string>().ToLower().Trim();
            this.Type = row.Cell(3).GetValue<string>().ToLower().Trim();
            this.Category = new FloodlightProperty(row);
            this.Session = row.Cell(7).GetValue<string>().ToLower().Trim();
        }
        public Floodlight(XElement xmlElement) {
            this.Source = xmlElement.Element("source") != null ? xmlElement.Element("source").Value : string.Empty;
            this.Type = xmlElement.Element("type") != null ? xmlElement.Element("type").Value : string.Empty;

            this.Category = new FloodlightProperty(xmlElement.Element("category"));

            this.Session = xmlElement.Element("session") != null ? xmlElement.Element("session").Value : string.Empty;
        }

        public static explicit operator XElement(Floodlight source) {
            var floodlight = new XElement("floodlight");

            if (!string.IsNullOrEmpty(source.Source)) {
                floodlight.Add(new XElement("source", source.Source));
            }

            if (!string.IsNullOrEmpty(source.Type)) {
                floodlight.Add(new XElement("type", source.Type));
            }

            floodlight.Add((XElement)source.Category);

            if (!string.IsNullOrEmpty(source.Session)) {
                floodlight.Add(new XElement("session", source.Session));
            }

            return floodlight;
        }
    }
}