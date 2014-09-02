using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using ClosedXML.Excel;

namespace ExcelToXML.Model {
    public class Omniture {
        public string ParentId { get; set; }
        public string Domain { get; set; }
        public string SiteSection { get; set; }
        public string SubSection { get; set; }
        public string ActiveState { get; set; }
        public string ObjectValue { get; set; }
        public string ObjectDescription { get; set; }
        public string CallToAction { get; set; }
        public string Events { get; set; }
        public string Skip { get; set; }
        public Dictionary<string, string> Exceptions { get; set; }

        public Boolean isEmpty { 
            get {
                return string.IsNullOrEmpty(this.Domain) &&
                       string.IsNullOrEmpty(this.SiteSection) &&
                       string.IsNullOrEmpty(this.SubSection) &&
                       string.IsNullOrEmpty(this.ActiveState) &&
                       string.IsNullOrEmpty(this.ObjectValue) &&
                       string.IsNullOrWhiteSpace(this.ObjectDescription) &&
                       string.IsNullOrEmpty(this.CallToAction) &&
                       string.IsNullOrEmpty(this.Events);
            }
        }

        public Omniture() {
            this.ParentId = string.Empty;
            this.Domain = string.Empty;
            this.SiteSection = string.Empty;
            this.SubSection = string.Empty;
            this.ActiveState = string.Empty;
            this.ObjectValue = string.Empty;
            this.ObjectDescription = string.Empty;
            this.CallToAction = string.Empty;
            this.Events = string.Empty;
            this.Skip = string.Empty;
            this.Exceptions = new Dictionary<string, string>();
        }
        public Omniture(IXLRow row) {
            this.ParentId = row.Cell(2).GetValue<string>().ToLower().Trim();
            this.Domain = row.Cell(3).GetValue<string>().ToLower().Trim();
            this.SiteSection = row.Cell(4).GetValue<string>().ToLower().Trim();
            this.SubSection = row.Cell(5).GetValue<string>().ToLower().Trim();
            this.ActiveState = row.Cell(6).GetValue<string>().ToLower().Trim();
            this.ObjectValue = row.Cell(7).GetValue<string>().ToLower().Trim();
            this.ObjectDescription = row.Cell(8).GetValue<string>().ToLower().Trim();
            this.CallToAction = row.Cell(9).GetValue<string>().ToLower().Trim();
            this.Events = row.Cell(11).GetValue<string>().ToLower().Trim();

            this.Skip = string.Empty;

            this.Exceptions = new Dictionary<string, string>();
        }

        public static explicit operator XElement(Omniture source) {
            var omniture = new XElement("omniture");

            if (!string.IsNullOrEmpty(source.ParentId)) {
                omniture.Add(new XElement("parentId", source.ParentId));
            }

            if (!string.IsNullOrEmpty(source.SiteSection)) {
                omniture.Add(new XElement("siteSection", source.SiteSection));
            }

            if (!string.IsNullOrEmpty(source.SubSection)) {
                omniture.Add(new XElement("subSection", source.SubSection));
            }

            if (!string.IsNullOrEmpty(source.ActiveState)) {
                omniture.Add(new XElement("activeState", source.ActiveState));
            }

            if (!string.IsNullOrEmpty(source.ObjectValue)) {
                omniture.Add(new XElement("object", source.ObjectValue));
            }

            if (!string.IsNullOrEmpty(source.ObjectDescription)) {
                omniture.Add(new XElement("objectDescription", source.ObjectDescription));
            }

            if (!string.IsNullOrEmpty(source.CallToAction)) {
                omniture.Add(new XElement("callToAction", source.CallToAction));
            }

            if (!string.IsNullOrEmpty(source.Events)) {
                omniture.Add(new XElement("events", source.Events));
            }

            if (!string.IsNullOrEmpty(source.Skip)) {
                omniture.Add(new XElement("skip", source.Skip));
            }

            if (source.Exceptions != null && source.Exceptions.Count() > 0) {
                var exceptions = new XElement("exceptions");
                exceptions.Add(source.Exceptions.Select(x => new XElement(x.Key, x.Value)));
                omniture.Add(exceptions);
            }

            return omniture;
        }
    }
}