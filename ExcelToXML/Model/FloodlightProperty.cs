using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Xml.Linq;

namespace ExcelToXML.Model {
    public class FloodlightProperty {
        public string DefaultValue { get; set; }
        public string VariableName { get; set; }
        public Dictionary<string, string> Items;

        public FloodlightProperty(IXLRow row) {
            this.DefaultValue = row.Cell(6).GetValue<string>().ToLower().Trim();
            this.VariableName = row.Cell(4).GetValue<string>().ToLower().Trim();

            if (!string.IsNullOrEmpty(this.VariableName)) {
                this.Items = new Dictionary<string, string>();
                this.Items.Add(
                    row.Cell(5).GetValue<string>().ToLower().Trim(),
                    row.Cell(6).GetValue<string>().ToLower().Trim()
                );
            }
        }
        public FloodlightProperty(XElement xmlElement){
            if (xmlElement != null) {
                if (xmlElement.Elements("value").Count() > 0) {
                    this.Items = xmlElement.Descendants("value").ToDictionary<XElement, string, string>(
                        x => x.Attribute("for").Value,
                        x => x.Value
                    );
                    this.VariableName = xmlElement.Attribute("varyby").Value;
                } else {
                    this.DefaultValue = xmlElement.Value;
                }
            }
        }

        public static explicit operator XElement(FloodlightProperty source) {
            if (source.Items != null && source.Items.Count > 0) {
                var category = new XElement("category");
                category.SetAttributeValue("varyby", source.VariableName);

                foreach (var item in source.Items) {
                    var value = new XElement("value", item.Value);
                    value.SetAttributeValue("for", item.Key);

                    category.Add(value);
                }

                return category;
            } else {
                return new XElement("category", source.DefaultValue);
            }
        }
    }
}