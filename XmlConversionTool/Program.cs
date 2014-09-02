using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace XmlConversionTool {
    class Program {
        static void Main(string[] args) {
            XElement originalFile = XElement.Load(@"C:\Filez-Globant\workspace\goodyear-webapp\src\main\webapp\WEB-INF\xml\en_US\tracking.xml");
            XElement newDefinitions = new XElement("definitions");

            IEnumerable<string> ids = originalFile.Elements("tracking").Select(x => (string)x.Attribute("id"));

            foreach (string id in ids) {
                Console.Write("Processing {0} ...", id);

                XElement trackingElement = originalFile.Elements("tracking").Single(x => x.Attribute("id").Value == id);

                XElement newTrackingElement = new XElement("tracking");
                newTrackingElement.SetAttributeValue("id", id);

                ProcessOmniture(trackingElement, newTrackingElement);

                ProcessFloodlight(trackingElement, newTrackingElement);

                ProcessEngagementTags(trackingElement, newTrackingElement);

                newDefinitions.Add(newTrackingElement);

                Console.Write(" Processed :)");
                Console.WriteLine();
            }

            XDocument convertedFile = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), newDefinitions);
            convertedFile.Save(@"c:\tmp\tracking.xml");
        }

        private static void ProcessEngagementTags(XElement trackingElement, XElement newTrackingElement) {
            var engagementTags = trackingElement.Elements().Where(x => x.Name == "engagement");
            if (engagementTags != null && engagementTags.Count() > 0) {

                foreach (var engagement in engagementTags) {
                    var floodlight = engagement.Elements().Select(x => x);
                    if (floodlight != null && floodlight.Count() > 0) {
                        newTrackingElement.Add(new XElement("floodlight", floodlight));
                    }
                }
            }
        }

        private static void ProcessOmniture(XElement trackingElement, XElement newTrackingElement) {
            XElement omniture = new XElement("omniture");

            var element = trackingElement.Elements().Where(x => x.Name == "siteSection").SingleOrDefault();
            if (element != null && !string.IsNullOrEmpty(element.Value)) {
                omniture.Add(new XElement(element.Name, element.Value));
            }

            element = trackingElement.Elements().Where(x => x.Name == "subSection").SingleOrDefault();
            if (element != null && !string.IsNullOrEmpty(element.Value)) {
                omniture.Add(new XElement(element.Name, element.Value));
            }

            element = trackingElement.Elements().Where(x => x.Name == "activeState").SingleOrDefault();
            if (element != null && !string.IsNullOrEmpty(element.Value)) {
                omniture.Add(new XElement(element.Name, element.Value));
            }

            element = trackingElement.Elements().Where(x => x.Name == "object").SingleOrDefault();
            if (element != null && !string.IsNullOrEmpty(element.Value)) {
                omniture.Add(new XElement(element.Name, element.Value));
            }

            element = trackingElement.Elements().Where(x => x.Name == "objectDescription").SingleOrDefault();
            if (element != null && !string.IsNullOrEmpty(element.Value)) {
                omniture.Add(new XElement(element.Name, element.Value));
            }

            element = trackingElement.Elements().Where(x => x.Name == "callToAction").SingleOrDefault();
            if (element != null && !string.IsNullOrEmpty(element.Value)) {
                omniture.Add(new XElement(element.Name, element.Value));
            }

            element = trackingElement.Elements().Where(x => x.Name == "skip").SingleOrDefault();
            if (element != null && !string.IsNullOrEmpty(element.Value)) {
                omniture.Add(new XElement(element.Name, element.Value));
            }

            element = trackingElement.Elements().Where(x => x.Name == "events").SingleOrDefault();
            if (element != null && !string.IsNullOrEmpty(element.Value)) {
                omniture.Add(new XElement(element.Name, element.Value));
            }

            var omnitureExceptions = trackingElement.Elements().Where(x => x.Name == "omniture").Select(x => x.Elements());
            if (omnitureExceptions != null && omnitureExceptions.Count() > 0) {
                omniture.Add(new XElement("exceptions", omnitureExceptions));
            }

            if (omniture.Elements().Count() > 0) {
                newTrackingElement.Add(omniture);
            }
        }
        private static void ProcessFloodlight(XElement trackingElement, XElement newTrackingElement) {
            XElement floodlight = new XElement("floodlight");

            var element = trackingElement.Elements().Where(x => x.Name == "src").SingleOrDefault();
            if (element != null && !string.IsNullOrEmpty(element.Value)) {
                floodlight.Add(new XElement("source", element.Value));
            }

            element = trackingElement.Elements().Where(x => x.Name == "type").SingleOrDefault();
            if (element != null && !string.IsNullOrEmpty(element.Value)) {
                floodlight.Add(new XElement(element.Name, element.Value));
            }

            element = trackingElement.Elements().Where(x => x.Name == "cat").SingleOrDefault();
            if (element != null && !string.IsNullOrEmpty(element.Value)) {
                floodlight.Add(new XElement("category", element.Value));
            }

            element = trackingElement.Elements().Where(x => x.Name == "session").SingleOrDefault();
            if (element != null && !string.IsNullOrEmpty(element.Value)) {
                floodlight.Add(new XElement(element.Name, element.Value));
            }

            var varyby = trackingElement.Elements().Where(x => x.Name == "varyby");
            if (varyby != null && varyby.Elements().Count() > 0) {
                foreach (var by in varyby) {

                    if (by != null && by.Elements().Count() > 0) {

                        var cat = new XElement("cat");
                        cat.SetAttributeValue("varyby", by.Attribute("variable").Value);
                        foreach (var item in by.Elements()) {
                            XElement value = new XElement("value", item.Value);
                            value.SetAttributeValue("for", item.Attribute("item").Value);
                            cat.Add(value);
                        }
                        floodlight.Add(cat);
                    }
                }
            }
            if (floodlight.Elements().Count() > 0) {
                newTrackingElement.Add(floodlight);
            }
        }
    }
}