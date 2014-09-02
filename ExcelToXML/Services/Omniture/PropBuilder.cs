using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Dynamic;

namespace ExcelToXML.Services.Omniture {
    public class PropBuilder : DynamicObject {
        private Dictionary<string, string> properties = new Dictionary<string, string>();

        public PropBuilder(ExcelToXML.Model.Omniture omniture) {
            properties.Add("pageName", string.Format("{0} | {1} | {2} | {3} | {4} | {5}",
                omniture.Domain,
                omniture.SiteSection,
                omniture.SubSection,
                omniture.ActiveState,
                omniture.ObjectValue.ToLower() == omniture.ActiveState ? string.Empty : omniture.ObjectValue,
                omniture.ObjectDescription
            ).Trim(new char[] { '|', ' ' }).Replace(" |  | ", " | ").ToLower());

            properties.Add("prop1", omniture.SiteSection.ToLower());
            
            properties.Add("prop2", omniture.SubSection.ToLower());

            properties.Add("prop3", string.Format("{0} | {1}",
                omniture.SiteSection,
                omniture.SubSection
            ).Trim(new char[] { '|', ' ' }).Replace(" |  | ", " | ").ToLower());

            properties.Add("prop4", omniture.ActiveState.ToLower());

            properties.Add("prop5", string.Format("{0} | {1}",
                omniture.SubSection,
                omniture.ActiveState
            ).Trim(new char[] { '|', ' ' }).Replace(" |  | ", " | ").ToLower());

            properties.Add("prop6", string.Format("{0} | {1} | {2}",
                omniture.SiteSection,
                omniture.SubSection,
                omniture.ActiveState
            ).Trim(new char[] { '|', ' ' }).Replace(" |  | ", " | ").ToLower());

            properties.Add("prop7", omniture.ObjectValue.ToLower());

            properties.Add("prop8", string.Format("{0} | {1}",
                omniture.SiteSection,
                omniture.ActiveState
            ).Trim(new char[] { '|', ' ' }).Replace(" |  | ", " | ").ToLower());

            properties.Add("prop9", string.Format("{0} | {1} | {2}",
                omniture.SubSection,
                omniture.ActiveState,
                omniture.ObjectValue
            ).Trim(new char[] { '|', ' ' }).Replace(" |  | ", " | ").ToLower());

            properties.Add("prop10", string.Format("{0} | {1} | {2} | {3}",
                omniture.SiteSection,
                omniture.SubSection,
                omniture.SiteSection,
                omniture.ActiveState
            ).Trim(new char[] { '|', ' ' }).Replace(" |  | ", " | ").ToLower());

            properties.Add("prop11", omniture.ObjectDescription.ToLower());

            properties.Add("prop12", string.Format("{0} | {1}",
                omniture.ObjectValue,
                omniture.ObjectDescription
            ).Trim(new char[] { '|', ' ' }).Replace(" |  | ", " | ").ToLower());

            properties.Add("prop13", string.Format("{0} | {1} | {2}",
                omniture.ActiveState,
                omniture.ObjectValue,
                omniture.ObjectDescription
            ).Trim(new char[] { '|', ' ' }).Replace(" |  | ", " | ").ToLower());

            properties.Add("prop14", string.Format("{0} | {1} | {2} | {3}",
                omniture.SubSection,
                omniture.ActiveState,
                omniture.ObjectValue,
                omniture.ObjectDescription
            ).Trim(new char[] { '|', ' ' }).Replace(" |  | ", " | ").ToLower());

            properties.Add("prop15", string.Format("{0} | {1} | {2} | {3} | {4}",
                omniture.SiteSection,
                omniture.SubSection,
                omniture.ActiveState,
                omniture.ObjectValue,
                omniture.ObjectDescription
            ).Trim(new char[] { '|', ' ' }).Replace(" |  | ", " | ").ToLower());

            properties.Add("prop16", string.Empty);
            properties.Add("prop17", string.Empty);
            properties.Add("prop18", string.Empty);
            properties.Add("prop19", string.Empty);
            properties.Add("prop20", string.Empty);
            properties.Add("prop21", string.Empty);
            properties.Add("prop22", string.Empty);
            properties.Add("prop23", string.Empty);
            properties.Add("prop24", string.Empty);
            properties.Add("prop25", string.Empty);
            properties.Add("prop26", string.Empty);
            properties.Add("prop27", string.Empty);
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result) {
            result = this.properties.Where(x => x.Key == binder.Name).Select(x => x.Value).FirstOrDefault();
            return result != null;
        }
    }
}