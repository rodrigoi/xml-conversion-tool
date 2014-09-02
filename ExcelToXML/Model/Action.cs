using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToXML.Model {
    public class Action {
        public string Id { get; set; }
        public List<Omniture> Omniture { get; set; }
        public List<Floodlight> Floodlight { get; set; }
    }
}