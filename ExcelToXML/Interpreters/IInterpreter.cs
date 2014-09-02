using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToXML.Interpreters {
    public interface IInterpreter {
        String FileTimeStamp { get; set; }
        void CreateTrackingXml(string inputFileName, string trackingOutputFileName, string metadataOutputFileName);
    }
}