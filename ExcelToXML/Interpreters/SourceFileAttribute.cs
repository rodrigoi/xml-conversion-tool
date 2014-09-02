using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToXML.Interpreters {
    [AttributeUsage(AttributeTargets.Class, AllowMultiple=true)]
    public class SourceFileExtensionAttribute : Attribute {
        public string FileExtension { get; set; }

        public SourceFileExtensionAttribute(string fileExtension) {
            this.FileExtension = fileExtension;
        }
    }
}