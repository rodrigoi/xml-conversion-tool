using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelToXML.Interpreters;
using System.IO;

namespace ExcelToXML.Factories {
    public sealed class InterpreterFactory {
        private InterpreterFactory() { }

        public static IInterpreter GetInterpreterFor(string fileName) {
            var extension = Path.GetExtension(fileName);

            var interpreterType = System.Reflection.Assembly.GetExecutingAssembly().GetTypes()
                .Where(t => t.GetCustomAttributes(typeof(SourceFileExtensionAttribute), true).Length > 0)
                .Where(t => t.GetCustomAttributesData().Where(cad => cad.ConstructorArguments.First().Value.ToString().ToLower() == extension).SingleOrDefault() != null)
                .SingleOrDefault();

            if (interpreterType != null) {
                return System.Activator.CreateInstance(interpreterType) as IInterpreter;
            }
            return null;
        }
    }
}