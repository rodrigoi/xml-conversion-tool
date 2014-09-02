using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using System.Xml.Linq;
using ClosedXML.Excel;
using System.IO;
using ExcelToXML.Model;
using System.Dynamic;
using ExcelToXML.Services.Omniture;
using ExcelToXML.Interpreters;
using ExcelToXML.Factories;

namespace ExcelToXML {
    class Program {
        static void Main(string[] args) {
            Trace.Listeners.Clear();
            Trace.AutoFlush = true;
            Trace.Listeners.Add(new ConsoleTraceListener(false));

            var fileTimeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");

            Stopwatch watch = new Stopwatch();
            watch.Start();

            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-US");

            if (args.Length > 0 && args.Length == 3) {
                string sourceFileName = args[0];
                string trackingDestinationFileName = args[1];
                string metadataDestinationFileName = args[2];

                var excelFolder = Path.GetDirectoryName(sourceFileName);
                Trace.Listeners.Add(new TextWriterTraceListener(Path.Combine(excelFolder, string.Format("Log-{0}.txt", fileTimeStamp))));

                if (File.Exists(sourceFileName)) {
                    IInterpreter sourceInterpreter = InterpreterFactory.GetInterpreterFor(sourceFileName);

                    if (sourceInterpreter != null) {
                        sourceInterpreter.FileTimeStamp = fileTimeStamp;
                        sourceInterpreter.CreateTrackingXml(sourceFileName, trackingDestinationFileName, metadataDestinationFileName);
                    } else {
                        Trace.WriteLine("There's no interpreter for that file extension.");
                    }
                } else {
                    Trace.WriteLine("The source file do not exists. Please, try again with a different file.");
                }
            } else {
                Trace.WriteLine("You have to specify an imput excel spreadsheet and an output xml file ;)");
            }

            watch.Stop();
            Trace.WriteLine(String.Format("Time is an illusion, but it took {0} to do all the work.", watch.Elapsed.ToString()));
            Trace.WriteLine("So long and thanks for all the fish. Press any key to exit.");
            Console.ReadKey();
        }
    }
}