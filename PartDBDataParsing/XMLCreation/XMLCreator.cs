using PartDBDataParsing.ExcelParser.SheetType;
using System.Collections.Generic;
using System.Xml.Linq;
using System;
using System.IO;
using System.Reflection;

namespace PartDBDataParsing.XMLCreation
{
    public class XMLCreator
    {
        public static void SaveXML(IEnumerable<CommonExcelData> parts)
        {
            var doc = new XDocument();

            var root = new XElement(
                "PartSpecDB",
                new XAttribute("Company", "Pentacube"),
                new XAttribute("Version", "0.3"));

            // root의 하위 항목에 Parts들을 삽입
            foreach (var part in parts)
                root.Add(part.CreateXmlElement());

            doc.Add(root);

            var outputDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ouput");
            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory).Create();

            var outputFilename = Assembly.GetExecutingAssembly().GetName().Name + ".xml";
            var outputPath = Path.Combine(outputDirectory, outputFilename);
            doc.Save(outputPath);
        }
    }
}
