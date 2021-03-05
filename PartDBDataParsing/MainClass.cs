using PartDBDataParsing.ExcelParser.SheetType;
using PartDBDataParsing.ExcelParser;
using System;
using System.Collections.Generic;
using System.Linq;
using JetBrains.Annotations;
using PartDBDataParsing.XMLCreation;

namespace PartDBDataParsing
{
    public class MainClass
    {
        public static void Main(string[] args)
        {
            // args의 첫 인자는 반드시 엑셀 파일의 경로를 의미함
            var sourceFileName = args.FirstOrDefault();
            if (sourceFileName == null)
                return;

            var parts = GetStorage(sourceFileName);  // 1. 엑셀 -> 클래스
            XMLCreator.SaveXML(parts);               // 2. 클래스 -> XML File

            //Console.ReadKey(); // 콘솔창 자동 닫힘 방지
        }

        [NotNull, ItemNotNull]
        private static IEnumerable<CommonExcelData> GetStorage([NotNull] string sourceFile)
        {
            return DefaultExcelParser.Parse(sourceFile).SelectMany(ParseSheet);
        }

        [NotNull, ItemNotNull]
        private static IEnumerable<CommonExcelData> ParseSheet(ParsedExcelData excelData)
        {
            switch (excelData.SheetName.ToUpper())
            {
                case "CHIP RESISTER": return ChipResistor.Parse(excelData);
                case "ALUMI ELEC CHIP CAP": return AlumiElecChipCap.Parse(excelData);
                case "CHIP CAP": return ChipCap.Parse(excelData);
                case "MOSFET": return MOSFET.Parse(excelData);
                case "INDUCTOR COIL": return InductorCoil.Parse(excelData);
                case "SCHOTTKY DIODE": return SchottkyDiode.Parse(excelData);
                case "SWITCHING DIODE": return SwitchingDiode.Parse(excelData);

                case "ARRY TR":
                case "GENERAL PURPOSE TR":
                    return BJT.Parse(excelData);

                case "OP AMP":
                case "REGULATOR":
                    return IC.Parse(excelData);

                // TODO: 타입 미정(보류)
                case "SWITCHING TR DIGITAL TR":
                case "DCDC PMIC":
                default:
                    return NoSpec.Parse(excelData);
            }
        }
    }
}
