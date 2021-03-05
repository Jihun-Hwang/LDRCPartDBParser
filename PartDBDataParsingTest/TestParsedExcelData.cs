using NUnit.Framework;
using PartDBDataParsing.ExcelParser;
using PartDBDataParsing.ExcelParser.SheetType;
using Pentacube.CAx.ECAD.LcadIf;
using Pentacube.Infrastructure.Service;
using System.Collections.Generic;

namespace PartDBDataParsingTest
{
    [TestFixture]
    public class TestParsedExcelData
    {
        [TestCase("ABC123", "ABC123")]
        [TestCase("a", "a")]
        public void TestCreate(string sheetName, string expectedName)
        {
            // given

            // when
            var actual = new ParsedExcelData(sheetName, new List<string[]>());

            // then
            Assert.AreEqual(expectedName, actual.SheetName);
        }

        [TestCase("330uF", 330, eSiPrefix.u)]
        [TestCase("150V", 150, eSiPrefix.None)]
        [TestCase("1Kohm", 1, eSiPrefix.k)]
        [TestCase("/35", 35, eSiPrefix.None)]
        public void TestConvertUnit(string input, double value, eSiPrefix unit)
        {
            // when
            var actual = CommonExcelData.ConvertUnit(input);

            // then
            Assert.AreEqual(value, actual.Value);
            Assert.AreEqual(unit, actual.Prefix);
        }

        [TestCase(eSiPrefix.None, eCapacitanceUnit.F)]
        [TestCase(eSiPrefix.m, eCapacitanceUnit.mF)]
        [TestCase(eSiPrefix.u, eCapacitanceUnit.uF)]
        [TestCase(eSiPrefix.µ, eCapacitanceUnit.uF)]
        [TestCase(eSiPrefix.p, eCapacitanceUnit.pF)]
        public void TestConvertUnitType(eSiPrefix from, eCapacitanceUnit to)
        {
            // when
            //var convertedVal = CommonExcelData.ConvertSItoCapacitanceUnit(from);  // public으로 수정후 테스트

            // then
            //Assert.AreEqual(to, convertedVal);
        }
    }
}