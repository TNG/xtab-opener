using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XtabFileOpener.TableContainer.SpreadsheetTableContainer.ExcelTableContainer;
using Microsoft.Office.Interop.Excel;
using Moq;

namespace XtabFileOpenerTest
{

    [TestClass]
    public class ExcelTableTest
    {
        ExcelTable excelTable;

        [TestInitialize]
        public void setUp()
        {
            DateTime dt1 = new DateTime(2016, 12, 31);
            DateTime dt2 = new DateTime(2016, 12, 31, hour: 23, minute: 59, second: 59);     
            DateTime dt3 = new DateTime(2016, 12, 31, hour:23, minute:59, second:59, millisecond:9);
            object[,] contents = new object[1, 4] { { "text value", dt1, dt2, dt3 } };

            var cells = new Mock<Microsoft.Office.Interop.Excel.Range>();
            cells.Setup(r => r.get_Value(It.IsAny<object>())).Returns(contents);
            excelTable = new ExcelTable("name", cells.Object);
        }

        [TestMethod]
        public void testConversionOfDateTimeObjectsDuringConstruction()
        {
            Assert.AreEqual("text value", excelTable.TableArray[0, 0]);
            Assert.AreEqual("2016-12-31 00:00:00", excelTable.TableArray[0, 1]);
            Assert.AreEqual("2016-12-31 23:59:59", excelTable.TableArray[0, 2]);
            Assert.AreEqual("2016-12-31 23:59:59.009000000", excelTable.TableArray[0, 3]);
        }
    }
}
