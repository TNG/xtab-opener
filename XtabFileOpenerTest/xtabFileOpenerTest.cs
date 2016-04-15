using System;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using XtabFileOpenerTest.SpreadsheetMock;
using XtabFileOpener.TableContainer;
using XtabFileOpener.TableContainer.XmlTableContainer;
using System.Reflection;

namespace XtabFileOpenerTest
{
    /// <summary>
    /// tests the whole functionality of the xtabFileOpener
    /// </summary>
    [TestClass]
    public class xtabFileOpenerTest
    {
        private static readonly string fileName = "XtabFileOpenerTest.Resources.TestFile.xtab";
        private static readonly string addFileName = "XtabFileOpenerTest.Resources.TestFileAdd.xtab";
        private static readonly string assertFileName = "XtabFileOpenerTest.Resources.TestFileAssert.xtab";

        private FileInfo file;
        private TableContainer originFileCon;
        private TableContainer addFileCon;
        private TableContainer assertFileCon;

        private XtabFileOpener.XtabFileOpener opener;
        private SpreadsheetAdapterMock spreadSheetMock;
        private Thread mainThread;

        [TestInitialize]
        public void setUp()
        {
            createTmpFileFromResource();
            opener = new XtabFileOpener.XtabFileOpener(file.FullName);
            spreadSheetMock = new SpreadsheetAdapterMock();
            opener.openFile(spreadSheetMock.mock.Object);
            mainThread = new Thread(() => opener.waitForClosing());
            mainThread.Start();

            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(fileName))
            {
                originFileCon = XmlTableContainer.createFromStream(stream, fileName);
            }

            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(addFileName))
            {
                addFileCon = XmlTableContainer.createFromStream(stream, addFileName);
            }
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(assertFileName))
            {
                assertFileCon = XmlTableContainer.createFromStream(stream, assertFileName);
            }
        }

        private void createTmpFileFromResource()
        {
            using (Stream resource = Assembly.GetExecutingAssembly().GetManifestResourceStream(fileName))
            {
                file = new FileInfo(Path.GetTempFileName());
                using (Stream output = File.OpenWrite(file.FullName))
                {
                    resource.CopyTo(output);
                }
            }
        }

        [TestCleanup]
        public void endTest()
        {
            spreadSheetMock.mock.Object.closeSpreadsheet();
            Thread.Sleep(1000);
            file.Delete();
            Assert.AreEqual(mainThread.IsAlive, false);
        }

        [TestMethod]
        public void testEditXtabFileWithEndSaving()
        {
            editFileWithAddFile(false);
            spreadSheetMock.mock.Object.saveSpreadsheet();
            Assert.AreEqual(XmlTableContainer.createFromFile(file), assertFileCon);
        }

        [TestMethod]
        public void testEditXtabFileWithOftenSaving()
        {
            editFileWithAddFile(true);
            Assert.AreEqual(XmlTableContainer.createFromFile(file), assertFileCon);
        }

        [TestMethod]
        public void testEditXtabFileWithoutSaving()
        {
            editFileWithAddFile(false);
            Assert.AreEqual(XmlTableContainer.createFromFile(file), originFileCon);
        }

        private void editFileWithAddFile(bool saveAfterEveryEdit)
        {
            foreach (Table table in addFileCon)
            {
                int numOfSheet = spreadSheetMock.getNumberOfSheet(table.Name);
                if (numOfSheet != -1)
                    addTableToFile(table, numOfSheet);
                else
                {
                    numOfSheet = spreadSheetMock.mock.Object.addSheetBehind(table.Name);
                    spreadSheetMock.mock.Object.setContentOfSheet(numOfSheet, table.TableArray, true);
                }
                if (saveAfterEveryEdit) spreadSheetMock.mock.Object.saveSpreadsheet();
            }
        }

        private void addTableToFile(Table table, int numOfSheet)
        {
            spreadSheetMock.addRangeToSheet(numOfSheet, (string[,])table.TableArray);
        }
    }
}
