using System;
using System.Text;
using XtabFileOpener.Spreadsheet;
using XtabFileOpener.Spreadsheet.SpreadsheetAdapter;
using XtabFileOpener.XtabFile;

namespace XtabFileOpener
{
    /// <summary>
    /// enables opening an xtab-file
    /// </summary>
    public class XtabFileOpener
    {
        private readonly XtabFile.XtabFile fileHandler;
        private SpreadsheetHandler spreadsheetHandler;

        private const string noValidFileExceptionMessage = "You can not open the file without setting a valid xtab-file";

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fullFileName">full file-name of the xtab-file</param>
        public XtabFileOpener(string fullFileName)
        {
            fileHandler = new XtabFile.XtabFile(fullFileName);
        }

        /// <summary>
        /// tries to open the initialized xtab-file
        /// </summary>
        /// <param name="spreadsheetAdapter">SpreadsheetAdapter, that the xtab-file should be opened with</param>
        public void openFile(SpreadsheetAdapter spreadsheetAdapter)
        {
            spreadsheetHandler = new SpreadsheetHandler(fileHandler.TableContainer, spreadsheetAdapter, true);
            spreadsheetHandler.spreadsheet.save += fileHandler.saveChanges;
            spreadsheetHandler.openSpreadsheet();
        }

        /// <summary>
        /// waits for the xtab editor being closed by the user
        /// </summary>
        public void waitForClosing()
        {
            spreadsheetHandler.waitForClosing();
        }
    }
}
