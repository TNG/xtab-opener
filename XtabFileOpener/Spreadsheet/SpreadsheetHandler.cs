using System;
using System.Text;
using System.Threading;
using XtabFileOpener.TableContainer;
using XtabFileOpener.XtabFile;

namespace XtabFileOpener.Spreadsheet
{
    /// <summary>
    /// enables open the tables with a program that can show and edit tables
    /// </summary>
    internal class SpreadsheetHandler
    {
        internal SpreadsheetAdapter.SpreadsheetAdapter spreadsheet
        {
            get; private set;
        }

        internal SpreadsheetHandler(TableContainer.TableContainer tableContainer, SpreadsheetAdapter.SpreadsheetAdapter spreadsheet, bool autosizeColumns)
        {
            this.spreadsheet = spreadsheet;
            spreadsheet.createSpreadsheet(tableContainer.Name);
            createWorkbookFromTableContainer(tableContainer, autosizeColumns);
        }
        
        internal void openSpreadsheet()
        {
            spreadsheet.saveSpreadsheet();
            //spreadsheet.save += onSave;
            spreadsheet.startListeningToSaving();
            spreadsheet.closed += onClosed;
            spreadsheet.startListeningToClosing();
            spreadsheet.show();
        }

        private void createWorkbookFromTableContainer(TableContainer.TableContainer tableContainer, bool autoSizeColumns)
        {
            //necessary if Excel opens several sheets at the beginning
            closeExistingSheetsExceptOne();

            bool standardSheetExists = spreadsheet.SheetCount >= 1;
            //necessary to rename the standard sheet with an unusual name, so that no table of the xtab-file has the same name
            if (standardSheetExists)
                spreadsheet.renameSheet(0, "DefaultSheetWithInimitableName");

            foreach (Table table in tableContainer)
                addTableToWorkbook(table, autoSizeColumns);

            //should never occur: if there is no table, then add default table
            if (tableContainer.Count == 0)
                spreadsheet.addSheetBehind(XtabFormat.default_table_name);

            if (standardSheetExists)
                spreadsheet.deleteSheet(0);

            spreadsheet.activateSheet(0);

            spreadsheet.createTableContainer();
        }

        private void closeExistingSheetsExceptOne()
        {
            while (spreadsheet.SheetCount >= 2)
                spreadsheet.deleteSheet(0);
        }

        private void addTableToWorkbook(Table table, bool autosize)
        {
            int number = spreadsheet.addSheetBehind(table.Name);
            if (!table.Empty)
                spreadsheet.setContentOfSheet(number, table.TableArray, autosize);
        }

        private bool closed = false;
        private void onClosed()
        {
            Monitor.Enter(this);
            closed = true;
            Monitor.Pulse(this);
            Monitor.Exit(this);
            spreadsheet.destroySpreadsheet();
        }

        /// <summary>
        /// lets the current thread wait until the Excel workbook is closed
        /// </summary>
        public void waitForClosing()
        {
            while (!closed)
                wait();
        }

        private void wait()
        {
            Monitor.Enter(this);
            Monitor.Wait(this);
            Monitor.Exit(this);
        }
    }
}
