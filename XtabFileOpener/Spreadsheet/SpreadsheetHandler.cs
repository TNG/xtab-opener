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
        internal SpreadsheetAdapter.SpreadsheetAdapter spreadsheetAdapter
        {
            get; private set;
        }

        internal SpreadsheetHandler(TableContainer.TableContainer tableContainer, SpreadsheetAdapter.SpreadsheetAdapter spreadsheetAdapter, bool autosizeColumns)
        {
            this.spreadsheetAdapter = spreadsheetAdapter;
            this.spreadsheetAdapter.createSpreadsheet(tableContainer.Name);
            createWorkbookFromTableContainer(tableContainer, autosizeColumns);
        }
        
        internal void openSpreadsheet()
        {
            spreadsheetAdapter.saveSpreadsheet();
            spreadsheetAdapter.startListeningToSaving();
            spreadsheetAdapter.closed += onClosed;
            spreadsheetAdapter.startListeningToClosing();
            spreadsheetAdapter.show();
        }

        private void createWorkbookFromTableContainer(TableContainer.TableContainer tableContainer, bool autoSizeColumns)
        {
            //necessary if Excel opens several sheets at the beginning
            closeExistingSheetsExceptOne();

            bool standardSheetExists = spreadsheetAdapter.SheetCount >= 1;
            //necessary to rename the standard sheet with an unusual name, so that no table of the xtab-file has the same name
            if (standardSheetExists)
                spreadsheetAdapter.renameSheet(0, "DefaultSheetWithInimitableName");

            foreach (Table table in tableContainer)
                addTableToWorkbook(table, autoSizeColumns);

            //should never occur: if there is no table, then add default table
            if (tableContainer.Count == 0)
                spreadsheetAdapter.addSheetBehind(XtabFormat.default_table_name);

            if (standardSheetExists)
                spreadsheetAdapter.deleteSheet(0);

            spreadsheetAdapter.activateSheet(0);

            spreadsheetAdapter.createTableContainer();
        }

        private void closeExistingSheetsExceptOne()
        {
            while (spreadsheetAdapter.SheetCount >= 2)
                spreadsheetAdapter.deleteSheet(0);
        }

        private void addTableToWorkbook(Table table, bool autosize)
        {
            int number = spreadsheetAdapter.addSheetBehind(table.Name);
            if (!table.Empty)
                spreadsheetAdapter.setContentOfSheet(number, table.TableArray, autosize);
        }

        private bool closed = false;
        private void onClosed()
        {
            Monitor.Enter(this);
            closed = true;
            Monitor.Pulse(this);
            Monitor.Exit(this);
            spreadsheetAdapter.destroySpreadsheet();
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
