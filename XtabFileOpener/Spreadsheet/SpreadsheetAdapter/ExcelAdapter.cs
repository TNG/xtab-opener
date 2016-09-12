using System;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using XtabFileOpener.TableContainer.SpreadsheetTableContainer.ExcelTableContainer;

namespace XtabFileOpener.Spreadsheet.SpreadsheetAdapter
{
    /// <summary>
    /// enables interacting with Excel
    /// </summary>
    internal class ExcelAdapter : SpreadsheetAdapter
    {
        private event applyChanges saveEvent;
        event applyChanges SpreadsheetAdapter.save
        {
            add { saveEvent += value; }
            remove { saveEvent -= value; }
        }

        private event spreadsheetClosed closedEvent;
        event spreadsheetClosed SpreadsheetAdapter.closed
        {
            add { closedEvent += value; }
            remove { closedEvent -= value; }
        }

        private FileInfo tmpFile;
        private Application excel;
        private Workbook workbook;

        private ExcelTableContainer excelTableContainer;

        internal ExcelAdapter() { }

        public void createSpreadsheet(string name)
        {
            excel = new Application();
            excel.DisplayAlerts = false;
            workbook = excel.Workbooks.Add(Type.Missing);
            excel.Caption = name;
        }

        public void closeSpreadsheet()
        {
            workbook.Close();
        }

        public void destroySpreadsheet()
        {
            //tmpFile.Delete();
            //necessary to stop process
            deleteObject(workbook);
            deleteObject(excel);
        }

        private void deleteObject(Object o)
        {
            try { while (Marshal.ReleaseComObject(o) > 0) ; }
            catch { }
            finally { o = null; }
        }

        public void saveSpreadsheet() { }

        private void saveWorkbookAs()
        {
            tmpFile = new FileInfo(Path.GetTempPath() + Path.GetRandomFileName());
            workbook.SaveAs(tmpFile);
        }

        public void startListeningToSaving()
        {
            excel.WorkbookBeforeSave += new AppEvents_WorkbookBeforeSaveEventHandler(onSave);   
        }

        private void onSave(Workbook savedWorkbook, bool SaveAsUI, ref bool Cancel)
        {
            raiseSaveEvent();
            //cancel saving, so that no xls-file is created
            Cancel = true;
        }

        public void startListeningToClosing()
        {
            excel.WorkbookDeactivate += onClose;
        }

        private void onClose(Workbook wb)
        {
            new Thread(waitForRealClosing).Start();
        }

        /// <summary>
        /// waits until the Excel workbook is really closed 
        /// (when the WorkbookDeactivate-event is fired, the application is not yet closed completely)
        /// </summary>
        private void waitForRealClosing()
        {
            while (excel.Workbooks.Count != 0)
                Thread.Sleep(10);
            raiseClosedEvent();
        }

        public void show()
        {
            excel.Visible = true;
            workbook.Activate();
        }

        public int SheetCount
        {
            get { return workbook.Sheets.Count; }
        }

        public int addSheetBehind(string name)
        {
            Worksheet sheet = workbook.Sheets.Add(After: workbook.Sheets.get_Item(workbook.Sheets.Count));
            sheet.Name = name;
            return workbook.Sheets.Count - 1;
        }

        public void setContentOfSheet(int number, object[,] content, bool autosizeColumns)
        {
            Worksheet sheet = workbook.Sheets[number + 1];
            Range startCell = sheet.Cells.get_Item(1, 1);
            Range endCell = sheet.Cells.get_Item(content.GetLength(0), content.GetLength(1));
            Range rangeToSet = sheet.get_Range(startCell, endCell);
            rangeToSet.Value = content;
            sheet.Cells.NumberFormat = "@";
            if (autosizeColumns)
                sheet.UsedRange.Columns.AutoFit();
        }

        public void deleteSheet(int number)
        {
            ((Worksheet)workbook.Sheets[number + 1]).Delete();
        }

        public void activateSheet(int number)
        {
            ((Worksheet)workbook.Sheets[number + 1]).Select(Type.Missing);
        }

        public void renameSheet(int number, string name)
        {
            ((Worksheet)workbook.Sheets[number + 1]).Name = name;
        }

        public TableContainer.TableContainer tableContainer
        {
            get { return excelTableContainer;  }
        }

        public void createTableContainer()
        {
            excelTableContainer = new ExcelTableContainer(workbook.Name, workbook.Sheets);
        }

        private void raiseSaveEvent()
        {
            if (saveEvent != null) saveEvent(tableContainer);
        }

        private void raiseClosedEvent()
        {
            if (closedEvent != null) closedEvent();
        }
    }
}
