using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XtabFileOpener.TableContainer;

namespace XtabFileOpener.Spreadsheet.SpreadsheetAdapter
{
    /// <summary>
    /// delegate that can be called when a spreadsheet was changed
    /// </summary>
    /// <param name="tableContainer">TableContainer with the whole data</param>
    public delegate void applyChanges(TableContainer.TableContainer tableContainer);
    /// <summary>
    /// delegate that can be called when the spreadsheet is closed
    /// </summary>
    public delegate void spreadsheetClosed();

    /// <summary>
    /// describes the composition of a class that allows interacting with a spreadsheet
    /// </summary>
    public interface SpreadsheetAdapter
    {
        /// <summary>
        /// event that is raised when the spreadsheet is saved
        /// </summary>
        event applyChanges save;

        /// <summary>
        /// event that is raised when the spreadsheet is closed
        /// </summary>
        event spreadsheetClosed closed;

        /// <summary>
        /// creates a new spreadsheet
        /// </summary>
        /// <param name="name">name of the spreadsheet</param>
        void createSpreadsheet(string name);

        /// <summary>
        /// closes the spreadsheet
        /// </summary>
        void closeSpreadsheet();

        /// <summary>
        /// destroys the spreadsheet permanently
        /// </summary>
        void destroySpreadsheet();

        /// <summary>
        /// saves the spreadsheet to a file
        /// </summary>
        void saveSpreadsheet();

        /// <summary>
        /// starts listening to the the workbook being saved
        /// </summary>
        void startListeningToSaving();

        /// <summary>
        /// starts listening to the workbook being closed
        /// </summary>
        void startListeningToClosing();

        /// <summary>
        /// sets the workbook visible to the user
        /// </summary>
        void show();

        /// <summary>
        /// number of sheets in this spreadsheet
        /// </summary>
        int SheetCount { get; }

        /// <summary>
        /// adds a new sheet to the workbook
        /// </summary>
        /// <param name="name">name of the new sheet</param>
        /// <returns>number of the new sheet, beginning at 0</returns>
        int addSheetBehind(string name);

        /// <summary>
        /// sets the content of a sheet in the workbook; if the sheet is not empty, all concerned cells are overridden
        /// </summary>
        /// <param name="number">number of the sheet, beginning at 0</param>
        /// <param name="content">new content of the sheet</param>
        /// <param name="autosizeColumns">whether the columns in the sheet should adapt their width to the new content</param>
        void setContentOfSheet(int number, object[,] content, bool autosizeColumns);

        /// <summary>
        /// deletes a sheet
        /// </summary>
        /// <param name="number">number of the sheet that should be deleted, beginning at 0</param>
        void deleteSheet(int number);

        /// <summary>
        /// sets a sheet to the active one
        /// </summary>
        /// <param name="number">number of the sheet that should be activated, beginning at 0</param>
        void activateSheet(int number);

        /// <summary>
        /// renames a sheet in the workbook
        /// </summary>
        /// <param name="number">number of the sheet, that should be renamed</param>
        /// <param name="name">new name of the sheet</param>
        void renameSheet(int number, string name);

        /// <summary>
        /// TableContainer that represents this spreadsheet
        /// </summary>
        TableContainer.TableContainer tableContainer { get; }

        /// <summary>
        /// creates the TableContainer
        /// </summary>
        void createTableContainer();
    }
}
