using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace XtabFileOpener.TableContainer.SpreadsheetTableContainer.ExcelTableContainer
{
    /// <summary>
    /// Implementation of Table, that manages an Excel table
    /// </summary>
    internal class ExcelTable : Table
    {
        internal ExcelTable(string name, Range cells) : base(name)
        {
            tableArray = cells.Value2;
            firstRowContainsColumnNames = true;
        }
    }
}
