using System;
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
            /*
             * "The only difference between this property and the Value property is that the Value2 property
             *  doesn’t use the Currency and Date data types. You can return values formatted with these 
             *  data types as floating-point numbers by using the Double data type."
             * */
            tableArray = cells.Value;
            firstRowContainsColumnNames = true;
        }
    }
}
