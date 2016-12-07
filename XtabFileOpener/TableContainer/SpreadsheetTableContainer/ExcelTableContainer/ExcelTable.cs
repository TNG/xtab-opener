using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.CompilerServices;

[assembly: InternalsVisibleTo("XtabFileOpenerTest")]
namespace XtabFileOpener.TableContainer.SpreadsheetTableContainer.ExcelTableContainer
{
    /// <summary>
    /// Implementation of Table, that manages an Excel table.
    /// Cells that are date formatted in Excel will be converted to ISO strings.
    /// </summary>
    internal class ExcelTable : Table
    {
        internal ExcelTable(string name, Range cells) : base(name)
        {
            // TODO: This if-clause looks like it could need some refactoring.
            // We need this extra logic because cells.Value is of type string or double 
            // if the range only contains a single cell. 
            if (cells.Value is object[,]) {
                tableArray = cells.Value;
            } else {                 
                tableArray = new object[1, 1];
                tableArray[0, 0] = cells.Value;
            }
            firstRowContainsColumnNames = true;
            ConvertDateTimesToIsoFormat();
        }

        private void ConvertDateTimesToIsoFormat()
        {
            string iso_time_format = "{0:yyyy-MM-dd HH:mm:ss}";
            string iso_time_format_with_miliseconds = "{0:yyyy-MM-dd HH:mm:ss.fffffff}00";            
            if (tableArray == null) { // Happens when table is empty
                return;
            }
            for (int i = tableArray.GetLowerBound(0); i <= tableArray.GetUpperBound(0); i++)
            {
                for (int j = tableArray.GetLowerBound(1); j <= tableArray.GetUpperBound(1); j++)
                {
                    object cell = tableArray[i, j];
                    if (cell is DateTime && cell != null)
                    {
                        DateTime dt = (DateTime) cell;
                        string format = (dt.Millisecond > 0) ? iso_time_format_with_miliseconds : iso_time_format;
                        tableArray[i, j] = String.Format(format, dt);
                    }
                }
            }
        }
    }
}
