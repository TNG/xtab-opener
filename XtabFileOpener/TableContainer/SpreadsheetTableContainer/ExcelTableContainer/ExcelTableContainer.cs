using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace XtabFileOpener.TableContainer.SpreadsheetTableContainer.ExcelTableContainer
{
    /// <summary>
    /// Implementation of TableContainer, that contains Excel database tables
    /// </summary>
    internal class ExcelTableContainer : TableContainer
    {
        internal ExcelTableContainer(string name, Sheets sheets) : base(name)
        {
            this.sheets = sheets;
        }

        private Sheets sheets;
        public override IEnumerator<Table> GetEnumerator()
        {
            //take only the used range in the Excel sheet
            foreach (Worksheet sheet in sheets)
                yield return new ExcelTable(sheet.Name, sheet.UsedRange);
        }

        public override int Count
        {
            get { return sheets.Count; }
        }
    }
}
