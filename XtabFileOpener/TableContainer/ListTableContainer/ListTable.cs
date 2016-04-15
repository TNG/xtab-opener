using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XtabFileOpener.TableContainer.ListTableContainer
{
    internal class ListTable : Table
    {
        internal ListTable(string name, string[,] tableArray):base(name)
        {
            this.tableArray = tableArray;
        }
    }
}
