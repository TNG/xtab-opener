using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XtabFileOpener.TableContainer.ListTableContainer
{
    internal class ListTable : Table
    {
        /// <summary>
        /// Table consisting of a simple two dimensional array
        /// </summary>
        /// <param name="name"></param>
        /// <param name="tableArray"></param>
        internal ListTable(string name, string[,] tableArray):base(name)
        {
            this.tableArray = tableArray;
        }
    }
}
