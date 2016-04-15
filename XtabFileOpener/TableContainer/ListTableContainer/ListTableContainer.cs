using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XtabFileOpener.TableContainer.ListTableContainer
{
    /// <summary>
    /// TableContainer that manages its tables using lists
    /// </summary>
    public class ListTableContainer : TableContainer
    {
        private IEnumerable<KeyValuePair<string, string[,]>> tables;

        /// <summary>
        /// </summary>
        /// <param name="_name">name of the TableContainer</param>
        /// <param name="_tables">tables of the TableContainer</param>
        public ListTableContainer(string name, IEnumerable<KeyValuePair<string, string[,]>> tables) : base(name)
        {
            this.tables = tables;
        }

        public override int Count
        {
            get { return tables.Count(); }
        }

        public override IEnumerator<Table> GetEnumerator()
        {
            foreach (KeyValuePair<string, string[,]> table in tables)
                yield return new ListTable(table.Key, table.Value);
        }
    }
}
