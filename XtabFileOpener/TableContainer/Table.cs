using System;
using System.Text;
using System.Threading.Tasks;

namespace XtabFileOpener.TableContainer
{
    /// <summary>
    /// represents a database tables
    /// </summary>
    public abstract class Table
    {
        /// <summary>
        /// one based array with the whole content of the table, including columns
        /// </summary>
        protected object[,] tableArray;

        /// <summary>
        /// defines whether the TableArray contains the column names in tis first row
        /// </summary>
        public bool firstRowContainsColumnNames { get; protected set; }

        public Table(string name)
        {
            this.name = name;
        }

        private string name;
        public string Name
        {
            get { return name; }
        }

        /// <summary>
        /// number of columns of the table
        /// </summary>
        public int Width
        {
            get { return tableArray.GetLength(1); }
        }

        /// <summary>
        /// number of rows of the table including the row with the column names
        /// </summary>
        public int Height
        {
            get { return tableArray.GetLength(0); }
        }

        /// <summary>
        /// get an array with one-based indexing
        /// </summary>
        /// <returns></returns>
        public virtual object[,] TableArray
        {
            get { return tableArray; }
        }

        /// <summary>
        /// checks whether the table is empty
        /// </summary>
        public bool Empty
        {
            get { return Height == 0 || Width == 0; }
        }

        public override bool Equals(object obj)
        {
            Table tab = obj as Table;
            object[,] arr1 = TableArray;
            object[,] arr2 = tab.TableArray;
            return name.Equals(tab.name) && ListTableContainer.Array2D.areEqual(arr1, arr2);
        }
    }
}
