using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XtabFileOpener.TableContainer.ListTableContainer;

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
        public object[,] tableArray { get; protected set; }

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
        /// creates a string array of this table, including the column names
        /// </summary>
        /// <returns></returns>
        //[Obsolete]
        //public string[,] toArray()
        //{
        //    string[,] result = new string[Height, maxWidth];

        //    int i = 0;
        //    int j = 0;
        //    foreach (string value in Columns)
        //    {
        //        result[i, j] = value;
        //        j++;
        //    }

        //    i += Columns.Exists ? 1 : 0;
        //    foreach (Row row in this)
        //    {
        //        j = 0;
        //        foreach (string value in row)
        //        {
        //            result[i, j] = value;
        //            j++;
        //        }
        //        i++;
        //    }
        //    Array2D.replaceNullByEmptyString(result);
        //    return result;
        //}

        /// <summary>
        /// get an array with one-based indexing
        /// </summary>
        /// <returns></returns>
        public virtual object[,] getTableArray()
        {
            return tableArray;
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
            object[,] arr1 = getTableArray();
            object[,] arr2 = tab.getTableArray();
            return ListTableContainer.Array2D.areEqual(arr1, arr2);
        }
    }
}
