using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XtabFileOpener.TableContainer
{
    /// <summary>
    /// Represents a container of database tables; the container is not static, that means it only contains references on tables.
    /// the consequence is that iterating over the tables at different times can yield different tables
    /// </summary>
    public abstract class TableContainer : IEnumerable<Table>
    {
        public TableContainer(string name)
        {
            this.name = name;
        }
        
        private string name;
        public string Name
        {
            get { return name; }
        }

        public abstract IEnumerator<Table> GetEnumerator();

        /// <summary>
        /// enables iterating over the tables of this container
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        /// <summary>
        /// number of the tables
        /// </summary>
        public abstract int Count
        {
            get;
        }

        /// <summary>
        /// checks whether the given TableContainer is equals to this one concerning the name and content of every table
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public override bool Equals(object other)
        {
            TableContainer otherContainer = other as TableContainer;
            var bothCons = this.Zip(otherContainer, (t1, t2) => new { Table1 = t1, Table2 = t2 });
            if (this.Count != otherContainer.Count)
            {
                return false;
            }
            foreach (var com in bothCons)
            {
                if (!com.Table1.Equals(com.Table2)) return false;
            }
            return true;
        }
    }
}
