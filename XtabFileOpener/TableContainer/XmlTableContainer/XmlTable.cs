using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using XtabFileOpener.TableContainer.ListTableContainer;

namespace XtabFileOpener.TableContainer.XmlTableContainer
{
    /// <summary>
    /// Implementation of Table, that manages an XML-based table
    /// </summary>
    public class XmlTable : Table
    {
        public XmlTable(string name, IEnumerable<XElement> rows, IEnumerable<XElement> columns) : base(name)
        {
            createTableArray(rows, columns);
        }

        private void createTableArray(IEnumerable<XElement> rows, IEnumerable<XElement> columns)
        {
            firstRowContainsColumnNames = columns.Count() > 0;

            int height = rows.Count() + (firstRowContainsColumnNames ? 1 : 0);
            int width = getMaxWidth(rows, columns.Count());

            tableArray = new string[height, width];

            int i = tableArray.GetLowerBound(0);
            int j = tableArray.GetLowerBound(1);
            if (firstRowContainsColumnNames)
            {
                foreach (XElement column in columns)
                {
                    tableArray[i, j] = column.Value;
                    j++;
                }
                i++;
            }

            foreach (XElement row in rows)
            {
                j = tableArray.GetLowerBound(1);
                foreach (XElement value in row.Elements())
                {
                    tableArray[i, j] = value.Value;
                    j++;
                }
                i++;
            }

            Array2D.replaceNullByEmptyString(tableArray);
        }

        private int getMaxWidth(IEnumerable<XElement> rows, int columnsCount)
        {
            int maxLength = 0;
            foreach (XElement row in rows)
            {
                int actLength = row.Elements().Count();
                if (actLength > maxLength) maxLength = actLength;
            }
            if (columnsCount > maxLength)
                maxLength = columnsCount;
            return maxLength;
        }
    }
}
