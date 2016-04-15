using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using XtabFileOpener.XtabFile;

namespace XtabFileOpener.TableContainer.XmlTableContainer
{
    /// <summary>
    /// Implementation of TableContainer, that contains XML database tables 
    /// </summary>
    public class XmlTableContainer : TableContainer
    {
        public XmlTableContainer(string name, IEnumerable<XElement> tables) : base(name)
        {
            this.tables = tables;
        }

        private IEnumerable<XElement> tables;
        public override IEnumerator<Table> GetEnumerator()
        {
            foreach (XElement table in tables)
                yield return new XmlTable(table.Attribute(XtabFormat.table_name).Value, 
                    table.Elements(XtabFormat.row), table.Elements(XtabFormat.column));
        }

        public override int Count
        {
            get { return tables.Count(); }
        }

        public static TableContainer createFromStream(Stream stream, string name)
        {
            XDocument doc = XDocument.Load(stream);
            TableContainer result = new XmlTableContainer(name, doc.Root.Elements());
            return result;
        }

        public static TableContainer createFromFile(FileInfo file)
        {
            XDocument doc = XDocument.Load(file.FullName);
            TableContainer result = new XmlTableContainer(file.Name, doc.Root.Elements());
            return result;
        }
    }
}
