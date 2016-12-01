using System;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Reflection;
using XtabFileOpener.TableContainer;
using XtabFileOpener.TableContainer.XmlTableContainer;


namespace XtabFileOpener.XtabFile
{
    /// <summary>
    /// represents an xtab-file and provides methods to work with it
    /// </summary>
    internal class XtabFile
    {
        public const string suffix = ".xtab";

        private FileInfo file;

        private XDocument document;

        private readonly TableContainer.TableContainer tableContainer;

        private const string invalideFileMessage = "The given file is no valid xtab-file";

        private const string schemaPath = "XtabFileOpener.Res.xtabSchema.xsd";

        internal XtabFile(string fullFileName)
        {
            file = new FileInfo(fullFileName);
            createXDocument();
            tableContainer = new XmlTableContainer(file.Name, document.Root.Elements());
        }

        private void createXDocument()
        {
            //check wether the file is empty
            if (file.Length == 0)
                createNewEmptyDocument();
            else
                loadExistingDocument();
        }

        private void createNewEmptyDocument()
        {
            XElement root = new XElement(XtabFormat.dataset);
            root.Add(new XElement(XtabFormat.table, new XAttribute(XtabFormat.table_name, XtabFormat.default_table_name)));
            document = new XDocument(new XDeclaration(XtabFormat.head[0], XtabFormat.head[1], XtabFormat.head[2]), root);
        }

        private void loadExistingDocument()
        {
            XmlSchemaSet schemas = new XmlSchemaSet();
            schemas.Add("", XmlReader.Create(new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream(schemaPath))));
            document = XDocument.Load(file.FullName);
            document.Validate(schemas, (o, e) =>
            {
                document = null;
                file = null;
                throw new ArgumentException(invalideFileMessage);
            });
        }

        /// <summary>
        /// TableContainer, that corresponds to the xtab-file
        /// </summary>
        internal TableContainer.TableContainer TableContainer
        {
            get { return tableContainer; }
        }

        internal void saveChanges(TableContainer.TableContainer tc)
        {
            XElement root = new XElement(XtabFormat.dataset);
            XDocument newDoc = new XDocument(new XDeclaration(XtabFormat.head[0], XtabFormat.head[1], XtabFormat.head[2]), root);
            foreach (Table table in tc)
                addTableToElement(root, table);

            saveDocument(newDoc);
        }

        private static void addTableToElement(XElement root, Table table)
        {
            XElement xTable = new XElement(XtabFormat.table, new XAttribute(XtabFormat.table_name, table.Name));
            root.Add(xTable);
            object[,] array = table.TableArray;
            if (array != null)
            {
                int i = array.GetLowerBound(0);
                if (table.firstRowContainsColumnNames)
                {
                    for (int j = array.GetLowerBound(1); j <= array.GetUpperBound(1); j++)
                        addColumnToElement(xTable, array[i, j]);
                    i++;
                }

                for (; i <= array.GetUpperBound(0); i++)
                    addRowToElement(xTable, array, i);
            }
        }

        private static void addRowToElement(XElement xTable, object[,] array, int row)
        {
            XElement xRow = new XElement(XtabFormat.row);
            xTable.Add(xRow);
            for (int j = array.GetLowerBound(1); j <= array.GetUpperBound(1); j++)
                addValueToElement(xRow, array[row, j]);
        }

        private static void addValueToElement(XElement xRow, object value)
        {
            xRow.Add(sanitizeRowValue(value));
        }

        private static object sanitizeRowValue(object value)
        {
            char[] charactersThatAskForCdata = { '\u0020', '\n', '\r', '\t', '&', '<' };
            if (value == null)
            {
                return new XElement(XtabFormat.nullvalue);
            }
            if (value.ToString().IndexOfAny(charactersThatAskForCdata) > -1)
            {
                return new XElement(XtabFormat.value, new XCData(value.ToString()));
            }
            return new XElement(XtabFormat.value, value.ToString());
        }

        private static void addColumnToElement(XElement xTable, object column)
        {
            xTable.Add(new XElement(XtabFormat.column, column == null ? "" : column.ToString()));
        }

        private void saveDocument(XDocument newDoc)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = XtabFormat.encoding;
            settings.Indent = true;
            using (var writer = XmlWriter.Create(file.FullName, settings)) //new XmlTextWriter(file.FullName, XtabFormat.encoding))
            {
                newDoc.Save(writer);
            }
        }
    }
}
