using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XtabFileOpener.XtabFile
{
    /// <summary>
    /// contains constants, which are names of XML-elements of an xtab-file
    /// </summary>
    public class XtabFormat
    {
        public static readonly Encoding encoding = new UTF8Encoding(false);
        /// <summary>
        /// first line of the xtab file
        /// </summary>
        public static readonly string[] head = new string[] { "1.0", encoding.BodyName, "" };
        public const string dataset = "dataset";
        public const string table = "table";
        public const string table_name = "name";
        public const string default_table_name = "Default";
        public const string column = "column";
        public const string row = "row";
        public const string value = "value";
        public const string nullvalue = "null";
    }
}
