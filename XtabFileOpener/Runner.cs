using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XtabFileOpener.Spreadsheet.SpreadsheetAdapter;

namespace XtabFileOpener
{
    /// <summary>
    /// runs the whole xtabOpener-program by containing a main-method
    /// </summary>
    public class Runner
    {
        /// <summary>
        /// tries to open the file with the path at the first position in the argument array
        /// </summary>
        /// <param name="args">
        /// array with the path of the xtab-file at the first position
        /// </param>
        public static void Main(String[] args)
        {
            if (args.Length == 1)
            {
                try
                {
                    XtabFileOpener opener = new XtabFileOpener(args[0]);
                    opener.openFile(new ExcelAdapter());
                    opener.waitForClosing();
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
