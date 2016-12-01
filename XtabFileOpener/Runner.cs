using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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
            } else if (args.Length == 0)
            {
                System.Windows.Forms.MessageBox.Show("Please give a single filename as parameter", "Missing parameter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } else
            {
                System.Windows.Forms.MessageBox.Show("Please give only a single filename as parameter. I received multiple.", "More than one parameter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
