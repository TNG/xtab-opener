﻿using System;
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
                    showErrorDialog(ex);
                }
            }
            else if (args.Length == 0)
            {
                MessageBox.Show("Please give a single filename as parameter", "Missing parameter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("Please give only a single filename as parameter. I received multiple.", "More than one parameter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// display an error message box containing the stack trace of an (unhandled) exception
        /// <param name="ex">
        /// The exception to render
        /// </param>
        /// <param name="when">
        /// string describing the action that lead to the exception (for example 'during opening the file')
        /// </param>
        /// </summary>
        public static void showErrorDialog(Exception ex, String when="")
        {            
            MessageBox.Show(new System.Windows.Forms.Form() { TopMost = true },             
                "The unhandled exception " + ex.Message + " occured in the XtabFileOpener " + when +
                "\n\nClick in this popup and You can use ctrl + c to copy the following stacktrace.\n\n" +
                ex.ToString(),
                "Error during saving",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        /// <summary>
        /// display an error message box containing the stack trace of an (unhandled) exception
        /// use this method if an excel instance is already present toe prevent the dialog from appearing in the background
        /// <param name="excel">
        /// the excel instance that should display the dialog
        /// </param>
        /// <param name="ex">
        /// The exception to render
        /// </param>
        /// <param name="when">
        /// string describing the action that lead to the exception (for example 'during opening the file')
        /// </param>
        /// </summary>
        public static void showErrorDialog(Microsoft.Office.Interop.Excel.Application excel, Exception ex, String when = "")
        {
            NativeWindow xlMain = new NativeWindow();
            xlMain.AssignHandle(new IntPtr(excel.Hwnd));
            MessageBox.Show(xlMain,
                "The unhandled exception " + ex.Message + " occured in the XtabFileOpener " + when +
                "\n\nClick in this popup and You can use ctrl + c to copy the following stacktrace.\n\n" +
                ex.ToString(),
                "Error during saving",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            xlMain.ReleaseHandle();
        }

        /// <summary>
        /// display an error message box containing the stack trace of an (unhandled) exception
        /// use this method if an excel instance is already present toe prevent the dialog from appearing in the background
        /// <param name="excel">
        /// the excel instance that should display the dialog
        /// </param>
        /// <param name="message">
        /// the message body
        /// </param>
        /// <param name="title">
        /// the title of the box
        /// </param>
        /// </summary>
        public static void showMessageInExcel(Microsoft.Office.Interop.Excel.Application excel, String message, String title)
        {
            NativeWindow xlMain = new NativeWindow();
            xlMain.AssignHandle(new IntPtr(excel.Hwnd));
            MessageBox.Show(xlMain, message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            xlMain.ReleaseHandle();    
        }

    }

}
