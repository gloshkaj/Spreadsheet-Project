using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace Spreadsheet
{
    public static class Errors
    {
        public static ListBox listbox;
        /*
        NAME

            RecordError - Records the collected error message in the list box.

        SYNOPSIS

            public static void RecordError(string a_msg);

        DESCRIPTION

            This function will collect the error message and put it in the list box

        RETURNS

           None
        */
        public static void RecordError(string a_msg)
        {
            listbox = Application.OpenForms["frmSpreadsheet"].Controls["lstErrors"] as ListBox;
            // Access form control and record collected error message
            listbox.Items.Add(a_msg);
        }
        /*
        NAME

            DisplayErrors - Displays all the error messages in separate message boxes upon exiting the program.

        SYNOPSIS

            public static void DisplayErrors();

        DESCRIPTION

            This function will display each error message in a separate message box upon exiting the program.

        RETURNS

           None

        */
        public static void DisplayErrors()
        {
            // Access form control
            listbox = Application.OpenForms["frmSpreadsheet"].Controls["lstErrors"] as ListBox;
            if (listbox.Items.Count == 0) // If list box has no elements in it when program is exited print success message and exit program
            {
                MessageBox.Show("User did not make any mistakes.");
                return;
            }
            // Otherwise print each item from listbox in separate messageboxes
            foreach (string item in listbox.Items)
            {
                MessageBox.Show(item);
            }
        }
    }
}
