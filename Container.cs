using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
// Code adapted from http://stackoverflow.com/questions/1032775/sorting-mixed-numbers-and-strings
public class MixedNumbersAndStringsComparer : IComparer<string>
{
    /// <summary>
    /// Compare: sorts numbers by numbers and strings by strings.
    /// </summary>
    /// <param name="x"></param>
    /// <param name="y"></param>
    /// <returns>Dictionary order if not numeric. Numeric order otherwise</returns>
    public int Compare(string a_val1, string a_val2)
    {
        double xVal, yVal;
        // If values are not doubles sort them the string way. Otherwise sort numerically.
        if (double.TryParse(a_val1, out xVal) && double.TryParse(a_val2, out yVal))
            return xVal.CompareTo(yVal);
        else
            return string.Compare(a_val1, a_val2);
    }
}
// End of adapted code
namespace Spreadsheet
{
    class Container
    {
        private object m_value; // Value to insert
        private int m_row; // Row number
        private int m_cell; // Column number
        public Container()
        {
            m_value = "";
            m_row = m_cell = 0;
        }
        /*
NAME

Insert - Inserts the value in the right hand text box into the cell specified by the left hand text box.

SYNOPSIS

public void Insert (object a_newVal, int a_cell, int a_row);

DESCRIPTION

This function will insert the data from the right hand text box into the cell specified by the left hand text box.
RETURNS

None. Reports error if cell to insert into is non empty.
*/
        public void Insert(object a_newVal, int a_cell, int a_row)
        {
            // Access datagridview
            DataGridView dgv = Application.OpenForms["frmSpreadsheet"].Controls["dgvSpreadsheet"] as DataGridView;
            //Report error in listbox if cell is not empty
            if (Convert.ToString(dgv[a_cell, a_row].Value) != String.Empty)
            {
                MessageBox.Show("Cannot insert into non-empty cell");
                Errors.RecordError("Cannot insert into non-empty cell");
                return;
            }
            m_value = a_newVal;
            m_cell = a_cell;
            m_row = a_row;
            // Record the value in the spreadsheet at the specified cell.
            dgv[m_cell, m_row].Value = m_value;
        }
        /*
NAME

Sort - Sort in ascending dictionary order.

SYNOPSIS

public void Sort ();

DESCRIPTION

This function will iterate through the rows and columns to place the non empty cell contents into an arraylist and then sort the arraylist in dictionary order.
Removes existing elements and places them column by column starting from the first column.
RETURNS

None.
*/
        public void Sort() // Sorting in this case will be done in dictionary order. Sorting for the median in the aggregate class will be numerical order
        {
            // Hold the sorted list in an array list
            List<string> dgvVals = new List<string>();
            // Access datagridview
            DataGridView dgv = Application.OpenForms["frmSpreadsheet"].Controls["dgvSpreadsheet"] as DataGridView;
            for (int i = 0; i < dgv.Columns.Count - 1; i++)
            {
                for (int j = 0; j < dgv.Rows.Count - 1; j++)
                {
                    if (Convert.ToString(dgv[i, j].Value) != String.Empty)
                    {
                        // For each column in the spreadsheet, insert contents of each non empty cell from the column
                        // into arraylist and remove existing values
                        dgvVals.Add(Convert.ToString(dgv[i, j].Value));
                        Remove(i, j);
                    }
                    
                }
            }
            dgvVals.Sort(new MixedNumbersAndStringsComparer()); // Sorting is done in DICTIONARY/NUMERIC ORDER (ex. 103 would come before 110, 20 would be before hello world, etc.)
            // Specifically because we allow non numeric values in this case.
            int count = 0;
            // Replace first column, second column, up to nth column, with sorted list.
            for (int i = 0; i < dgvVals.Count; i++)
            {
                dgv[count, i].Value = dgvVals[i];
                if (i >= dgv.Rows.Count && i % dgv.Rows.Count == 0) // If we reached the end of a column go to the next column
                {
                    count++;
                }
            }
        }
        /*
NAME

SortDescending - Sort in descending dictionary order.

SYNOPSIS

public void SortDescending ();

DESCRIPTION

This function will iterate through the rows and columns to place the non empty cell contents into an arraylist and then sort the arraylist in descending dictionary order.
Removes existing elements and places them column by column starting from the first column.
RETURNS

None.
*/
        public void SortDescending()
        {
            // Hold reversed sorted list in arraylist
            List<string> dgvVals = new List<string>();
            DataGridView dgv = Application.OpenForms["frmSpreadsheet"].Controls["dgvSpreadsheet"] as DataGridView;
            for (int i = 0; i < dgv.Columns.Count - 1; i++)
            {
                for (int j = 0; j < dgv.Rows.Count - 1; j++)
                {
                    if (Convert.ToString(dgv[i, j].Value) != String.Empty)
                    {
                        // For each column in spreadsheet, place contents of each non empty cell in arraylist and remove existing content from spreadsheet
                        dgvVals.Add(Convert.ToString(dgv[i, j].Value));
                        Remove(i, j);
                    }

                }
            }
            dgvVals.Sort(new MixedNumbersAndStringsComparer()); // Sorting is done in DICTIONARY ORDER/Numeric (ex. 103 would come before 20, Hello World would be before 20, etc.)
            dgvVals.Reverse(); // Reverse list
            int count = 0;
            for (int i = 0; i < dgvVals.Count; i++)
            {
                dgv[count, i].Value = dgvVals[i];
                // Replace each cell in first column, second column, up to nth column, with sorted list.
                if (i >= dgv.Rows.Count && i % dgv.Rows.Count == 0) // If we reach end of column go to the next column.
                {
                    count++;
                }
            }
        }
        /*
NAME

Update - Overwrites the non-empty cell specified by the left hand textbox with the value in the right hand text box.

SYNOPSIS

public void Update (object a_newVal, int a_cell, int a_row);

DESCRIPTION

This function will update the contents of the specified non-empty cell with the data from the right hand text box
RETURNS

None. Reports error if cell to overwrite is empty.
*/
        public void Update(object a_val, int a_cell, int a_row) // If the cell is empty, report error that we cannot update an empty cell. Otherwise change the cell's value
        {
            // Access spreadsheet control
            DataGridView dgv = Application.OpenForms["frmSpreadsheet"].Controls["dgvSpreadsheet"] as DataGridView;
            m_value = a_val;
            // If empty report error that we cannot update empty cell.
            if (Convert.ToString(dgv[a_cell, a_row].Value) == String.Empty)
            {
                MessageBox.Show("Cannot update empty cell");
                Errors.RecordError("Cannot update from empty cell");
                return;
            }
            dgv[a_cell, a_row].Value = m_value;
        }
        /*
NAME

Remove - Deletes the content in the cell specified by LH text box.

SYNOPSIS

public void Remove (int a_cell, int a_row);

DESCRIPTION

This function will delete the value from the cell specified by the user in the LH text box.
RETURNS

None. Reports error if cell to delete content from is empty.
*/
        public void Remove(int a_cell, int a_row)
        {
            // Access form control
            DataGridView dgv = Application.OpenForms["frmSpreadsheet"].Controls["dgvSpreadsheet"] as DataGridView;
            dgv[a_cell, a_row].Value = ""; // Remove value from spreadsheet. Error checking for empty cells is done already.
        }
        /*
NAME

Search - Checks if a value from the right hand text box exists in the spreadsheet.

SYNOPSIS

public bool Search (object a_val);

DESCRIPTION

This function will loop through spreadsheet contents to check if the value exists in the spreadsheet.
RETURNS

True if value exists, false if it does not. 0 goes into cell specified by LH text box if false, 1 if true.
*/
        public bool Search(object a_val)
        {
            // Hold values in array list.
            ArrayList dgvVals = new ArrayList();
            DataGridView dgv = Application.OpenForms["frmSpreadsheet"].Controls["dgvSpreadsheet"] as DataGridView;
            for (int i = 0; i < dgv.Columns.Count - 1; i++)
            {
                for (int j = 0; j < dgv.Rows.Count - 1; j++)
                {
                    if (Convert.ToString(dgv[i, j].Value) != String.Empty)
                    {
                        // Add contents of each non empty cell for each column, into arraylist
                        dgvVals.Add(Convert.ToString(dgv[i, j].Value));
                    }

                }
            }
            // If the list contains the value then it exists in the spreadsheet.
            if (dgvVals.Contains(a_val))
            {
                return true;
            }
            return false;
        }
        /*
NAME
CountDoublesLessThan - Counts numeric values less than the input from the right hand text box.

SYNOPSIS

public int CountDoublesLessThan (double a_val);

DESCRIPTION

This function will count the doubles less than a value given in the right hand text box and place the counter in the cell
given in the left hand text box.
RETURNS
The counter. Reports error if a_val cannot be converted to double.
*/
        public int CountDoublesLessThan(double a_val)
        {
            // List to hold values
            List<string> dgvVals = new List<string>();
            DataGridView dgv = Application.OpenForms["frmSpreadsheet"].Controls["dgvSpreadsheet"] as DataGridView;
            for (int i = 0; i < dgv.Columns.Count - 1; i++)
            {
                for (int j = 0; j < dgv.Rows.Count - 1; j++)
                {
                    // For each column place contents of each non empty cell into array list
                    if (Convert.ToString(dgv[i, j].Value) != String.Empty)
                    {
                        dgvVals.Add(Convert.ToString(dgv[i, j].Value));
                    }

                }
            }
            int count = 0;
            double val;
            foreach(string item in dgvVals)
            {
                // For each item in the list, check if it is less than input. If it is, increment counter.
                // If the value cannot be converted to double, move on.
                if (!Double.TryParse(item, out val))
                {
                    continue;
                }
                else
                {
                    val = Convert.ToDouble(item);
                }
                if (val < a_val)
                {
                    count++;
                }
            }
            return count; // Return the count of items less than the input.
        }
        /*
NAME
CountDoublesGreaterThan - Counts numeric values greater than the input from the right hand text box.

SYNOPSIS

public int CountDoublesLessThan (double a_val);

DESCRIPTION

This function will count the doubles greater than a value given in the right hand text box and place the counter in the cell
given in the left hand text box.
RETURNS
The counter. Reports error if a_val cannot be converted to double.
*/
        public int CountDoublesGreaterThan(double a_val)
        {
            // List to hold values
            List<string> dgvVals = new List<string>();
            // Access control
            DataGridView dgv = Application.OpenForms["frmSpreadsheet"].Controls["dgvSpreadsheet"] as DataGridView;
            for (int i = 0; i < dgv.Columns.Count - 1; i++)
            {
                for (int j = 0; j < dgv.Rows.Count - 1; j++)
                {
                    // Insert contents of each non empty cell, for each column, into array list.
                    if (Convert.ToString(dgv[i, j].Value) != String.Empty)
                    {
                        dgvVals.Add(Convert.ToString(dgv[i, j].Value));
                    }

                }
            }
            int count = 0;
            double val;
            foreach (string item in dgvVals)
            {
                // For each item, check if it is greater than the input. If it is increment the counter.
                // If the value cannot be converted to double, move on.
                if (!Double.TryParse(item, out val))
                {
                    continue;
                }
                else
                {
                    val = Convert.ToDouble(item);
                }
                if (val > a_val)
                {
                    count++;
                }
            }
            return count; // Return number of cells with values greater than input.
        }
        /*
NAME
CountAll - Counts non empty cells in the spreadsheet.

SYNOPSIS

public int CountAll ();

DESCRIPTION

This function will count all the non empty cells, place their content in an array list, 
and place the counter in the cell given in the left hand text box.
RETURNS
The length of the arraylist.
*/
        public int CountAll()
        {
            // List of values
            List<string> dgvVals = new List<string>();
            DataGridView dgv = Application.OpenForms["frmSpreadsheet"].Controls["dgvSpreadsheet"] as DataGridView;
            for (int i = 0; i < dgv.Columns.Count - 1; i++)
            {
                for (int j = 0; j < dgv.Rows.Count - 1; j++)
                {
                    // For each column place contents of each non empty cell in array list
                    if (Convert.ToString(dgv[i, j].Value) != String.Empty)
                    {
                        dgvVals.Add(Convert.ToString(dgv[i, j].Value));
                    }

                }
            }
            return dgvVals.Count; // Return length of array list
        }
        /*
NAME
CountOccurences - Counts occurences of input string from the right hand text box.

SYNOPSIS

public int CountOccurences (string a_val);

DESCRIPTION

This function will count the number of occurences of a string given in the right hand text box and places the counter in the cell
given in the left hand text box.
RETURNS

The counter.
*/
        public int CountOccurences(string a_val)
        {
            // List of values
            List<string> dgvVals = new List<string>();
            DataGridView dgv = Application.OpenForms["frmSpreadsheet"].Controls["dgvSpreadsheet"] as DataGridView;
            for (int i = 0; i < dgv.Columns.Count - 1; i++)
            {
                for (int j = 0; j < dgv.Rows.Count - 1; j++)
                {
                    // For each column place contents of each non empty cell in array list
                    if (Convert.ToString(dgv[i, j].Value) != String.Empty)
                    {
                        dgvVals.Add(Convert.ToString(dgv[i, j].Value));
                    }

                }
            }
            int count = 0;
            foreach (string item in dgvVals)
            {
                // For each item in the arraylist check if it is equal to the input string. If it is, increment the count
                if (item.Equals(a_val))
                {
                    count++;
                }
            }
            // Return the number of occurences of data. Don't have to report error for non numeric data in this case.
            return count;
        }
    }
}