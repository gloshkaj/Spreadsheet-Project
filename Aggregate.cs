using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace Spreadsheet
{
    class Aggregate
    {
        private double m_value;
        private List<double> m_calcs;
        Spreadsheet spreadsheet;
        public Aggregate()
        {
            m_value = 0;
            m_calcs = new List<double>();
            spreadsheet = new Spreadsheet();
        }
        /*
NAME

    Max - Calculates the maximum of a list of values from the right hand text box.

SYNOPSIS

    public double Max (Dictionary<string, int> a_cols);

DESCRIPTION

    This function will calculate the maximum of multiple values inside the right hand text box. Excel function: =Max(a[0], a[1], .... a[n])

RETURNS

   The maximum.
*/
        public double Max(Dictionary<string, int> a_cols)
        {
            TextBox textbox = Application.OpenForms["frmSpreadsheet"].Controls["txtEnter"] as TextBox;
            string vals = textbox.Text;
            // Access form control, get values, and check function signature.
            if (!vals.StartsWith("=Max("))
            {
                Errors.RecordError(vals + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            }
            if (!vals.EndsWith(")"))
            {
                Errors.RecordError(vals + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            } // Report error if bad signature or missed parenthesis. Split by comma.
            string val2 = vals.Remove(0, 5);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            try
            {
                foreach (string n in elts)
                {
                    // For each value in the list, convert to double.
                    // Get values of any column variables or arithmetic expressions.
                    double num = Convert.ToDouble(spreadsheet.Compute(n, a_cols));
                    m_calcs.Add(num);
                }
                m_value = m_calcs.Max();
                return m_value; // Return maximum
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
        /*
NAME

    Min - Calculates the minimum of a list of values from the right hand text box.

SYNOPSIS

    public double Min (Dictionary<string, int> a_cols);

DESCRIPTION

    This function will calculate the minimum of multiple values inside the right hand text box. Excel function: =Min(a[0], a[1], .... a[n])

RETURNS

   The minimum.
*/
        public double Min(Dictionary<string, int> a_cols)
        {
            TextBox textbox = Application.OpenForms["frmSpreadsheet"].Controls["txtEnter"] as TextBox;
            string vals = textbox.Text;
            // Access form control and get values from text box.
            if (!vals.StartsWith("=Min("))
            {
                Errors.RecordError(vals + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            }
            if (!vals.EndsWith(")"))
            {
                Errors.RecordError(vals + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            } // Report error if bad function signature. Split by comma and remove function signature and parenthesis.
            string val2 = vals.Remove(0, 5);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            try
            {
                foreach (string n in elts)
                {
                    // For each value in the list, add to the running min. Get value of all arithmetic expressions or column names.
                    double num = Convert.ToDouble(spreadsheet.Compute(n, a_cols));
                    m_calcs.Add(num);
                }
                m_value = m_calcs.Min();
                return m_value; // Return the minimum.
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
        /*
NAME

    Avg - Calculates the average of a list of values from the right hand text box.

SYNOPSIS

    public double Avg (Dictionary<string, int> a_cols);

DESCRIPTION

    This function will calculate the average of multiple values inside the right hand text box. Excel function: =Avg(a[0], a[1], .... a[n])

RETURNS

   The average.
*/
        public double Avg(Dictionary<string, int> a_cols)
        {
            TextBox textbox = Application.OpenForms["frmSpreadsheet"].Controls["txtEnter"] as TextBox;
            string vals = textbox.Text;
            // Access form control and get text box values.
            if (!vals.StartsWith("=Avg("))
            {
                Errors.RecordError(vals + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            }
            if (!vals.EndsWith(")"))
            {
                Errors.RecordError(vals + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            }
            // Report error if bad function signature. Remove closing parenthesis and function signature. Split by comma.
            string val2 = vals.Remove(0, 5);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            try
            {
                foreach (string n in elts)
                {
                    // Put each parameter in list. Get values of all arithmetic expressions or column names.
                    double num = Convert.ToDouble(spreadsheet.Compute(n, a_cols));
                    m_calcs.Add(num);
                }
                m_value = m_calcs.Average();
                return m_value; // Return the mean of the list of values.
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
        /*
NAME

Med - Calculates the median of a list of values from the right hand text box.

SYNOPSIS

public double Med (Dictionary<string, int> a_cols);

DESCRIPTION

This function will calculate the median of multiple values inside the right hand text box. Excel function: =Med(a[0], a[1], .... a[n])

RETURNS

The median if an odd number of values were given. The average of the middle two numbers otherwise.
*/
        public double Med(Dictionary<string, int> a_cols)
        {
            TextBox textbox = Application.OpenForms["frmSpreadsheet"].Controls["txtEnter"] as TextBox;
            string vals = textbox.Text;
            // Access form controls and text box string. Report error if bad function signature.
            if (!vals.StartsWith("=Med("))
            {
                Errors.RecordError(vals + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            }
            if (!vals.EndsWith(")"))
            {
                Errors.RecordError(vals + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            }
            // Remove function signature and closing parenthesis. Split by comma.
            string val2 = vals.Remove(0, 5);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            try
            {
                foreach (string n in elts)
                {
                    // For each parameter in function, add to list. Get values of all arithmetic expressions or column names.
                    double num = Convert.ToDouble(spreadsheet.Compute(n, a_cols));
                    m_calcs.Add(num);
                }

                m_calcs.Sort(); // Must sort the list to get the median.
                int count = m_calcs.Count;
                if (count % 2 != 0) // If count is odd then return the median as the middle number.
                {
                    m_value = m_calcs[count / 2];
                    return m_value;
                }
                else // Otherwise return the median as the average of the two middle numbers.
                {
                    double m_medsum = m_calcs[count / 2] + m_calcs[(count / 2 - 1)];
                    return (m_medsum * 1.0) / 2;  
                }
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
        /*
NAME

Sdv - Calculates the standard deviation of a list of values from the right hand text box.

SYNOPSIS

public double Sdv (Dictionary<string, int> a_cols);

DESCRIPTION

This function will calculate the standard deviation of multiple values inside the right hand text box. Excel function: =Sdv(a[0], a[1], .... a[n])

RETURNS

The standard deviation.
*/
        public double Sdv(Dictionary<string, int> a_cols)
        {
            TextBox textbox = Application.OpenForms["frmSpreadsheet"].Controls["txtEnter"] as TextBox;
            string vals = textbox.Text;
            // Access form control and get text box values. Report error if bad function signature.
            if (!vals.StartsWith("=Sdv("))
            {
                Errors.RecordError(vals + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            }
            if (!vals.EndsWith(")"))
            {
                Errors.RecordError(vals + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            }
            // Remove function signature and closing parenthesis. Split by comma.
            string val2 = vals.Remove(0, 5);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            try
            {
                foreach (string n in elts)
                {
                    // For each parameter in function, add to list. Get values of all arithmetic expressions and column names.
                    double num = Convert.ToDouble(spreadsheet.Compute(n, a_cols));
                    m_calcs.Add(num);
                }
                // Code adapted from http://stackoverflow.com/questions/5336457/how-to-calculate-a-standard-deviation-array
                double average = m_calcs.Average(); // Compute the average to determine standard deviation
                double sumOfDerivation = 0;
                foreach (double m_value in m_calcs)
                {
                    // Add the square of each value's difference from average in the list to the running sum of derivation.
                    sumOfDerivation += Math.Pow(Math.Abs(m_value - average), 2);
                }
                double sumOfDerivationAverage = sumOfDerivation / (m_calcs.Count - 1); // Variance = sumofDerivation divided by count - 1
                // The standard deviation is the squareroot of the variance.
                // Return the standard deviation. Note: this will be the sample standard deviation
                // and not the population standard deviation.
                return Math.Sqrt(sumOfDerivationAverage);
                // End of adapted code
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
        /*
  NAME

  Var - Calculates the variance of a list of values from the right hand text box.

  SYNOPSIS

  public double Var (Dictionary<string, int> a_cols);

  DESCRIPTION

  This function will calculate the variance of multiple values inside the right hand text box. Excel function: =Sdv(a[0], a[1], .... a[n])

  RETURNS

  The variance.
  */
        public double Var(Dictionary<string, int> a_cols)
        {
            TextBox textbox = Application.OpenForms["frmSpreadsheet"].Controls["txtEnter"] as TextBox;
            string vals = textbox.Text;
            // Access form control and get text box values. Report error if bad function signature.
            if (!vals.StartsWith("=Var("))
            {
                Errors.RecordError(vals + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            }
            if (!vals.EndsWith(")"))
            {
                Errors.RecordError(vals + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            }
            // Remove function signature and closing parenthesis. Split by comma.
            string val2 = vals.Remove(0, 5);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            try
            {
                foreach (string n in elts)
                {
                    // For each parameter in function, add to list. Get values of all arithmetic expressions and column names.
                    double num = Convert.ToDouble(spreadsheet.Compute(n, a_cols));
                    m_calcs.Add(num);
                }
                // Code adapted from http://stackoverflow.com/questions/5336457/how-to-calculate-a-standard-deviation-array
                double average = m_calcs.Average(); // Compute the average to determine variance
                double sumOfDerivation = 0;
                foreach (double m_value in m_calcs)
                {
                    // Add the square of each value's difference from average in the list to the running sum of derivation.
                    sumOfDerivation += Math.Pow(Math.Abs(m_value - average), 2);
                }
                return sumOfDerivation / (m_calcs.Count - 1); // Variance = sumofDerivation divided by count - 1
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
    }
}