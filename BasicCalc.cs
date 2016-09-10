using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace Spreadsheet
{
    class BasicCalc
    {
        private List<double> m_calcs;
        private double m_value;
        Spreadsheet spreadsheet; // Must have this for column names
        public BasicCalc()
        {
            m_calcs = new List<double>();
            m_value = 0;
            spreadsheet = new Spreadsheet(); // Need this to convert strings such as A3 to corresponding cell value.
        }
        /*
        NAME

            Sum - Calculates the sum of a list of values from the right hand text box.

        SYNOPSIS

            public double Sum (Dictionary<string, int> a_cols);

        DESCRIPTION

            This function will calculate the sum of multiple values inside the right hand text box. Excel function: =Sum(a[0], a[1], .... a[n])

        RETURNS

           The sum.
        */
        public double Sum(Dictionary<string, int> a_cols)
        {
            TextBox textbox = Application.OpenForms["frmSpreadsheet"].Controls["txtEnter"] as TextBox;
            string vals = textbox.Text;
            // Get values from text box and access for control
            if (!vals.StartsWith("=Sum(")) // Report error if required function signature is not present
            {
                Errors.RecordError(vals + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            }
            if (!vals.EndsWith(")")) // Report error if no closing parenthesis
            {
                Errors.RecordError(vals + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            }
            string val2 = vals.Remove(0, 5);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(','); // Remove function signature and parenthesis and get array of values
            try
            {
                foreach (string n in elts)
                {
                    // For each parameter passed to function, if it is a column variable or arithmetic expression, get its value,
                    // and add it to list
                    double num = Convert.ToDouble(spreadsheet.Compute(n, a_cols));
                    m_calcs.Add(num);
                }
                // Return the sum of the values
                m_value = m_calcs.Sum();
                return m_value;
            }
            catch(FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
        /*
NAME

    Diff - Calculates the difference between values from the right hand text box.

SYNOPSIS

    public double Sum (Dictionary<string, int> a_cols);

DESCRIPTION

    This function will calculate the difference of values inside the right hand text box. Excel function: =Sub(a[0], a[1], ... a[n])

RETURNS

   The difference.
*/
        public double Diff(Dictionary<string, int> a_cols)
        {
            TextBox textbox = Application.OpenForms["frmSpreadsheet"].Controls["txtEnter"] as TextBox;
            string vals = textbox.Text;
            // Get values from text box, access form control.
            // Report error if function signature does not match or if there are no closing parenthesis.
            if (!vals.StartsWith("=Sub("))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            }
            if (!vals.EndsWith(")"))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            }
            string val2 = vals.Remove(0, 5);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            // Remove function signature and parenthesis and split array by commas
            try
            {
                foreach (string n in elts)
                {
                    // For each parameter passed to function, if it is a column variable or arithmetic expression, get its value,
                    // and add it to list
                    double num = Convert.ToDouble(spreadsheet.Compute(n, a_cols));
                    m_calcs.Add(num);
                }
                m_value = Convert.ToDouble(m_calcs[0]);
                // For each value in the list subtract it from running difference.
                for (int i = 1; i < m_calcs.Count; i++) {
                    m_value -= m_calcs[i];
                }
                return m_value;
            }
            catch(FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }

        }
        /*
NAME

Mult - Calculates the product between values from the right hand text box.

SYNOPSIS

public double Mult (Dictionary<string, int> a_cols);

DESCRIPTION

This function will calculate the product of values inside the right hand text box. Excel function: =Mult(a[0], a[1], ... a[n])

RETURNS

The product
*/
        public double Mult(Dictionary<string, int> a_cols)
        {
            TextBox textbox = Application.OpenForms["frmSpreadsheet"].Controls["txtEnter"] as TextBox;
            string vals = textbox.Text;
            // Get values from text box and access form control.
            // Report error if function signature does not match or if parenthesis dont match
            if (!vals.StartsWith("=Mult("))
            {
                Errors.RecordError("Improper String");
                MessageBox.Show("Improper String");
                return -1;
            }
            if (!vals.EndsWith(")"))
            {
                Errors.RecordError("Improper String");
                MessageBox.Show("Improper String");
                return -1;
            }
            string val2 = vals.Remove(0, 6);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            try
            {
                // Split the list by commas.
                // Remove function signature and parenthesis.
                foreach (string n in elts)
                {
                    // For each parameter passed to function, if it is a column variable or arithmetic expression, get its value,
                    // and add it to list
                    double num = Convert.ToDouble(spreadsheet.Compute(n, a_cols));
                    m_calcs.Add(num);
                }
                m_value = 1;
                // For each value in the list, multiply it to running product. Return final product.
                foreach (double number in m_calcs)
                {
                    m_value *= number;
                }
                return m_value;
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

Div - Calculates the quotient between two values from the right hand text box.

SYNOPSIS

public double Div (Dictionary<string, int> a_cols);

DESCRIPTION

This function will calculate the quotient of two values inside the right hand text box. Excel function: =Div(a, b)

RETURNS

The quotient. -1 if dividing by zero.
*/
        public double Div(Dictionary<string, int> a_cols)
        {
            TextBox textbox = Application.OpenForms["frmSpreadsheet"].Controls["txtEnter"] as TextBox;
            string vals = textbox.Text;
            // Access form control and get values from text box.
            if (!vals.StartsWith("=Div("))
            {
                Errors.RecordError("Improper String");
                MessageBox.Show("Improper String");
                return -1;
            }
            if (!vals.EndsWith(")"))
            {
                Errors.RecordError("Improper String");
                MessageBox.Show("Improper String");
                return -1;
            }
            // Report error if function signature doesn't match. Remove function signature and parenthesis and split string by comma.
            string val2 = vals.Remove(0, 5);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            if (elts.Length != 2)
            {
                Errors.RecordError("Must only have two elements");
                MessageBox.Show("Must only have two elements");
                return -1;
            }
            try
            {
                // Report error if not two elements.
                // If column name or arithmetic expression, get the values.
                double a_val = Convert.ToDouble(spreadsheet.Compute(elts[0], a_cols));
                double b_val = Convert.ToDouble(spreadsheet.Compute(elts[1], a_cols));
                if (b_val == 0) // Cannot divide by zero.
                {
                    Errors.RecordError("Cannot divide by zero");
                    MessageBox.Show("Cannot divide by zero");
                    return -1;
                }
                m_value = (a_val * 1.0) / b_val;
                return m_value; // Return quotient.
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

Mod - Calculates the remainder between dividing a by b from the right hand text box.

SYNOPSIS

public double Mod (Dictionary<string, int> a_cols);

DESCRIPTION

This function will calculate the remainder of dividing a by b inside the right hand text box. Excel function: =Mod(a, b)

RETURNS

The remainder. -1 if first number is negative or second number is less than or equal to zero.
*/
        public double Mod(Dictionary<string, int> a_cols)
        {
            TextBox textbox = Application.OpenForms["frmSpreadsheet"].Controls["txtEnter"] as TextBox;
            string vals = textbox.Text;
            // Get text box values and access control
            // Report error if wrong function signature
            if (!vals.StartsWith("=Mod("))
            {
                Errors.RecordError("Improper String");
                MessageBox.Show("Improper String");
                return -1;
            }
            if (!vals.EndsWith(")"))
            {
                Errors.RecordError("Improper String");
                MessageBox.Show("Improper String");
                return -1;
            }
            string val2 = vals.Remove(0, 5);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            // Report error if not two elements. Split array by commas and remove function signature and parenthesis.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must only have two elements");
                MessageBox.Show("Must only have two elements");
                return -1;
            }
            try
            {
                // If column variable or arithmetic expression, get its value. Return remainder.
                double a_val = Convert.ToDouble(spreadsheet.Compute(elts[0], a_cols));
                double b_val = Convert.ToDouble(spreadsheet.Compute(elts[1], a_cols));
                int mod1, mod2;
                // Report error if using doubles or floats in modulo.
                if (!int.TryParse(Convert.ToString(a_val), out mod1) || !int.TryParse(Convert.ToString(b_val), out mod2))
                {
                    MessageBox.Show("Cannot modulo with doubles or floats!");
                    Errors.RecordError("Cannot modulo with doubles or floats!");
                    return -1;
                } 
                if (b_val <= 0 || a_val < 0) // If any value is negative or the second value is not positive report error that we cannot use negative
                    // numbers in modulo.
                {
                    Errors.RecordError("Cannot modulo by zero");
                    MessageBox.Show("Cannot modulo by zero");
                    return -1;
                }
                m_value = a_val % b_val;
                return m_value;
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