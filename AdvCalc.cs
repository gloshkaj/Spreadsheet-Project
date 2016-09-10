using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace Spreadsheet
{
    class AdvCalc
    {
        private double m_value;
        Spreadsheet spreadsheet;
        public AdvCalc()
        {
            m_value = 0;
            spreadsheet = new Spreadsheet();
        }
/*
NAME

    Sqrt - Calculates the square root of a number, variable, or arithmetic expression from the right hand text box.

SYNOPSIS

    public double Sqrt (string a_vals, Dictionary<string, int> a_cols);

DESCRIPTION

    This function will calculate the square root of the value inside the right hand text box

RETURNS

   The square root of the number. -1 if the input was negative.
*/
        public double Sqrt(string a_vals, Dictionary<string, int> a_cols)
        {
            if (!a_vals.StartsWith("=Sqrt("))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            }
            if (!a_vals.EndsWith(")"))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            } // Report error if wrong function signature. Split string by comma.
            string val2 = a_vals.Remove(0, 6);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            if (elts.Length != 1) // Make sure there is only one element
            {
                Errors.RecordError("Must only have one element");
                MessageBox.Show("Must only have one element");
                return -1;
            }
            try
            {
                if (Convert.ToDouble(spreadsheet.Compute(elts[0], a_cols)) < 0)
                    // Report error if number is negative. If it is an arithmetic expression or column variable get it's value.
                {
                    MessageBox.Show("Cannot take square root of negative number!");
                    Errors.RecordError("Cannot take square root of negative number!");
                    return -1;
                }
                m_value = Math.Sqrt(Convert.ToDouble(spreadsheet.Compute(elts[0], a_cols)));
                return m_value; // Return the square root.
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

            SqrOf - Calculates the square of a number, variable, or arithmetic expression from the right hand text box.

        SYNOPSIS

            public double SqrOf (string a_vals, Dictionary<string, int> a_cols);

        DESCRIPTION

            This function will calculate the square of the value (n ^ 2) inside the right hand text box

        RETURNS

           The square of the number.
        */
        public double SqrOf(string a_vals, Dictionary<string, int> a_cols)
        {
            if (!a_vals.StartsWith("=Sqr("))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            }
            if (!a_vals.EndsWith(")"))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            } // Report error if bad function signature. Split string by comma.
            string val2 = a_vals.Remove(0, 5);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            if (elts.Length != 1) // Report error if there is not only one element.
            {
                Errors.RecordError("Must only have one element");
                MessageBox.Show("Must only have one element");
                return -1;
            }
            try
            {
                m_value = Math.Pow(Convert.ToDouble(spreadsheet.Compute(elts[0], a_cols)), 2);
                return m_value; // Return value ^ 2. If the value is an arithmetic expression or a column variable first get its value.
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

            AbsVal - Calculates the absolute value of a number, variable, or arithmetic expression from the right hand text box.

        SYNOPSIS

            public double AbsVal (string a_vals, Dictionary<string, int> a_cols);

        DESCRIPTION

            This function will calculate the absolute value of the value inside the right hand text box

        RETURNS

           The absolute value.
        */
        public double AbsVal(string a_vals, Dictionary<string, int> a_cols)
        {
            if (!a_vals.StartsWith("=Abs("))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            }
            if (!a_vals.EndsWith(")"))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            } // Report error if bad function signature. Split string by regex.
            string val2 = a_vals.Remove(0, 5);
            string val3 = val2.Remove(val2.Length - 1);
            string pattern = @"[,|?|:|;|!|']";
            string[] elts = Regex.Split(val3, pattern);
            if (elts.Length != 1) // Report error if not one element.
            {
                Errors.RecordError("Must only have one element");
                MessageBox.Show("Must only have one element");
                return -1;
            }
            try
            {
                m_value = Math.Abs(Convert.ToDouble(spreadsheet.Compute(elts[0], a_cols)));
                return m_value; // Return absolute value. If value is an arithmetic expression or column name, get its value.
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

            Exp - Calculates x to the y power of two numbers, variables, or arithmetic expressions from the right hand text box.

        SYNOPSIS

            public double Exp (string a_vals, Dictionary<string, int> a_cols);

        DESCRIPTION

            This function will calculate x to the y power of two values inside the right hand text box

        RETURNS

           The absolute value.
        */
        public double Exp(string a_vals, Dictionary<string, int> a_cols)
        {
            if (!a_vals.StartsWith("=Exp("))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            } // Report error if bad function signature.
            if (!a_vals.EndsWith(")"))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            }
            // Split string by comma and remove function signature and parenthesis.
            string val2 = a_vals.Remove(0, 5);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            if (elts.Length != 2) // Report error if not exactly two values.
            {
                Errors.RecordError("Must only have two elements");
                MessageBox.Show("Must only have two elements");
                return -1;
            }
            try
            {
                // If the values are column variables or arithmetic expressions, get their values.
                double a_val = Convert.ToDouble(spreadsheet.Compute(elts[0], a_cols));
                double b_val = Convert.ToDouble(spreadsheet.Compute(elts[1], a_cols));
                m_value = Math.Pow(a_val, b_val);
                return m_value; // Return x ^ y.
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
    }
    // For this class, BasicCalc, Aggregate, and Spreadsheet, if we reach any of the catch blocks, then we have entered non-numeric data or the parentheses don't match.
    // Report format error in these cases.
}