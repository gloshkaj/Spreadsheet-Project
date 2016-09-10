using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace Spreadsheet
{
    class Spreadsheet
    {
        private object m_value;
        private List<double> m_calcs;
        public Spreadsheet()
        {
            m_value = 0;
            m_calcs = new List<double>();
        }
        // For the purpose of the Excel functions, if we reached a catch block, report error that there is non numeric data and stop.
        // All values taken in from Excel functions in this class are assumed to be from spreadsheet cells.
        /*
        NAME

            Compute - Converts column variables to value in corresponding cells.

        SYNOPSIS

            public double Compute (string a_exp, Dictionary<string, int> a_cols);

        DESCRIPTION

            This function will compute the arithmetic expressions by replacing column variables with the values inside the cells.

        RETURNS

           The result of the converted arithmetic expression.
        */
        public string Compute(string a_exp, Dictionary<string, int> a_cols)
        {
            // Access form control
            DataGridView dgv = Application.OpenForms["frmSpreadsheet"].Controls["dgvSpreadsheet"] as DataGridView;
            // Code adapted from  http://stackoverflow.com/questions/3210393/how-do-i-remove-all-non-alphanumeric-characters-from-a-string-except-dash
            string newStr = a_exp.Trim(); // Remove spaces
            // Find all digits in the array. If there are none, stop.
            char[] arr = newStr.ToCharArray();
            char[] arr2 = Array.FindAll<char>(arr, (c => (char.IsDigit(c))));
            if (arr2.Length == 0)
            {
                return a_exp;
            }
            a_exp = newStr.ToUpper();
            // Convert to uppercase and find all letters, digits, division signs, and whitespace and put results in new string.
            arr = Array.FindAll<char>(arr, (c => (char.IsLetterOrDigit(c) || c.Equals('/') || char.IsWhiteSpace(c))));
            string str = new string(arr);
            str = Regex.Replace(str, "\\b" + '/' + "\\b", " / "); // If there are any division signs without anything else put whitespace on each side of it. 
            // End of Code Taken from Internet
            string[] elts = str.Split(Char.Parse(" ")); // Split by whitespace
            Dictionary<string, string> termList = new Dictionary<string, string>(); // Dictionary to hold old term and new term
            int cellMatches = 0; // Number of column variables in string.
            foreach (string item in elts)
            {
                string newItem = item.ToUpper();
                string term = newItem; // Convert to uppercase. If empty or starts with a digit move on.
                if (term == String.Empty) continue;
                if (char.IsDigit(term[0])) continue;
                if (term.Equals("/")) continue;
                // Report error if the term begins with something other than A, B, or C and not followed by a digit.
                if (!term[0].Equals('A') && !term[0].Equals('B') && !term[0].Equals('C') && !char.IsDigit(term[1]))
                {
                    MessageBox.Show("Cell out of bounds!");
                    Errors.RecordError("Cell out of bounds!");
                    return a_exp;
                }
                foreach (KeyValuePair<string, int> kvp in a_cols)
                {
                    string nonums = Regex.Replace(term, "[0-9]", "");
                    if (nonums.Equals(kvp.Key)) // If the term begins with a certain key in the dictionary,
                    {
                        // Replace it with the value followed by a comma and split by the comma.
                        term = term.Replace(kvp.Key, kvp.Value + ",");
                        string[] vals = term.Split(',');
                        int val1 = Convert.ToInt32(vals[0]); // Column
                        int val2 = Convert.ToInt32(vals[1]); // Row
                        // Report error if out of range of the control.
                        if (val1 < 0 || val2 < 1 || val1 > dgv.Columns.Count - 1 || val2 > dgv.Rows.Count)
                        {
                            MessageBox.Show("Bad cell supplied!");
                            Errors.RecordError("Bad Cell Choice!");
                            return a_exp;
                        }
                        // Report error if computing from empty cell
                        if (Convert.ToString(dgv[val1, val2 - 1].Value) == String.Empty)
                        {
                            MessageBox.Show("Cannot compute from empty cell!");
                            Errors.RecordError("Cannot compute from empty cell!");
                            return a_exp;
                        }
                        double number;
                        // Report error if computing from cell with non-numeric data
                        if (!Double.TryParse(Convert.ToString(dgv[val1, val2 - 1].Value), out number)) {
                            MessageBox.Show("Cannot compute from cell with non-numeric data");
                            Errors.RecordError("Cannot compute from cell with non-numeric data");
                            return a_exp;
                        }
                        // Replace old term with number in the cell.
                        term = term.Replace(term, dgv[val1, val2 - 1].Value.ToString());
                        // Add to dictionary and increment number of column variables
                        termList.Add(newItem, term);
                        cellMatches++;
                        break;
                    }
                }
            }
            if (cellMatches > 0) // If we found a column variable,
            {
                // Replace each item in the variables array with the number in its corresponding cell. Using regex. Only replaces EXACT MATCHES
                foreach (KeyValuePair<string,string> kvp in termList)
                {
                    string pattern = "\\b" + kvp.Key + "\\b";
                    a_exp = Regex.Replace(a_exp, pattern, kvp.Value);
                }    
            }
            // Return the result of the computed arithmetic expression
            return ComputeBasicCalcFromCell(a_exp);
        }
/// <summary>
/// string ComputeBasicCalcFromCell(string a_content): Computes the converted arithmetic expression.
/// </summary>
/// <param name="a_content"></param>
/// <returns>The result of the converted arithmetic expression</returns>
        public string ComputeBasicCalcFromCell(string a_content)
        {
            // Get the result of the converted expression.
            return Expression.Evaluate(a_content);
        }
/// <summary>
/// void DisplayInCell(string a_value): Displays the result of a calculation in the cell.
/// </summary>
/// <param name="a_value"></param>
        public void DisplayInCell(string a_value)
        {
            DataGridView dgv = Application.OpenForms["frmSpreadsheet"].Controls["dgvSpreadsheet"] as DataGridView;
            m_value = a_value; // Access spreadsheet control and put the result of the computation in the cell.
            dgv.CurrentCell.Value = Convert.ToString(m_value);
        }
/// <summary>
/// CalcSumFromCell: Gets the sum of values from the cell.
/// </summary>
/// <param name="a_content"></param>
/// <param name="a_cols"></param>
/// <returns>Running sum</returns>
        public double CalcSumFromCell(string a_content, Dictionary<string, int> a_cols)
        {
            if (!a_content.EndsWith(")"))
            {
                Errors.RecordError(a_content + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            } // Report error if no closing parenthesis specified.
            string val2 = a_content.Remove(0, 4);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(','); // Remove function signature and parenthesis and split by comma.
            try
            {
                foreach (string n in elts)
                { // For each parameter add to list. If there are column variables or arithmetic expressions get their values.
                    string str = n;
                    string content = Compute(str, a_cols);
                    double num = Convert.ToDouble(content);
                    m_calcs.Add(num);
                }
                m_value = m_calcs.Sum();
                // Return the sum
                return Convert.ToDouble(m_value);
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
        /// <summary>
        /// CalcDiffFromCell: Calculate running difference
        /// </summary>
        /// <param name="a_content"></param>
        /// <param name="a_cols"></param>
        /// <returns>Running difference</returns>
        public double CalcDiffFromCell(string a_content, Dictionary<string, int> a_cols)
        {
            if (!a_content.EndsWith(")"))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            } // Same as before, report error if missing closing parenthesis.
            string val2 = a_content.Remove(0, 4);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            try
            {
                foreach (string n in elts)
                {
                    // For each parameter passed to function, if it is a column variable or arithmetic expression, get its value,
                    // and add it to list
                    double num = Convert.ToDouble(Compute(n, a_cols));
                    m_calcs.Add(num);
                }
                double diff = m_calcs[0];
                // Subtract each value of the list from running difference
                for (int i = 1; i < m_calcs.Count; i++)
                {
                    diff -= m_calcs[i];
                }
                m_value = diff; // Return difference
                return Convert.ToDouble(m_value);
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
/// <summary>
/// CalcProductFromCell: Calculates running product
/// </summary>
/// <param name="a_content"></param>
/// <param name="a_cols"></param>
/// <returns>Running Product</returns>
        public double CalcProductFromCell(string a_content, Dictionary<string, int> a_cols)
        {
            if (!a_content.EndsWith(")"))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            } // Report error if no closing parenthesis. Split by comma and remove function signature
            string val2 = a_content.Remove(0, 5);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            try
            {
                foreach (string n in elts)
                {
                    // For each parameter passed to function, if it is a column variable or arithmetic expression, get its value,
                    // and add it to list
                    double num = Convert.ToDouble(Compute(n, a_cols));
                    m_calcs.Add(num);
                }
                double product = 1;
                // For each number in the list, multiply it to running product and return product.
                foreach (double number in m_calcs)
                {
                    product *= number;
                }
                m_value = product;
                return Convert.ToDouble(m_value);
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
/// <summary>
/// CalcQuotientFromCell: Calculates quotient of two values in cell.
/// </summary>
/// <param name="a_content"></param>
/// <param name="a_cols"></param>
/// <returns>Quotient. -1 if dividing by zero.</returns>
        public double CalcQuotientFromCell(string a_content, Dictionary<string, int> a_cols)
        {
            if (!a_content.EndsWith(")"))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            } // Report error if no closing parenthesis.
            string val2 = a_content.Remove(0, 4);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            if (elts.Length != 2) // Split by comma and remove function signature. Report error if not exactly two elements.
            {
                Errors.RecordError("Must only have two elements");
                MessageBox.Show("Must only have two elements");
                return -1;
            }
            try
            {
                // If values are column names or arithmetic expressions convert them
                double a_val = Convert.ToDouble(Compute(elts[0], a_cols));
                double b_val = Convert.ToDouble(Compute(elts[1], a_cols));
                if (b_val == 0)
                { // Report error if dividing by zero.
                    Errors.RecordError("Cannot divide by zero");
                    MessageBox.Show("Cannot divide by zero");
                    return -1;
                }
                m_value = (a_val * 1.0) / b_val; // Return quotient
                return Convert.ToDouble(m_value);
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
/// <summary>
/// CalcModuloFromCell: Calculates remainder of division.
/// </summary>
/// <param name="a_content"></param>
/// <param name="a_cols"></param>
/// <returns>the remainder</returns>
        public double CalcModuloFromCell(string a_content, Dictionary<string, int> a_cols)
        {
            if (!a_content.EndsWith(")"))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            } // Report error if bad function signature. Split by comma
            string val2 = a_content.Remove(0, 4);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            if (elts.Length != 2) // Must only have two elements
            {
                Errors.RecordError("Must only have two elements");
                MessageBox.Show("Must only have two elements");
                return -1;
            }
            try
            { // Convert any column names or arithmetic expressions
                double a_val = Convert.ToDouble(Compute(elts[0], a_cols));
                double b_val = Convert.ToDouble(Compute(elts[1], a_cols));
                // Report error if using doubles or floats in modulo.
                int mod1, mod2;
                if (!int.TryParse(Convert.ToString(a_val), out mod1) || !int.TryParse(Convert.ToString(b_val), out mod2))
                {
                    MessageBox.Show("Cannot modulo with doubles or floats!");
                    Errors.RecordError("Cannot modulo with doubles or floats!");
                    return -1;
                }
                if (a_val < 0 || b_val <= 0) // If a is negative or b is non positive report error and return -1.
                {
                    Errors.RecordError("Cannot modulo by non positive number");
                    MessageBox.Show("Cannot modulo by non positive number");
                    return -1;
                }
                m_value = a_val % b_val;
                return Convert.ToDouble(m_value);
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
/// <summary>
/// CalcExpFromCell: calculates x to the y power
/// </summary>
/// <param name="a_content"></param>
/// <param name="a_cols"></param>
/// <returns>x ^ y</returns>
        public double CalcExpFromCell(string a_content, Dictionary<string, int> a_cols)
        {
            if (!a_content.EndsWith(")"))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            }
            string val2 = a_content.Remove(0, 4);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(','); // Check function signature and split by comma.
            if (elts.Length != 2) // verify that we have exactly two parameters
            {
                Errors.RecordError("Must only have two elements");
                MessageBox.Show("Must only have two elements");
                return -1;
            }
            try
            { // Convert any column names or arithmetic expressions. Return x ^ y.
                double a_val = Convert.ToDouble(Compute(elts[0], a_cols));
                double b_val = Convert.ToDouble(Compute(elts[1], a_cols));
                m_value = Math.Pow(a_val, b_val);
                return Convert.ToDouble(m_value);
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
/// <summary>
/// CalcAbsValFromCell: Calculates absolute value of a number
/// </summary>
/// <param name="a_content"></param>
/// <param name="a_cols"></param>
/// <returns>Absolute Value</returns>
        public double CalcAbsValFromCell(string a_content, Dictionary<string, int> a_cols)
        {
            if (!a_content.EndsWith(")"))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            } // Split by comma and check function signature. Make sure there is only one parameter
            string val2 = a_content.Remove(0, 4);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            if (elts.Length != 1)
            {
                Errors.RecordError("Must only have one element");
                MessageBox.Show("Must only have one element");
                return -1;
            }
            try
            {
                m_value = Math.Abs(Convert.ToDouble(Compute(elts[0], a_cols))); // Return absolute value of converted string.
                return Convert.ToDouble(m_value);
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
/// <summary>
/// CalcPowTwoFromCell: Calculates the square of a number
/// </summary>
/// <param name="a_content"></param>
/// <param name="a_cols"></param>
/// <returns>x ^ 2</returns>
        public double CalcPowTwoFromCell(string a_content, Dictionary<string, int> a_cols)
        {
            if (!a_content.EndsWith(")"))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            } // Check function signature and split by comma. Check that there is only one parameter.
            string val2 = a_content.Remove(0, 4);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            if (elts.Length != 1)
            {
                Errors.RecordError("Must only have one element");
                MessageBox.Show("Must only have one element");
                return -1;
            }
            try
            {
                m_value = Math.Pow(Convert.ToDouble(Compute(elts[0], a_cols)), 2); // Return square of converted element.
                return Convert.ToDouble(m_value);
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
/// <summary>
/// CalcSqrtFromCell: Calculates square root of number.
/// </summary>
/// <param name="a_content"></param>
/// <param name="a_cols"></param>
/// <returns>Square root. -1 if parameter was negative.</returns>
        public double CalcSqrtFromCell(string a_content, Dictionary<string, int> a_cols)
        {
            if (!a_content.EndsWith(")"))
            {
                Errors.RecordError("Improper Calculation");
                MessageBox.Show("Improper String");
                return -1;
            } // Check function signature and split by comma. Check that there is only one element.
            string val2 = a_content.Remove(0, 5);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            if (elts.Length != 1)
            {
                Errors.RecordError("Must only have one element");
                MessageBox.Show("Must only have one element");
                return -1;
            }
            try
            {
                // Return square root of converted string. Report error if number is negative. In that case return -1.
                if (Convert.ToDouble(Compute(elts[0], a_cols)) < 0)
                {
                    MessageBox.Show("Cannot take square root of negative number!");
                    Errors.RecordError("Cannot take square root of negative number!");
                    return -1;
                }
                m_value = Math.Sqrt(Convert.ToDouble(Compute(elts[0], a_cols)));
                return Convert.ToDouble(m_value);
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
/// <summary>
/// CalcMaxFromCell: Computes maximum of dataset
/// </summary>
/// <param name="a_content"></param>
/// <param name="a_cols"></param>
/// <returns>The maximum</returns>
        public double CalcMaxFromCell(string a_content, Dictionary<string, int> a_cols)
        {
            if (!a_content.EndsWith(")"))
            {
                Errors.RecordError(a_content + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            } // Check function signature and split parameters by comma.
            string val2 = a_content.Remove(0, 4);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            try
            {
                foreach (string n in elts)
                {
                    // Add each parameter to array list
                    double num = Convert.ToDouble(Compute(n, a_cols));
                    m_calcs.Add(num);
                }
                m_value = m_calcs.Max();
                // Return maximum. must convert to double
                return Convert.ToDouble(m_value);
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
/// <summary>
/// CalcMinFromCell: Calculates minimum from dataset
/// </summary>
/// <param name="a_content"></param>
/// <param name="a_cols"></param>
/// <returns>The minimum</returns>
        public double CalcMinFromCell(string a_content, Dictionary<string, int> a_cols)
        {
            if (!a_content.EndsWith(")"))
            {
                Errors.RecordError(a_content + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            } // Check function signature and split by comma.
            string val2 = a_content.Remove(0, 4);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            try
            {
                foreach (string n in elts)
                {
                    // Place each value from dataset into list
                    double num = Convert.ToDouble(Compute(n, a_cols));
                    m_calcs.Add(num);
                }
                m_value = m_calcs.Min(); // Return minimum
                return Convert.ToDouble(m_value);
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
/// <summary>
/// CalcMedFromCell: Calculates middle number from dataset
/// </summary>
/// <param name="a_content"></param>
/// <param name="a_cols"></param>
/// <returns>Median if there are an odd number of numbers, average of two middle numbers otherwise</returns>
        public double CalcMedFromCell(string a_content, Dictionary<string, int> a_cols)
        {
            if (!a_content.EndsWith(")"))
            {
                Errors.RecordError(a_content + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            } // Check function signature and split by comma.
            string val2 = a_content.Remove(0, 4);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            try
            {
                foreach (string n in elts)
                {
                    // Put each value into list.
                    double num = Convert.ToDouble(Compute(n, a_cols));
                    m_calcs.Add(num);
                }

                m_calcs.Sort(); // Must sort the list for median calculation
                int count = m_calcs.Count;
                if (count % 2 != 0)
                {
                    // If the count is odd then the median is the middle number
                    m_value = m_calcs[count / 2];
                    return Convert.ToDouble(m_value);
                }
                else
                {
                    // Otherwise it is the average of the two middle numbers.
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
/// <summary>
/// Calculates average of dataset
/// </summary>
/// <param name="a_content"></param>
/// <param name="a_cols"></param>
/// <returns>Average of dataset</returns>
        public double CalcAvgFromCell(string a_content, Dictionary<string, int> a_cols)
        {
            if (!a_content.EndsWith(")"))
            {
                Errors.RecordError(a_content + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            } // Check function signature and split by comma.
            string val2 = a_content.Remove(0, 4);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            try
            {
                foreach (string n in elts)
                {
                    // Place each value of dataset into list.
                    double num = Convert.ToDouble(Compute(n, a_cols));
                    m_calcs.Add(num);
                }
                m_value = m_calcs.Average(); // Return the average (sum / count)
                return Convert.ToDouble(m_value);
            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
/// <summary>
/// Calculates sample standard deviation from dataset
/// </summary>
/// <param name="a_content"></param>
/// <param name="a_cols"></param>
/// <returns>Standard deviation. Note: we are calculating the sample standard deviation.</returns>
        public double CalcSdvFromCell(string a_content, Dictionary<string, int> a_cols)
        {
            if (!a_content.EndsWith(")"))
            {
                Errors.RecordError(a_content + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            } // Check function signature and split by comma.
            string val2 = a_content.Remove(0, 4);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            try
            {
                foreach (string n in elts)
                {
                    // Place each value into list.
                    double num = Convert.ToDouble(Compute(n, a_cols));
                    m_calcs.Add(num);
                }
                // Code adapted from http://stackoverflow.com/questions/5336457/how-to-calculate-a-standard-deviation-array
                double average = m_calcs.Average(); // Get average
                double sumOfDerivation = 0;
                foreach (double value in m_calcs)
                {
                    // For each value in list add square of difference between value and average to running sum of derivation
                    sumOfDerivation += Math.Pow(Math.Abs(value - average), 2);
                }
                // variance is the sumofDerivation divided by count - 1
                double sumOfDerivationAverage = sumOfDerivation / (m_calcs.Count - 1);
                return Math.Sqrt(sumOfDerivationAverage); // Return the square root of that result
                // End of adapted code

            }
            catch (FormatException fe)
            {
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
/// <summary>
/// Calculates statistical variance of dataset
/// </summary>
/// <param name="a_content"></param>
/// <param name="a_cols"></param>
/// <returns>The statistical variance.</returns>
        public double CalcStatisticalVarianceFromCell(string a_content, Dictionary<string, int> a_cols)
        {
            if (!a_content.EndsWith(")"))
            {
                Errors.RecordError(a_content + " is an invalid operation");
                MessageBox.Show("Improper Calculation");
                return -1;
            } // Check function signature and split by comma.
            string val2 = a_content.Remove(0, 4);
            string val3 = val2.Remove(val2.Length - 1);
            string[] elts = val3.Split(',');
            try
            {
                foreach (string n in elts)
                {
                    // Place each value in list.
                    double num = Convert.ToDouble(Compute(n, a_cols));
                    m_calcs.Add(num);
                }
                // Code adapted from http://stackoverflow.com/questions/5336457/how-to-calculate-a-standard-deviation-array
                double average = m_calcs.Average(); // Must have average to calculate standard deviation.
                double sumOfDerivation = 0;
                foreach (double value in m_calcs)
                {
                    // For each value in list, add to running sum of derivation the square of difference between value and average
                    sumOfDerivation += Math.Pow(Math.Abs(value - average), 2);
                }
                // Return variance (sum of derivation divided by count - 1)
                return sumOfDerivation / (m_calcs.Count - 1);
                // End of adapted code

            }
            catch (FormatException fe)
            {
                // this is for error reporting. It comes up many times in the entire project.
                Errors.RecordError(fe.Message);
                MessageBox.Show(fe.Message);
                return -1;
            }
        }
    }
}