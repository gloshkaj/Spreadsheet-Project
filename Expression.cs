using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Spreadsheet
{
    public static class Expression
    {
        // Code adapted from http://www.geeksforgeeks.org/expression-evaluation/
        // *Expressions include spaces in between tokens and use infix notation*
/*
NAME

    Evaluate - evaluates the arithmetic expression.

SYNOPSIS

    string Evaluate( string a_expression );

DESCRIPTION

    This function will evaluate the arithmetic expression in "a_expression" given in the cells or the textbox
    and evaluate the answer

RETURNS

   The original expression if the expression is invalid, the result of the expression otherwise.
*/
        public static string Evaluate(string a_expression)
        {
            // Validate expression. Report error if it does not have supported operators
            if (!a_expression.Contains("+") && !a_expression.Contains("-") && !a_expression.Contains("*")
                && !a_expression.Contains("/") && !a_expression.Contains("%") && !a_expression.Contains("^") &&
                !a_expression.Contains(".") && !a_expression.Contains("(") && !a_expression.Contains(")"))
            {
                foreach (char c in a_expression)
                {
                    if (!char.IsDigit(c)) // Report error if there are unsupported characters
                    {
                        MessageBox.Show("Not a mathematical expression!");
                        Errors.RecordError("This is not a mathematical expression!");
                        return a_expression;
                    }
                }
                return a_expression;
            }
            try
            {
                char[] tokens = a_expression.ToCharArray();

                // Stack for numbers: 'values'
                Stack<double> values = new Stack<double>();

                // Stack for Operators: 'ops'
                Stack<char> ops = new Stack<char>();

                for (int i = 0; i < tokens.Length; i++)
                {
                    // Current token is a whitespace, skip it
                    if (tokens[i] == ' ')
                        continue;

                    // Current token is a number, push it to stack for numbers
                    if (tokens[i] >= '.' && tokens[i] <= '9' && !tokens[i].Equals('/'))
                    {
                        string sbuf = "";
                        // There may be more than one digits in number. There also may be decimal numbers and fractions
                        while (i < tokens.Length && tokens[i] >= '.' && tokens[i] <= '9')
                            sbuf += (tokens[i++]);
                        if (sbuf.Contains("/")) // If there is a fraction first solve the fraction!
                        {
                            string[] split = sbuf.Split('/');
                            double result = Convert.ToDouble(split[0]);
                            for (int j = 1; j < split.Length; j++)
                            {
                                result /= Convert.ToDouble(split[j]);
                            }
                            values.Push(result);
                        }
                        else
                        {
                            values.Push(Convert.ToDouble(sbuf));
                        }
                        if (i != tokens.Length && !tokens[i].Equals(Char.Parse(" "))) // Report error if term is not followed by whitespace
                        {
                            MessageBox.Show("Must have spaces in between numbers!");
                            Errors.RecordError("Term must be separated by a space");
                            return a_expression;
                        }
                        if (i != tokens.Length && !tokens[i+1].Equals('+') && !tokens[i+1].Equals('-') && // Report error if a number is not followed by an operator
                             !tokens[i+1].Equals('*') && !tokens[i+1].Equals('/') && !tokens[i+1].Equals('%') &&
                             !tokens[i+1].Equals('^') && !tokens[i + 1].Equals(')'))
                        {
                            MessageBox.Show("Must follow number by operator!");
                            Errors.RecordError("number must be followed by an operator");
                            return a_expression;
                        }
                    }
                    else if (tokens[i] == '-' && tokens[i+1] >= '.' && tokens[i+1] <= '9' && !tokens[i+1].Equals('/')) // Unary minus. Do the same as above with value * -1.
                    {
                        string sbuf = "-";
                        // Must set sbuf equal to "-" to deal with unary minus.
                        i++;
                        // There may be more than one digits in number. There also may be decimal numbers and fractions
                        while (i < tokens.Length && tokens[i] >= '.' && tokens[i] <= '9')
                            sbuf += (tokens[i++]);
                        if (sbuf.Contains("/")) // If there is a fraction first solve the fraction!
                        {
                            string[] split = sbuf.Split('/');
                            double result = Convert.ToDouble(split[0]);
                            for (int j = 1; j < split.Length; j++)
                            {
                                result /= Convert.ToDouble(split[j]);
                            }
                            // Push result onto numbers stack
                            values.Push(result);
                        }
                        else
                        {
                            values.Push(Convert.ToDouble(sbuf));
                        }
                        if (i != tokens.Length && !tokens[i].Equals(Char.Parse(" "))) // Report error if term is not followed by whitespace
                        {
                            MessageBox.Show("Must have spaces in between numbers!");
                            Errors.RecordError("Term must be separated by a space");
                            return a_expression;
                        }
                        if (i != tokens.Length && !tokens[i + 1].Equals('+') && !tokens[i + 1].Equals('-') && // Report error if a number is not followed by an operator
                             !tokens[i + 1].Equals('*') && !tokens[i + 1].Equals('/') && !tokens[i + 1].Equals('%') &&
                             !tokens[i + 1].Equals('^') && !tokens[i + 1].Equals(')'))
                        {
                            MessageBox.Show("Must follow number by operator!");
                            Errors.RecordError("number must be followed by an operator");
                            return a_expression;
                        }
                    }
                    // Current token is an opening brace, push it to 'ops'
                    else if (tokens[i] == '(')
                    {
                        ops.Push(tokens[i]);
                        if (i != tokens.Length - 1 && !tokens[i + 1].Equals(Char.Parse(" "))) // Report error if term is not followed by whitespace
                        {
                            MessageBox.Show("Must have spaces in between parentheses!");
                            Errors.RecordError("Term must be separated by a space");
                            return a_expression;
                        }
                    }


                    // Closing brace encountered, solve entire brace
                    else if (tokens[i] == ')')
                    {
                        if (i != tokens.Length - 1 && !tokens[i + 1].Equals(Char.Parse(" "))) // Report error if term is not followed by whitespace
                        {
                            MessageBox.Show("Must have spaces in between parentheses!");
                            Errors.RecordError("Term must be separated by a space");
                            return a_expression;
                        }
                        if (i != tokens.Length - 1 && !tokens[i + 2].Equals('+') && !tokens[i + 2].Equals('-') && // Report error if a closing parenthesis is not followed by an operator
                             !tokens[i + 2].Equals('*') && !tokens[i + 2].Equals('/') && !tokens[i + 2].Equals('%') &&
                             !tokens[i + 2].Equals('^') && !tokens[i + 2].Equals(')'))
                        {
                            MessageBox.Show("Must follow closing parenthesis by operator!");
                            Errors.RecordError("Closing parenthesis must be followed by an operator");
                            return a_expression;
                        }
                        while (ops.Peek() != '(') // Push values inside closing parenthesis
                            values.Push(ApplyOp(ops.Pop(), values.Pop(), values.Pop()));
                        ops.Pop();
                    }

                    // Current token is an operator.
                    else if (tokens[i] == '+' || tokens[i] == '-' ||
                             tokens[i] == '*' || tokens[i] == '/' || tokens[i] == '%' || tokens[i] == '^')
                    {
                        // While top of 'ops' has same or greater precedence to current
                        // token, which is an operator. Apply operator on top of 'ops'
                        // to top two elements in values stack
                        if (i != tokens.Length - 1 && !tokens[i + 1].Equals(Char.Parse(" "))) // Report error if term is not followed by whitespace
                        {
                            MessageBox.Show("Must have spaces in between operators!");
                            Errors.RecordError("Term must be separated by a space");
                            return a_expression;
                        }
                        while (ops.Count != 0 && HasPrecedence(tokens[i], ops.Peek()))
                            values.Push(ApplyOp(ops.Pop(), values.Pop(), values.Pop()));

                        // Push current token to 'ops'.
                        ops.Push(tokens[i]);
                    }
                    else // If it has anything other than these characters then report error that the expression is invalid!
                    {
                        MessageBox.Show("Expression has invalid characters");
                        Errors.RecordError("This is not a mathematical expression!");
                        return a_expression;
                    }
                }

                // Entire expression has been parsed at this point, apply remaining
                // ops to remaining values
                while (ops.Count != 0)
                    values.Push(ApplyOp(ops.Pop(), values.Pop(), values.Pop()));

                // Top of 'values' contains result, return it
                return values.Pop().ToString();
            }
            catch (Exception) // If we get here then the user messed up in the expression. Report syntax error.
            {
                MessageBox.Show("Syntax Error in Expression");
                Errors.RecordError("Syntax Error in expression");
                return a_expression;
            }
        }
        /*
       NAME

           HasPrecedence - Determines precedence of operators.

       SYNOPSIS

           bool HasPrecedence(char a_op1, char a_op2);

       DESCRIPTION

           This function will evaluate the precendence of the two operators, a_op1 and a_op2

       RETURNS
       
            True if 'a_op2' has higher or same precedence as 'a_op1', false otherwise.
       */
        public static bool HasPrecedence(char a_op1, char a_op2)
        {
            if (a_op2 == '(' || a_op2 == ')') // Parentheses return false
                return false;
            if ((a_op1 == '*' || a_op1 == '/' || a_op1 == '%') && (a_op2 == '+' || a_op2 == '-')) // If op2 is of lower precedence than op1 return false
                return false;
            if ((a_op1 == '^') && (a_op2 == '*' || a_op2 == '/' || a_op2 == '%' || a_op2 == '+' || a_op2 == '-')) // If op1 is exponentiation and op2 is anything else return false
                return false;
            else // If we get here than op2 is of higher precedence than op1
                return true;
        }
        /*
NAME

    ApplyOp - applys the given operator on two operands in the arithmetic expression.

SYNOPSIS

    double ApplyOp(char a_op, double a_num2, double a_num1);

DESCRIPTION

    This function will apply an operator 'a_op' on operands 'a_num1' 
    and 'a_num2'.

RETURNS

   The result. -1 if invalid.
*/
        public static double ApplyOp(char a_op, double a_num2, double a_num1)
        {
            switch (a_op)
            {
                case '+':
                    return a_num1 + a_num2; // addition
                case '-':
                    return a_num1 - a_num2; // subtraction
                case '*':
                    return a_num1 * a_num2; // multiplication
                case '/': // division
                    if (a_num2 == 0) // Report error if dividing by zero
                    {
                        MessageBox.Show("Cannot Divide By Zero");
                        Errors.RecordError("Cannot divide by zero");
                        return -1; // Division by zero will return -1
                    }
                    return a_num1 / a_num2;
                case '%': // modulo
                    if (a_num1 < 0 || a_num2 <= 0) // Report error if we are using modulo with a negative number
                    {
                        MessageBox.Show("Cannot use negative numbers for modulo");
                        Errors.RecordError("Cannot use negative numbers for modulo");
                        return -1; // modulo with a non positive number will return -1
                    }
                    int mod1, mod2;
                    // Report error if using doubles or floats in modulo.
                    if (!int.TryParse(Convert.ToString(a_num1), out mod1) || !int.TryParse(Convert.ToString(a_num2), out mod2))
                    {
                        MessageBox.Show("Cannot modulo with doubles or floats!");
                        Errors.RecordError("Cannot modulo with doubles or floats!");
                        return -1;
                    }
                    return a_num1 % a_num2;
                case '^': // exponentiation
                    return Math.Pow(a_num1, a_num2);
                default: // If we get here then we are using an unsupported operator
                    MessageBox.Show("Unsupported operator");
                    Errors.RecordError("Unsupported operator");
                    break;
            }
            return 0; // return zero if an unsupported operator is used
        }
    }
}
// End of adapted code