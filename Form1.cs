using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace Spreadsheet
{
    public partial class frmSpreadsheet : Form
    {
        public frmSpreadsheet()
        {
            InitializeComponent();
            this.dgvSpreadsheet.Rows.Add(104); // Add 104 rows to spreadsheet when form loads
        }
        int count = 0;
        int ccount = 0;
        Dictionary<string, int> m_cols = new Dictionary<string, int>(); // Add column names
        private void frmSpreadsheet_Load(object sender, EventArgs e)
        {
            MessageBox.Show("Welcome to my Senior Project!"); // Display welcome message
            int count = 0;
            // Name each column by letter for arithmetic expression evaluation and add to dictionary.
            foreach (DataGridViewColumn column in dgvSpreadsheet.Columns)
            {
                m_cols.Add(column.HeaderText, count);
                count++;
            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            BasicCalc basicCalc = new BasicCalc();
            string rowCell = txtCell.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = basicCalc.Sum(m_cols);
        }

        private void btnSub_Click(object sender, EventArgs e)
        {
            BasicCalc basicCalc = new BasicCalc();
            string rowCell = txtCell.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = basicCalc.Diff(m_cols);
        }

        private void btnMult_Click(object sender, EventArgs e)
        {
            BasicCalc basicCalc = new BasicCalc();
            string rowCell = txtCell.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0) {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = basicCalc.Mult(m_cols);
        }

        private void btnDiv_Click(object sender, EventArgs e)
        {
            BasicCalc basicCalc = new BasicCalc();
            string rowCell = txtCell.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = basicCalc.Div(m_cols);
        }

        private void btmMod_Click(object sender, EventArgs e)
        {
            BasicCalc basicCalc = new BasicCalc();
            string rowCell = txtCell.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = basicCalc.Mod(m_cols);
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            Container m_Cont = new Container();
            string rowCell = txtCell.Text;
            object value = txtEnter.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            m_Cont.Insert(value, ccount, count);
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            Container m_Cont = new Container();
            string rowCell = txtCell.Text;
            object value = txtEnter.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            m_Cont.Update(value, ccount, count);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            Container m_Cont = new Container();
            string rowCell = txtCell.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            string currentCell = System.Convert.ToString(dgvSpreadsheet[ccount, count].Value);
            // Report error if cell is empty.
            if (currentCell == String.Empty)
            {
                Errors.RecordError("No value to delete exists at this position");
                MessageBox.Show("Cannot delete from an empty cell");
                return;
            }
            else
            {
                m_Cont.Remove(ccount, count);
            }

        }

        private void btnSort_Click(object sender, EventArgs e)
        {
            Container m_Cont = new Container();
            m_Cont.Sort(); // Sort the list in ascending order
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            Container m_Cont = new Container();
            string value = txtEnter.Text;
            string rowCell = txtCell.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            if (m_Cont.Search(value)) // If the value exists the cell should have 1, otherwise 0.
            {
                MessageBox.Show("This value exists in the spreadsheet.");
                dgvSpreadsheet[ccount, count].Value = 1;
            }
            else
            {
                MessageBox.Show("This value does not exist in the spreadsheet.");
                dgvSpreadsheet[ccount, count].Value = 0;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Errors.DisplayErrors();
            MessageBox.Show("Thank you for using my spreadsheet application!");
            Application.Exit(); // Display any errors, a thank you message, and exit the program.
        }

        private void btnAvg_Click(object sender, EventArgs e)
        {
            Aggregate dataset = new Aggregate();
            string rowCell = txtCell.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = dataset.Avg(m_cols);
        }

        private void btnMax_Click(object sender, EventArgs e)
        {
            Aggregate dataset = new Aggregate();
            string rowCell = txtCell.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = dataset.Max(m_cols);
        }

        private void btnMin_Click(object sender, EventArgs e)
        {
            Aggregate dataset = new Aggregate();
            string rowCell = txtCell.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = dataset.Min(m_cols);
        }

        private void btnMed_Click(object sender, EventArgs e)
        {
            Aggregate dataset = new Aggregate();
            string rowCell = txtCell.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = dataset.Med(m_cols);
        }

        private void btnStdev_Click(object sender, EventArgs e)
        {
            Aggregate dataset = new Aggregate();
            string rowCell = txtCell.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = System.Convert.ToDouble(dataset.Sdv(m_cols));
        }

        private void btnExp_Click(object sender, EventArgs e)
        {
            AdvCalc calc = new AdvCalc();
            string rowCell = txtCell.Text;
            string content = txtEnter.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box and data from right hand text box. Report error if there are not exactly two elements in left hand text box.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = calc.Exp(content, m_cols);
        }

        private void btnsqr_Click(object sender, EventArgs e)
        {
            AdvCalc calc = new AdvCalc();
            string rowCell = txtCell.Text;
            string content = txtEnter.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box and data in right hand text box. 
            // Report error if there are not exactly two elements in left hand text box.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = calc.Sqrt(content, m_cols);
        }

        private void btnabs_Click(object sender, EventArgs e)
        {
            AdvCalc calc = new AdvCalc();
            string rowCell = txtCell.Text;
            string content = txtEnter.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box and data from right hand text box.
            // Report error if there are not exactly two elements in left hand text box.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = calc.AbsVal(content, m_cols);
        }

        private void btnpowtwo_Click(object sender, EventArgs e)
        {
            AdvCalc calc = new AdvCalc();
            string rowCell = txtCell.Text;
            string content = txtEnter.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box and data from right hand text box. Report error if there are not exactly two elements in left hand text box.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = calc.SqrOf(content, m_cols);
        }

        private void btnCount_Click(object sender, EventArgs e)
        {
            Container m_Cont = new Container();
            string rowCell = txtCell.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = m_Cont.CountAll();
        }

        private void btnCountWithPred_Click(object sender, EventArgs e)
        {
            Container m_Cont = new Container();
            string rowCell = txtCell.Text;
            double value;
            // Report error if querying non doubles.
            if (!Double.TryParse(txtEnter.Text, out value)) {
                MessageBox.Show("Cannot query non double");
                Errors.RecordError("Cannot query non double!");
                return;
            }
            double val = System.Convert.ToDouble(txtEnter.Text);
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box and data from right hand text box. Report error if there are not exactly two elements in left hand text box.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = m_Cont.CountDoublesLessThan(val);
        }

        private void btnCountOccurences_Click(object sender, EventArgs e)
        {
            Container m_Cont = new Container();
            string rowCell = txtCell.Text;
            string val = txtEnter.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box and data from right hand text box. Report error if there are not exactly two elements in left hand text box.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = m_Cont.CountOccurences(val);
        }

        private void btnCountGreater_Click(object sender, EventArgs e)
        {
            Container m_Cont = new Container();
            string rowCell = txtCell.Text;
            double value;
            // Report error if querying non doubles.
            if (!Double.TryParse(txtEnter.Text, out value))
            {
                MessageBox.Show("Cannot query non double");
                Errors.RecordError("Cannot query non double!");
                return;
            }
            double val = System.Convert.ToDouble(txtEnter.Text);
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box and data from right hand text box. Report error if there are not exactly two elements in left hand text box.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = m_Cont.CountDoublesGreaterThan(val);
        }

        private void btnSortDescending_Click(object sender, EventArgs e)
        {
            Container m_Cont = new Container();
            m_Cont.SortDescending(); // Sort in descending order
        }

        private void btnArithExp_Click(object sender, EventArgs e)
        {
            Spreadsheet spreadsheet = new Spreadsheet();
            string content = txtEnter.Text;
            string rowCell = txtCell.Text;
            string newContent = "";
            // Get cell numbers from left hand text box and data from right hand text box. Report error if there are not exactly two elements in left hand text box.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (!content.StartsWith("=")) // Report error if the expression does not start with an equals sign
            {
                MessageBox.Show("Must start expression with equals sign!");
                Errors.RecordError("Must have equals sign at beginning of expression!");
                return;
            }
            newContent = content.Remove(0, 1); // Remove the equals sign before evaluating
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = spreadsheet.Compute(newContent, m_cols); // Place the result of the arithmetic expression in the specified cell.
        }

        private void btnVariance_Click(object sender, EventArgs e)
        {
            Aggregate dataset = new Aggregate();
            string rowCell = txtCell.Text;
            string splitter = " ";
            string[] elts = rowCell.Split(Char.Parse(splitter));
            // Get cell numbers from left hand text box. Report error if there are not exactly two elements.
            // Also report error if at least one is negative or not numeric. Place the result of the specified calculation in the cell.
            if (elts.Length != 2)
            {
                Errors.RecordError("Must have exactly two elements, row and column #.");
                MessageBox.Show("Must have exactly two elements, row and column #.");
                return;
            }
            int num1, num2;
            if (!int.TryParse(elts[0], out num1) || !int.TryParse(elts[1], out num2))
            {
                MessageBox.Show("Cannot have non-numeric column or row");
                Errors.RecordError("Cannot have non-numeric column or row");
                return;
            }
            ccount = System.Convert.ToInt32(elts[0]);
            count = System.Convert.ToInt32(elts[1]);
            if (ccount < 0 || count < 0)
            {
                MessageBox.Show("Invalid column or row");
                Errors.RecordError("Invalid column or row");
                return;
            }
            dgvSpreadsheet[ccount, count].Value = System.Convert.ToDouble(dataset.Var(m_cols));
        }

        private void dgvSpreadsheet_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            Spreadsheet spreadsheet = new Spreadsheet(); // When editing is done the result should be displayed inside the cell that they just edited.
            if (System.Convert.ToString(dgvSpreadsheet[e.ColumnIndex, e.RowIndex].Value) == String.Empty)
            {
                return;
            } // Stop if the cell was empty
            string content = dgvSpreadsheet[e.ColumnIndex, e.RowIndex].Value.ToString(); // Get the value from the cell.
            // The result would be displayed when the user leaves the cell.
            string newContent = "";
            string value = "";
            if (!content.StartsWith("=")) // Stop if the string does not start with the equals sign.
            {
                return;
            }
            newContent = content.Remove(0, 1).Trim(); // Remove all leading and trailing whitespace.
            if (newContent.StartsWith("Sum(")) // Addition
            {
                value = spreadsheet.CalcSumFromCell(newContent, m_cols).ToString();
            }
            else if (newContent.StartsWith("Sub(")) // Subtraction. Though there is no Excel function for subtract, it it very similar to the sum function.
            {
                value = spreadsheet.CalcDiffFromCell(newContent, m_cols).ToString();
            }
            else if (newContent.StartsWith("Mult(")) // Multiplication
            {
                value = spreadsheet.CalcProductFromCell(newContent, m_cols).ToString();
            }
            else if (newContent.StartsWith("Div(")) // Division
            {
                value = spreadsheet.CalcQuotientFromCell(newContent, m_cols).ToString();
            }
            else if (newContent.StartsWith("Mod(")) // Modulo
            {
                value = spreadsheet.CalcModuloFromCell(newContent, m_cols).ToString();
            }
            else if (newContent.StartsWith("Exp(")) // Exponentiation
            {
                value = spreadsheet.CalcExpFromCell(newContent, m_cols).ToString();
            }
            else if (newContent.StartsWith("Sqr(")) // Square (x ^ 2)
            {
                value = spreadsheet.CalcPowTwoFromCell(newContent, m_cols).ToString();
            }
            else if (newContent.StartsWith("Sqrt(")) // Square root
            {
                value = spreadsheet.CalcSqrtFromCell(newContent, m_cols).ToString();
            }
            else if (newContent.StartsWith("Abs(")) // Absolute value
            {
                value = spreadsheet.CalcAbsValFromCell(newContent, m_cols).ToString();
            }
            else if (newContent.StartsWith("Max(")) // Maximum
            {
                value = spreadsheet.CalcMaxFromCell(newContent, m_cols).ToString();
            }
            else if (newContent.StartsWith("Min(")) // Minimum
            {
                value = spreadsheet.CalcMinFromCell(newContent, m_cols).ToString();
            }
            else if (newContent.StartsWith("Med(")) // Median
            {
                value = spreadsheet.CalcMedFromCell(newContent, m_cols).ToString();
            }
            else if (newContent.StartsWith("Avg(")) // Average
            {
                value = spreadsheet.CalcAvgFromCell(newContent, m_cols).ToString();
            }
            else if (newContent.StartsWith("Sdv(")) // Standard Deviation
            {
                value = spreadsheet.CalcSdvFromCell(newContent, m_cols).ToString();
            }
            else if (newContent.StartsWith("Var(")) // Variance
            {
                value = spreadsheet.CalcStatisticalVarianceFromCell(newContent, m_cols).ToString();
            }
            else if (newContent == String.Empty) // Report error if nothing was typed after equals sign. Empty the cell.
            {
                MessageBox.Show("Don't have any expression!");
                Errors.RecordError("Must input a formula!");
                spreadsheet.DisplayInCell("");
                return;
            }
            else if (newContent.StartsWith("(") || newContent.StartsWith("-") || newContent.StartsWith(".") || char.IsLetterOrDigit(newContent[0]))
            {
                // Arithmetic expression, if the string starts with a parenthesis, a unary minus, a decimal point, a digit, or a letter.
                value = spreadsheet.Compute(newContent, m_cols);
            }
            else // If we get here report error that the user is using a function that is not supported by the program.
            {
                MessageBox.Show("Unsupported function");
                Errors.RecordError("Unsupported function");
                return;
            }
            // Display the result in the cell.
            spreadsheet.DisplayInCell(value);
        }
    }
}