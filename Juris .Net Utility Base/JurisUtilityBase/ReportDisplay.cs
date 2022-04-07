using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;


namespace JurisUtilityBase
{
    public partial class ReportDisplay : Form
    {
        public ReportDisplay(DataTable AR, DataTable allOthers, JurisUtility ju)
        {
            InitializeComponent();
            dataGridView1.DataSource = AR;
            dataGridView2.DataSource = allOthers;
            _jurisUtility = ju;
            //enable certain columns to be editable (fee, exp, etc)
            dataGridView1.ReadOnly = false;
            foreach (DataGridViewColumn dc in dataGridView1.Columns)
            {
                if (dc.Index > 11 && dc.Index < 20)
                    dc.ReadOnly = false;
                else
                    dc.ReadOnly = true;
            }
            //hide known balances because they are just for verification that their allocation amount does not exceed them
            dataGridView1.Columns["billFees"].Visible = false;
            dataGridView1.Columns["billCashExps"].Visible = false;
            dataGridView1.Columns["billNCashExps"].Visible = false;
            dataGridView1.Columns["billInt"].Visible = false;
            dataGridView1.Columns["billTax1"].Visible = false;
            dataGridView1.Columns["billTax2"].Visible = false;
            dataGridView1.Columns["billTax3"].Visible = false;
            dataGridView1.Columns["billSur"].Visible = false;
            dataGridView1.Columns["matsys"].Visible = false;
            string sql = "IF OBJECT_ID('tempdb.dbo.##ErrorsTool', 'U') IS NOT NULL DROP TABLE [dbo].[##ErrorsTool]";
            _jurisUtility.ExecuteNonQuery(0, sql);
            sql = "create table ##ErrorsTool (row int, err varchar(500))";
            _jurisUtility.ExecuteNonQuery(0, sql);


        }

        private JurisUtility _jurisUtility;

        List<DataError> de = new List<DataError>();
        DataError error;
        int row = 0;

        private void buttonBack_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView2.DataSource = null;
            this.Close();
        }

        private void buttonPrint_Click(object sender, EventArgs e)
        {
            //click save
            //verify the total allocs per line = ar alloc
            //if not, stop and force them to while displaying check number
            //if so, update table and move on to actual updates
            foreach (DataGridViewRow dc in dataGridView1.Rows)
            {
                row++;
                string allocAmt = dc.Cells[11].Value.ToString();
                string checkNum = dc.Cells[0].Value.ToString();
                //ensure all their entered allocs are numbers
                if (!IsNumeric(dc.Cells[12].Value.ToString()) || !IsNumeric(dc.Cells[13].Value.ToString()) || !IsNumeric(dc.Cells[14].Value.ToString()) || !IsNumeric(dc.Cells[15].Value.ToString())
                     || !IsNumeric(dc.Cells[16].Value.ToString()) || !IsNumeric(dc.Cells[17].Value.ToString()) || !IsNumeric(dc.Cells[18].Value.ToString()) || !IsNumeric(dc.Cells[19].Value.ToString()))
                {
                    error = new DataError();
                    error.rowNum = row;
                    error.error = "All allocations for check number " + checkNum + " must be a valid a valid number";
                    de.Add(error);
                    break;
                }
                //ensure the allocs they entered are not greater than the amounts left to be allocated
                else if (Convert.ToDouble(dc.Cells[12].Value.ToString()) > Convert.ToDouble(dc.Cells[20].Value.ToString()))
                {
                    error = new DataError();
                    error.rowNum = row;
                    error.error = "The fee allocation for check number " + checkNum + " is greater than the amount left to be allocated. The remaining balance is " + Math.Round(Convert.ToDouble(dc.Cells[20].Value.ToString()), 2).ToString();
                    de.Add(error);
                    continue;
                }
                else if (Convert.ToDouble(dc.Cells[13].Value.ToString()) > Convert.ToDouble(dc.Cells[21].Value.ToString()))
                {
                    error = new DataError();
                    error.rowNum = row;
                    error.error = "The cash expense allocation for check number " + checkNum + " is greater than the amount left to be allocated. The remaining balance is " + Math.Round(Convert.ToDouble(dc.Cells[21].Value.ToString()), 2).ToString();
                    de.Add(error);
                    continue;
                }
                else if (Convert.ToDouble(dc.Cells[14].Value.ToString()) > Convert.ToDouble(dc.Cells[22].Value.ToString()))
                {
                    error = new DataError();
                    error.rowNum = row;
                    error.error = "The non cash expense allocation for check number " + checkNum + " is greater than the amount left to be allocated. The remaining balance is " + Math.Round(Convert.ToDouble(dc.Cells[22].Value.ToString()), 2).ToString();
                    de.Add(error);
                    continue;
                }
                else if (Convert.ToDouble(dc.Cells[15].Value.ToString()) > Convert.ToDouble(dc.Cells[24].Value.ToString()))
                {
                    error = new DataError();
                    error.rowNum = row;
                    error.error = "The tax1 allocation for check number " + checkNum + " is greater than the amount left to be allocated. The remaining balance is " + Math.Round(Convert.ToDouble(dc.Cells[24].Value.ToString()), 2).ToString();
                    de.Add(error);
                    continue;
                }
                else if (Convert.ToDouble(dc.Cells[16].Value.ToString()) > Convert.ToDouble(dc.Cells[25].Value.ToString()))
                {
                    error = new DataError();
                    error.rowNum = row;
                    error.error = "The tax2 allocation for check number " + checkNum + " is greater than the amount left to be allocated. The remaining fee balance is " + Math.Round(Convert.ToDouble(dc.Cells[25].Value.ToString()), 2).ToString();
                    de.Add(error);
                    continue;
                }
                else if (Convert.ToDouble(dc.Cells[17].Value.ToString()) > Convert.ToDouble(dc.Cells[26].Value.ToString()))
                {
                    error = new DataError();
                    error.rowNum = row;
                    error.error = "The tax3 allocation for check number " + checkNum + " is greater than the amount left to be allocated. The remaining balance is " + Math.Round(Convert.ToDouble(dc.Cells[26].Value.ToString()), 2).ToString();
                    de.Add(error);
                    continue;
                }
                else if (Convert.ToDouble(dc.Cells[18].Value.ToString()) > Convert.ToDouble(dc.Cells[23].Value.ToString()))
                {
                    error = new DataError();
                    error.rowNum = row;
                    error.error = "The interest allocation for check number " + checkNum + " is greater than the amount left to be allocated. The remaining balance is " + Math.Round(Convert.ToDouble(dc.Cells[23].Value.ToString()), 2).ToString();
                    de.Add(error);
                    continue;
                }
                else if (Convert.ToDouble(dc.Cells[19].Value.ToString()) > Convert.ToDouble(dc.Cells[27].Value.ToString()))
                {
                    error = new DataError();
                    error.rowNum = row;
                    error.error = "The surcharge allocation for check number " + checkNum + " is greater than the amount left to be allocated. The remaining balance is " + Math.Round(Convert.ToDouble(dc.Cells[27].Value.ToString()), 2).ToString();
                    de.Add(error);
                    continue;
                }
                else //all checks out...now see if total allocs = ar alloc
                {
                    double allocs = (Convert.ToDouble(dc.Cells[19].Value.ToString()) + Convert.ToDouble(dc.Cells[18].Value.ToString()) +
                        Convert.ToDouble(dc.Cells[17].Value.ToString()) + Convert.ToDouble(dc.Cells[16].Value.ToString()) + Convert.ToDouble(dc.Cells[15].Value.ToString()) +
                        Convert.ToDouble(dc.Cells[14].Value.ToString()) + Convert.ToDouble(dc.Cells[13].Value.ToString()) + Convert.ToDouble(dc.Cells[12].Value.ToString()));
                    allocs = Math.Round(allocs,2);
                    if (Convert.ToDouble(allocAmt) != allocs)
                    {

                        error = new DataError();
                        error.rowNum = row;
                        error.error = "The total allocations of for check number " + checkNum + " do not add up to the AR Allocation amount";
                        de.Add(error);
                        continue;

                    }

                }

            }
            DataTable dt = new DataTable();
            if (de.Count == 0)
            {
                //Adding the Columns.
                foreach (DataGridViewColumn column in dataGridView1.Columns)
                {
                    dt.Columns.Add(column.HeaderText, column.ValueType);
                }

                //Adding the Rows.
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    dt.Rows.Add();
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();
                    }
                }
                //get data out of datatable and into sql

                string sql = "delete from  ##CombineChecks where ttype = 1";//we no longer need all the multiple records...we just need 1 per check per, per invoice, per matter
                _jurisUtility.ExecuteNonQuery(0, sql);
                foreach (DataRow dr in dt.Rows)
                {
                    sql = "insert into ##CombineChecks (Deposit_Date, Check_Number, Check_Amount, Check_Date, [Payor], " +
                               " Client_Code, Matter_Code, Bank_Code, GL_Account, [Reference], " +
                               " Invoice_Date, Invoice_Number, Bill_Total, Bill_Balance, AR_Allocation, Fees, Cash_Exps, NonCash_Exps, Taxes1, " +
                               " Taxes2, Taxes3, Surcharge, Interest, PPD_Allocation, Trust_Allocation, Other_Allocation, TTYpe, batchNo,  " +
                                " billFees, billCashExps, billNCashExps, billInt, billTax1, billTax2, billTax3, billSur, matsys) " +
                                " values (convert(varchar,'" + dr[8].ToString() + "', 101), '" + dr[0].ToString() + "', " + dr[7].ToString() + ", convert(varchar, '" + dr[5].ToString() + "', 101), '" + dr[6].ToString() + "', " +
                                " '" + dr[1].ToString() + "', '" + dr[2].ToString() + "', '', '', '', " +
                                " convert(varchar, '" + dr[3].ToString() + "', 101), " + dr[4].ToString() + ", " + dr[10].ToString() + ", " + dr[9].ToString() + ", " +
                                dr[11].ToString() + ", " + dr[12].ToString() + ", " + dr[13].ToString() + ", " + dr[14].ToString() + ", " + dr[15].ToString() +
                                " ," + dr[16].ToString() + "," + dr[17].ToString() + "," + dr[18].ToString() + "," + dr[19].ToString() + ",0.00, 0.00, 0.00, 1, 0, " +
                                dr[20].ToString() + ", " + dr[21].ToString() + ", " + dr[22].ToString() + ", " + dr[23].ToString() + ", " + dr[24].ToString() + ", " +
                                dr[25].ToString() + ", " + dr[26].ToString() + ", " + dr[27].ToString() + ", " + dr[28].ToString() + 
                                ")";
                    _jurisUtility.ExecuteNonQuery(0, sql); 
                }

                this.Close();


            }
            else
            {
                MessageBox.Show("There were issues with some of the AR allocations. They will now be displayed.");
                string sql = "delete from ##ErrorsTool";
                _jurisUtility.ExecuteNonQuery(0, sql);

                foreach (DataError err in de)
                {
                    sql = "insert into ##ErrorsTool (row, err) values(" + err.rowNum.ToString() + ", '" + err.error + "')";
                    _jurisUtility.ExecuteNonQuery(0, sql);
                }
                sql = "select row as Row_Number, err as Error_Message from ##ErrorsTool";
                DataSet DSErr = new DataSet();
                DSErr = _jurisUtility.RecordsetFromSQL(sql);
                if (DSErr != null && DSErr.Tables.Count > 0 && DSErr.Tables[0].Rows.Count > 0)
                {

                    DataTable tabErr = DSErr.Tables[0];
                    ErrorDisplay ed = new ErrorDisplay(tabErr);
                    ed.ShowDialog();
                }
                row = 0;
                de.Clear();
            }


        }


        private static bool IsNumeric(object Expression)
        {
            double retNum = 0.00;

            try
            {
                bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
                return true;
            }
            catch (Exception dd) {
                MessageBox.Show(Expression.ToString());
                
                
                
                return false; }

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellLeave(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {

        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.ThrowException = false;
        }
    }
}
