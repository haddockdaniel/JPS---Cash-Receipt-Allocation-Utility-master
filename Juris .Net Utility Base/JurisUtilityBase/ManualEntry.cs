using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace JurisUtilityBase
{
    public partial class ManualEntry : Form
    {
        public ManualEntry(JurisUtility jutil, System.Drawing.Point ppt, int empsys, string empid)
        {
            InitializeComponent();
            _jurisUtility = jutil;
            pt = ppt;
            this.Location = pt;
            empsysnbr = empsys;
            employeeID = empid;
            richTextBox1.Text = "By choosing this option, you can select a range of bills that you wish to apply this receipt to. This will display all matters associated with the selected bills as long as they have an outstanding balance.";
            string sql = "SELECT case when max([CRBBatchNbr]) is null then 1 else max([CRBBatchNbr]) + 1 end as BN FROM [CashReceiptsBatch]";
            DataSet ds = jutil.RecordsetFromSQL(sql);
            try
            {

                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        batchNo = dr[0].ToString();
                    }
                }
                textBoxBatch.Text = "AUTO BATCH: " + batchNo + " Created by " + employeeID + " on " + DateTime.Now.ToShortDateString();
                textBoxCheckDate.Text = DateTime.Now.ToShortDateString();
                textBoxDepDate.Text = DateTime.Now.ToShortDateString();
            }
            catch (Exception ex1) { MessageBox.Show("There was a problem getting your batch numbers. The program will now exit." + "\r\n" + "Details: " + ex1.Message, "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error); System.Environment.Exit(1); }
            fillInComboBoxes();
        }

        JurisUtility _jurisUtility;
        private System.Drawing.Point pt;
        public int clisysnbr = 0;
        public int matsysnbr = 0;
        public int clisysnbrTo = 0;
        string batchNo = "";
        public int empsysnbr = 0;
        public string employeeID = "";

        private void labelAmountApplied_Click(object sender, EventArgs e)
        {

        }

        private void fillInComboBoxes()
        {
            DataSet myRSPC2;
            comboBoxBank.Items.Clear();
            string SQLPC2 = "select BnkCode + '    ' + left(BnkDesc, 30) as Bank from BankAccount order by BnkCode";
            myRSPC2 = _jurisUtility.RecordsetFromSQL(SQLPC2);

            if (myRSPC2 == null || myRSPC2.Tables.Count == 0 || myRSPC2.Tables[0].Rows.Count == 0)
                MessageBox.Show("There are no Banks. Correct and run the tool again", "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                foreach (DataRow dr in myRSPC2.Tables[0].Rows)
                    comboBoxBank.Items.Add(dr["Bank"].ToString());
                comboBoxBank.SelectedIndex = 0;
            }

            comboBoxClient.Items.Clear();
            myRSPC2.Clear();
            SQLPC2 = "select dbo.jfn_FormatClientCode(clicode) + '    ' + left(clireportingname, 30) as Client from client order by dbo.jfn_FormatClientCode(clicode)";
            myRSPC2 = _jurisUtility.RecordsetFromSQL(SQLPC2);

            if (myRSPC2 == null || myRSPC2.Tables.Count == 0 || myRSPC2.Tables[0].Rows.Count == 0)
                MessageBox.Show("There are no Clients. Correct and run the tool again", "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                foreach (DataRow dr in myRSPC2.Tables[0].Rows)
                    comboBoxClient.Items.Add(dr["Client"].ToString());
                comboBoxClient.SelectedIndex = 0;
            }

            comboBoxMatter.Items.Clear();
            myRSPC2.Clear();
            SQLPC2 = "select dbo.jfn_FormatMatterCode(MatCode) + '    ' + left(matreportingname, 30) as Matter from matter order by dbo.jfn_FormatMatterCode(MatCode)";
            myRSPC2 = _jurisUtility.RecordsetFromSQL(SQLPC2);

            if (myRSPC2 == null || myRSPC2.Tables.Count == 0 || myRSPC2.Tables[0].Rows.Count == 0)
                MessageBox.Show("There are no Matters. Correct and run the tool again", "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                foreach (DataRow dr in myRSPC2.Tables[0].Rows)
                    comboBoxMatter.Items.Add(dr["Matter"].ToString());
                comboBoxMatter.SelectedIndex = 0;
            }

            comboBoxGL.Items.Clear();
            myRSPC2.Clear();
            SQLPC2 = "select dbo.jfn_FormatChartOfAccount(ChartOfAccounts.ChtSysNbr) + '    ' + left(ChtDesc, 30) as GL from ChartOfAccounts order by dbo.jfn_FormatChartOfAccount(ChartOfAccounts.ChtSysNbr)";
            myRSPC2 = _jurisUtility.RecordsetFromSQL(SQLPC2);

            if (myRSPC2 == null || myRSPC2.Tables.Count == 0 || myRSPC2.Tables[0].Rows.Count == 0)
                MessageBox.Show("There are no GL Accounts. Correct and run the tool again", "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                foreach (DataRow dr in myRSPC2.Tables[0].Rows)
                    comboBoxGL.Items.Add(dr["GL"].ToString());
                comboBoxGL.SelectedIndex = 0;
            }
        }

        private void checkBoxPctOfBill_CheckedChanged(object sender, EventArgs e)
        {
            textBoxAllocPct.Enabled = checkBoxPctOfBill.Checked;
        }

        private void checkBoxUseMatter_CheckedChanged(object sender, EventArgs e)
        {
            textBoxMatter.Enabled = checkBoxUseMatter.Checked;
            buttonMatter.Enabled = checkBoxUseMatter.Checked;
        }

        private void buttonCliLookUp_Click(object sender, EventArgs e)
        {
            ClientLookUp cl = new ClientLookUp(_jurisUtility, pt);
            cl.ShowDialog();
            if (cl.clientSelected)
            {
                clisysnbr = cl.clisysnbr;
                textBoxClient.Text = cl.clicode;
                labelClientName.Text = cl.clientName;
                labelClientName.Visible = true;
            }
            cl.Close();
        }

        private void buttonMatter_Click(object sender, EventArgs e)
        {
            if (clisysnbr != 0)
            {
                MatLookUp cl = new MatLookUp(_jurisUtility, pt, clisysnbr);
                cl.ShowDialog();
                if (cl.matterSelected)
                {
                    matsysnbr = cl.matsysnbr;
                    textBoxMatter.Text = cl.matcode;
                }
                else
                {
                    checkBoxUseMatter.Checked = false;
                    textBoxMatter.Enabled = false;
                    buttonMatter.Enabled = false;

                }
                cl.Close();
            }
            else
            {
                MessageBox.Show("A client must be selected first", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ManualEntry_Load(object sender, EventArgs e)
        {

        }

        private void buttonClientTo_Click(object sender, EventArgs e)
        {
            ClientLookUp cl = new ClientLookUp(_jurisUtility, pt);
            cl.ShowDialog();
            if (cl.clientSelected)
            {
                clisysnbrTo = cl.clisysnbr;
                textBoxClientTo.Text = cl.clicode;
            }
            cl.Close();
        }

        private void radioButtonBillRange_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonBillRange.Checked)
            {
                textBoxBillFrom.Enabled = true;
                textBoxBillTo.Enabled = true;
                textBoxClientTo.Enabled = false;
                buttonClientTo.Enabled = false;
                textBoxDateFrom.Enabled = false;
                textBoxDateTo.Enabled = false;
                clisysnbrTo = 0;
                textBoxClientTo.Text = "";
                richTextBox1.Text = "By choosing this option, you can select a range of bills that you wish to apply this receipt to. This will display all matters associated with the selected bills as long as they have an outstanding balance.";
            }
        }

        private void radioButtonClient_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonClient.Checked)
            {
                textBoxBillFrom.Enabled = false;
                textBoxBillTo.Enabled = false;
                textBoxClientTo.Enabled = true;
                buttonClientTo.Enabled = true;
                textBoxDateFrom.Enabled = false;
                textBoxDateTo.Enabled = false;
                textBoxBillFrom.Text = "";
                textBoxBillTo.Text = "";
                textBoxDateFrom.Text = "";
                textBoxDateTo.Text = "";
                richTextBox1.Text = "WARNING: When choosing this option, clients must be contiguous (meaning you would want the check applied to all clients between AND including the From and TO clients you choose). Failure to choose contiguous clients can cause numerous issues including data corruption." +
                    "\r\n" + "If you are certain this option is the one you need, enter your TO client code and a date range of bills you want to apply this receipt against. All clients between your selected From client (right below Payor in the tool) to the selected TO client whos bills fall in the date range will be selected.";
            }
        }

        private void radioButtonDateRange_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonDateRange.Checked)
            {
                textBoxBillFrom.Enabled = false;
                textBoxBillTo.Enabled = false;
                textBoxClientTo.Enabled = false;
                buttonClientTo.Enabled = false;
                textBoxDateFrom.Enabled = true;
                textBoxDateTo.Enabled = true;
                clisysnbrTo = 0;
                textBoxClientTo.Text = "";
                textBoxBillFrom.Text = "";
                textBoxBillTo.Text = "";
                richTextBox1.Text = "By choosing this option, you can select a range of bills based on the bill date that you wish to apply this receipt to. This will display all matters associated with the selected bills as long as they have an outstanding balance.";
            }
        }

        private void buttonBills_Click(object sender, EventArgs e)
        {
            if (verifyBoxes())//passed all initial checks
                if (verifyText()) //payor and matter checks
                    if (verifyLeftOverAllocs()) //ensure proper selections have been made
                    {
                        //all checks have been passed, load bill list form



                    }
        }

        private void buttonBack_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private bool verifyBoxes()
        {
            if (verifyDate(textBoxCheckDate.Text))
                if (verifyDate(textBoxDepDate.Text))
                    if (verifyMoney(textBoxCheckAmount.Text.Replace(",", "").Replace("$", "")))
                        if (!checkBoxPctOfBill.Checked || (checkBoxPctOfBill.Checked && verifyPercent(textBoxAllocPct.Text)))
                            if (verifyRangeSelection())
                                return true;
                            else
                            {
                                return false;
                            }
                        else
                        {
                            MessageBox.Show("The Allocation percent must be a valid number and be between 1 and 100", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }
                    else
                    {
                        MessageBox.Show("The Check Amount must be a valid dollar amount (number)", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                else
                {
                    MessageBox.Show("The Deposit Date must be a valid date (MM/DD/YYYY)", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            else
            {
                MessageBox.Show("The Check Date must be a valid date (MM/DD/YYYY)", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

        }

        private bool verifyRangeSelection()
        {
            if (radioButtonBillRange.Checked)
            {
                if (String.IsNullOrEmpty(textBoxBillFrom.Text) || String.IsNullOrEmpty(textBoxBillTo.Text))
                {
                    MessageBox.Show("To and From Bill Numbers must be filled in when selecting Bill Range", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else
                    return true;
            }
            else if (radioButtonClient.Checked)
            {
                if (String.IsNullOrEmpty(textBoxClient.Text) || String.IsNullOrEmpty(textBoxClientTo.Text))
                {
                    MessageBox.Show("Both Client boxes must have a client selected", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else
                {
                    if (string.IsNullOrEmpty(textBoxDateFrom.Text) || string.IsNullOrEmpty(textBoxDateTo.Text))
                    {
                        MessageBox.Show("To and From Dates must be filled in when selecting Date Range", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                    else
                    {
                        if (verifyDate(textBoxDateFrom.Text) && verifyDate(textBoxDateTo.Text))
                            return true;
                        else
                        {
                            MessageBox.Show("To and From Dates must be a proper date format (MM/DD/YYYY)", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }

                    }
                }


            }
            else // by date range
            {
                if (string.IsNullOrEmpty(textBoxDateFrom.Text) || string.IsNullOrEmpty(textBoxDateTo.Text))
                {
                    MessageBox.Show("To and From Dates must be filled in when selecting Date Range", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else
                {
                    if (verifyDate(textBoxDateFrom.Text) && verifyDate(textBoxDateTo.Text))
                        return true;
                    else
                    {
                        MessageBox.Show("To and From Dates must be a proper date format (MM/DD/YYYY)", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                }

            }
        }

        private bool verifyText()
        {
            if (!String.IsNullOrEmpty(textBoxPayor.Text) || checkBoxUseCliNickName.Checked)
                if (!checkBoxUseMatter.Checked || (checkBoxUseMatter.Checked && !string.IsNullOrEmpty(textBoxMatter.Text)))
                    return true;
                else
                {
                    MessageBox.Show("When the Use Matter Checkbox is selected, a matter must be chosen", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            else
            {
                MessageBox.Show("Payor is a required field. Please add a Payor or check the 'Use Client NickName' box", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;

            }
        }

        private bool verifyLeftOverAllocs()
        {
            if (radioButtonTrust.Checked)
            {
                string bCode = this.comboBoxBank.GetItemText(this.comboBoxBank.SelectedItem).Split(' ')[0];
                string SQLPC2 = "select * from bankaccount where bnkcode = '" + bCode + "' and BnkAcctType = 'T'";
                DataSet myRSPC2 = _jurisUtility.RecordsetFromSQL(SQLPC2);

                if (myRSPC2 == null || myRSPC2.Tables.Count == 0 || myRSPC2.Tables[0].Rows.Count == 0)
                {
                    MessageBox.Show("Choosing Trust for remaining balance requires a Trust account. Please check your selection", "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else
                    return true;
            }
            else if (radioButtonPPD.Checked)
            {
                string bCode = this.comboBoxBank.GetItemText(this.comboBoxBank.SelectedItem).Split(' ')[0];
                string SQLPC2 = "select * from bankaccount where bnkcode = '" + bCode + "' and BnkAcctType = 'O'";
                DataSet myRSPC2 = _jurisUtility.RecordsetFromSQL(SQLPC2);

                if (myRSPC2 == null || myRSPC2.Tables.Count == 0 || myRSPC2.Tables[0].Rows.Count == 0)
                {
                    MessageBox.Show("Choosing PPD for remaining balance requires an Operating account. Please check your selection", "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else
                    return true;
            }
            else if (radioButtonOther.Checked)
            {
                string bCode = this.comboBoxBank.GetItemText(this.comboBoxBank.SelectedItem).Split(' ')[0];
                string SQLPC2 = "select * from bankaccount where bnkcode = '" + bCode + "' and BnkAcctType = 'O'";
                DataSet myRSPC2 = _jurisUtility.RecordsetFromSQL(SQLPC2);

                if (myRSPC2 == null || myRSPC2.Tables.Count == 0 || myRSPC2.Tables[0].Rows.Count == 0)
                {
                    MessageBox.Show("Choosing PPD for remaining balance requires an Operating account. Please check your selection", "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else
                {
                    if (!string.IsNullOrEmpty(textBoxRef.Text))
                        return true;
                    else
                    {
                        MessageBox.Show("Choosing Other for remaining balance requires that reference be filled in", "Data Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
                    
            }
            else
                return true;
        }

        private bool verifyDate(string dt)
        {
            try
            {
                if (!string.IsNullOrEmpty(dt))
                {
                    Convert.ToDateTime(dt);
                    return true;
                }
                return false;
            }
            catch (Exception) { return false; }

        }

        private bool verifyMoney(string mn)
        {
            try
            {
                if (!string.IsNullOrEmpty(mn))
                {
                    Convert.ToDecimal(mn);
                    return true;
                }
                return false;
            }
            catch (Exception) { return false; }

        }

        private bool verifyPercent(string mn)
        {
            try
            {
                if (!string.IsNullOrEmpty(mn))
                {
                    decimal ff = Convert.ToDecimal(mn);
                    if (ff < 0 || ff > 100)
                        return false;
                    else
                        return true;
                }
                return false;
            }
            catch (Exception) { return false; }

        }

        private void radioButtonPPD_CheckedChanged(object sender, EventArgs e)
        {
            comboBoxBank.Enabled = false;
            comboBoxClient.Enabled = true;
            comboBoxGL.Enabled = false;
            comboBoxMatter.Enabled = true;
            textBoxRef.Enabled = false;
        }

        private void radioButtonManual_CheckedChanged(object sender, EventArgs e)
        {
            comboBoxBank.Enabled = false;
            comboBoxClient.Enabled = false;
            comboBoxGL.Enabled = false;
            comboBoxMatter.Enabled = false;
            textBoxRef.Enabled = false;
        }

        private void radioButtonTrust_CheckedChanged(object sender, EventArgs e)
        {
            comboBoxBank.Enabled = true;
            comboBoxClient.Enabled = true;
            comboBoxGL.Enabled = false;
            comboBoxMatter.Enabled = true;
            textBoxRef.Enabled = false;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            comboBoxBank.Enabled = true;
            comboBoxClient.Enabled = false;
            comboBoxGL.Enabled = true;
            comboBoxMatter.Enabled = false;
            textBoxRef.Enabled = true;
        }
    }
}
