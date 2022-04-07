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
    public partial class ClientLookUp : Form
    {
        public ClientLookUp(JurisUtility jutil, System.Drawing.Point ppt)
        {
            InitializeComponent();
            _jurisUtility = jutil;
            pt = ppt;
        }

        JurisUtility _jurisUtility;
        private System.Drawing.Point pt;
        public bool clientSelected = false;
        public int clisysnbr = 0;
        public string clicode = "";
        public string clientName = "";

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string sql = "select clisysnbr, dbo.jfn_FormatClientCode(clicode) as clicode, clireportingname from client where clireportingname like '%" + textBoxClient.Text + "%'";
            DataSet ds = _jurisUtility.RecordsetFromSQL(sql);

            dataGridView1.DataSource = ds.Tables[0];
            dataGridView1.Columns[0].Width = 1;
            dataGridView1.Columns[1].Width = 50;
            dataGridView1.Columns[2].Width = 250;
            this.dataGridView1.Columns[0].Visible = false;
        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int id = 0;
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows.Count > 1)
                MessageBox.Show("Please select one client to proceed", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                int index = e.RowIndex;
                dataGridView1.Rows[index].Selected = true;
                clisysnbr = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                clicode = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[1].Value.ToString();
                clientName = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Value.ToString();
                clientSelected = true;
            }
        }

        private void buttonCreateClient_Click(object sender, EventArgs e)
        {
            if (!clientSelected)
                MessageBox.Show("Please select one client to proceed", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                
                this.Hide();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int id = 0;
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows.Count > 1)
                MessageBox.Show("Please select one client to proceed", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                int index = e.RowIndex;
                dataGridView1.Rows[index].Selected = true;
                clisysnbr = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                clicode = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[1].Value.ToString();
                clientName = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Value.ToString();
                clientSelected = true;
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int id = 0;
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows.Count > 1)
                MessageBox.Show("Please select one client to proceed", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                int index = e.RowIndex;
                dataGridView1.Rows[index].Selected = true;
                clisysnbr = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                clicode = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[1].Value.ToString();
                clientName = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[2].Value.ToString();
                clientSelected = true;
            }
        }
    }
}
