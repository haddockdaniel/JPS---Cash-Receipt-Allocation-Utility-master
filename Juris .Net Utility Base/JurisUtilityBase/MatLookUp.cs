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
    public partial class MatLookUp : Form
    {
        public MatLookUp(JurisUtility jutil, System.Drawing.Point ppt, int clisys)
        {
            InitializeComponent();
            _jurisUtility = jutil;
            pt = ppt;
            clisysnbr = clisys;
        }

        JurisUtility _jurisUtility;
        private System.Drawing.Point pt;
        public bool matterSelected = false;
        public int clisysnbr = 0;
        public string clicode = "";
        public int matsysnbr = 0;
        public string matcode = "";

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void buttonExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int id = 0;
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows.Count > 1)
                MessageBox.Show("Please select one matter to proceed", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                int index = e.RowIndex;
                dataGridView1.Rows[index].Selected = true;
                matsysnbr = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                matcode = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[1].Value.ToString();
                matterSelected = true;
            }
        }

        private void buttonCreateClient_Click(object sender, EventArgs e)
        {
            if (!matterSelected)
                MessageBox.Show("Please select one matter to proceed", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {

                this.Hide();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int id = 0;
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows.Count > 1)
                MessageBox.Show("Please select one matter to proceed", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                int index = e.RowIndex;
                dataGridView1.Rows[index].Selected = true;
                matsysnbr = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                matcode = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[1].Value.ToString();
                matterSelected = true;
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int id = 0;
            if (dataGridView1.SelectedRows.Count == 0 || dataGridView1.SelectedRows.Count > 1)
                MessageBox.Show("Please select one matter to proceed", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                int index = e.RowIndex;
                dataGridView1.Rows[index].Selected = true;
                matsysnbr = Convert.ToInt32(dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value);
                matcode = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[1].Value.ToString();
                matterSelected = true;
            }
        }

        private void MatLookUp_Load(object sender, EventArgs e)
        {
            string sql = "";
                sql = "select matsysnbr, dbo.jfn_FormatMatterCode(matcode) as MatterCode, matreportingname as MatterName from matter where matclinbr = " + clisysnbr.ToString() + " order by dbo.jfn_FormatMatterCode(matcode)";
            DataSet ds = _jurisUtility.RecordsetFromSQL(sql);

            dataGridView1.DataSource = ds.Tables[0];
            dataGridView1.Columns[0].Width = 1;
            dataGridView1.Columns[1].Width = 50;
            dataGridView1.Columns[2].Width = 250;
            this.dataGridView1.Columns[0].Visible = false;
        }



    }
}
