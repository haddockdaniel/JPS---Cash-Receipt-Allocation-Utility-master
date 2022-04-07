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
    public partial class ErrorDisplay : Form
    {
        public ErrorDisplay(DataTable dt)
        {
            InitializeComponent();
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[1].Width = 850;

        }

        private void buttonBack_Click(object sender, EventArgs e)
        {
            this.Close();

        }


        private void buttonPrint_Click(object sender, EventArgs e)
        {
            var newline = System.Environment.NewLine;
            var tab = "\t";
            var clipboard_string = new StringBuilder();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    if (i == (row.Cells.Count - 1))
                        clipboard_string.Append(row.Cells[i].Value + newline);
                    else
                        clipboard_string.Append(row.Cells[i].Value + tab);
                }
            }

            Clipboard.SetText(clipboard_string.ToString());
            MessageBox.Show("Data is ready to be pasted into Excel or Text Editor", "Process Complete");
        }
    }
}
