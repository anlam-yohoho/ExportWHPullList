using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportWHPullList
{
    public partial class frmProgressBar : Form
    {
        public event Action ButtonAbortClicked;
        public frmProgressBar()
        {
            InitializeComponent();
        }

        private void btnAbort_Click(object sender, EventArgs e)
        {
            ButtonAbortClicked?.Invoke();
            this.Close();
        }

        public void UpdateProgress(int value, string message = "")
        {
            progressBar.Value = value;
            lblProgress.Text = $"Progress: {value}%";
            lblStatus.Text = message;
            Application.DoEvents(); // Forces UI refresh
            if (progressBar.Value == 100)
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }
    }
}
