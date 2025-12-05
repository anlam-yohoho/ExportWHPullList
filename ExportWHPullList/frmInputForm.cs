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
    public partial class frmInputForm : Form
    {
        public string InputText
        {
            get { return txtbInput.Text; }
            set { txtbInput.Text = value; }
        }

        public frmInputForm()
        {
            InitializeComponent();
            this.AcceptButton = btnOK;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            InputText = txtbInput.Text;
            this.DialogResult = DialogResult.OK;
            this.Hide();
        }

        public void SetForm(string title, string message)
        {
            this.Text = title;
            lblMessage.Text = message;
        }
    }
}
