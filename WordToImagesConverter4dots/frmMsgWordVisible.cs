using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace WordToImagesConverter4dots
{
    public partial class frmMsgWordVisible : WordToImagesConverter4dots.CustomForm
    {
        public frmMsgWordVisible()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.MsgWordVisible = !chkDoNotShowAgain.Checked;

            this.DialogResult = DialogResult.OK;
        }
    }
}
