using PowerAnalyzer.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerAnalyzer
{
    public partial class ChangeClientAuthentication : Form
    {
        public ChangeClientAuthentication()
        {
            InitializeComponent();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            if(txtUsername.Text.Trim() == "" || txtPassword.Text.Trim() == "")
            {
                MessageBox.Show("Username And Password Required");
            }
            else
            {
                Prefs.Set("username", txtUsername.Text);
                Prefs.Set("password", txtPassword.Text);
                MessageBox.Show("Creditial Changed Successfully");
                Hide();
            }
        }
    }
}
