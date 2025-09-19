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
    public partial class LoginForm : Form
    {
        public LoginForm()
        {
            InitializeComponent();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            if(txtUsername.Text == "PRAM_ELECTECH_RPN" && txtPassword.Text == "RAJ84_PRA90_NIT91")
            {
                txtUsername.Text = "";
                txtPassword.Text = "";

                ChangeClientAuthentication changeClientAuthentication = new ChangeClientAuthentication();
                changeClientAuthentication.Show();
            }
            else if(txtUsername.Text == Prefs.Get("username") && txtPassword.Text == Prefs.Get("password"))
            {
                MainFormV2 mainForm = new MainFormV2();
                mainForm.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Invalid Username or Password", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
