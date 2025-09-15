using Guna.UI2.WinForms;
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
    public partial class SettingsForm : Form
    {
        private MainFormV2 mainFormV2;

        public SettingsForm(MainFormV2 mainFormV2)
        {
            InitializeComponent();
            this.mainFormV2 = mainFormV2;   
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            for (int i = 1; i <= 54; i++)
            {
                Guna2TextBox txt = Controls.Find("m" + i, true).FirstOrDefault() as Guna2TextBox;

                if (txt != null)
                {
                    txt.Text = Prefs.Get("m" + i);
                }
            }
        }

        private void btnSavePreference_Click(object sender, EventArgs e)
        {
            for (int i = 1; i <= 54; i++)
            {
                Guna2TextBox txt = Controls.Find("m" + i, true).FirstOrDefault() as Guna2TextBox;

                if (txt != null)
                {
                    Prefs.Set("m" + i, txt.Text);
                }
            }

            mainFormV2.LoadPreference();
            Dispose();
        }
    }
}
