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

        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            SetAllCheckboxes(true, this);

            btnPhase1ShowHide.Text = "Phase 1 Hide";
            btnPhase2ShowHide.Text = "Phase 2 Hide";
            btnPhase3ShowHide.Text = "Phase 3 Hide";
            btnPhase4ShowHide.Text = "Phase 4 Hide";
        }

        private void btnClearAll_Click(object sender, EventArgs e)
        {
            SetAllCheckboxes(false, this);

            btnPhase1ShowHide.Text = "Phase 1 Show";
            btnPhase2ShowHide.Text = "Phase 2 Show";
            btnPhase3ShowHide.Text = "Phase 3 Show";
            btnPhase4ShowHide.Text = "Phase 4 Show";
        }

        private void PhaseWiseCheckboxStatus(string phaseTag, bool status, Control parent)
        {
            foreach (Control ctrl in parent.Controls)
            {
                if (ctrl is Guna2CheckBox checkBox && checkBox.Tag?.ToString() == phaseTag)
                {
                    checkBox.Checked = status;
                }                
            }
        }


        private void SetAllCheckboxes(bool check, Control parent)
        {
            foreach (Control ctrl in parent.Controls)
            {
                if (ctrl is Guna2CheckBox checkBox)
                {
                    checkBox.Checked = check;
                }

                if (ctrl.HasChildren)
                {
                    SetAllCheckboxes(check, ctrl);
                }
            }
        }

        private void btnPhase1ShowHide_Click(object sender, EventArgs e)
        {
            if(btnPhase1ShowHide.Text == "Phase 1 Show")
            {
                btnPhase1ShowHide.Text = "Phase 1 Hide";

                PhaseWiseCheckboxStatus("Phase1", true, this);
            }
            else
            {
                btnPhase1ShowHide.Text = "Phase 1 Show";
                PhaseWiseCheckboxStatus("Phase1", false, this);
            }
        }

        private void btnPhase2ShowHide_Click(object sender, EventArgs e)
        {
            if (btnPhase2ShowHide.Text == "Phase 2 Show")
            {
                btnPhase2ShowHide.Text = "Phase 2 Hide";
                PhaseWiseCheckboxStatus("Phase2", true, this);
            }
            else
            {
                btnPhase2ShowHide.Text = "Phase 2 Show";
                PhaseWiseCheckboxStatus("Phase2", false, this);
            }
        }

        private void btnPhase3ShowHide_Click(object sender, EventArgs e)
        {
            if (btnPhase3ShowHide.Text == "Phase 3 Show")
            {
                btnPhase3ShowHide.Text = "Phase 3 Hide";
                PhaseWiseCheckboxStatus("Phase3", true, this);
            }
            else
            {
                btnPhase3ShowHide.Text = "Phase 3 Show";
                PhaseWiseCheckboxStatus("Phase3", false, this);
            }
        }

        private void btnPhase4ShowHide_Click(object sender, EventArgs e)
        {
            if (btnPhase4ShowHide.Text == "Phase 4 Show")
            {
                btnPhase4ShowHide.Text = "Phase 4 Hide";
                PhaseWiseCheckboxStatus("Phase4", true, this);
            }
            else
            {
                btnPhase4ShowHide.Text = "Phase 4 Show";
                PhaseWiseCheckboxStatus("Phase4", false, this);
            }
        }
    }
}
