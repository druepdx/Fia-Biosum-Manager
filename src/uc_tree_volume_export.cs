using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
    public partial class uc_tree_volume_export : UserControl
    {
        private GisTools m_oGisTools = new GisTools();
        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultTreatmentOptimizerFile;
        string m_strMasterDb;
        string m_strSourceField;
        public bool bTerminateLoad = false;

        public uc_tree_volume_export(frmFCSTreeVolumeEdit frmTreeVolume)
        {
            InitializeComponent();
            this.m_oEnv = new env();
            load_values();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            ((frmDialog)ParentForm).ParentControl.Enabled = true;
            this.ParentForm.Close();
        }

        public void load_values()
        {

        }
        private void btnExportDirectory_Click(object sender, System.EventArgs e)
        {
            DialogResult result = this.folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                string strTemp = this.folderBrowserDialog1.SelectedPath;

                if (strTemp.Length > 0)
                {
                    this.txtRootDirectory.Text = strTemp;
                }
            }
        }
        //@ToDo: Enable if we get help for this simple screen
        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "TREATMENT_OPTIMIZER", "GIS_DATA" });
        }
    }


}
