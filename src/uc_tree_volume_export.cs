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
            this.lblDatabaseName.Text = $@"Database Name:{Tables.VolumeAndBiomass.ExportBiosumVolumesDatabase}";
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
                        if (!string.IsNullOrEmpty(txtRootDirectory.Text) && !string.IsNullOrEmpty(txtTableName.Text))
            {
                btnExportDirectory.Enabled = true;
            }
            else
            {
                btnExportDirectory.Enabled = false;
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
        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (this.txtTableName.Text.Length == 0)
            {
                MessageBox.Show("Enter a unique table name");
                this.txtTableName.Focus();;
                return;
            }
            else
            {
                System.Text.RegularExpressions.Regex rx =
                    new System.Text.RegularExpressions.Regex("^[a-zA-Z_][a-zA-Z0-9_]*$");

                System.Text.RegularExpressions.MatchCollection matches = rx.Matches(txtTableName.Text);
                if (matches.Count < 1)
                {
                    MessageBox.Show("The table name contains an invalid character. Only letters, numbers and underscores are permitted!!", "FIA Biosum");
                    this.txtTableName.Focus();
                    return;
                }
            }
            //
            //check for duplicate table name
            //
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strFile = $@"{this.txtRootDirectory.Text}\{Tables.VolumeAndBiomass.ExportBiosumVolumesDatabase}";
            if (System.IO.File.Exists(strFile))
            {
                string strConn = dataMgr.GetConnectionString(strFile);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    bool bExists = dataMgr.TableExist(conn,this.txtTableName.Text.Trim());
                    if (!bExists)
                    {
                        bExists = dataMgr.AttachedTableExist(conn, this.txtTableName.Text.Trim());
                    }
                    if (bExists)
                    {
                        MessageBox.Show("Cannot have a duplicate table name");
                        return;
                    }
                }
            }
            else
            {
                // Create database if it doesn't exist
                try
                {
                    dataMgr.CreateDbFile(strFile);
                }
                catch (Exception)
                {
                    MessageBox.Show("Unable to create the database in the directory provided. Please check the folder permissions or choose another folder.");
                    return;
                }
            }
        }
        private void txtTableName_TextChanged(object sender, EventArgs e)
        {
            ManageExportButton();
        }
        private void ManageExportButton()
        {
            if (!string.IsNullOrEmpty(txtRootDirectory.Text) && !string.IsNullOrEmpty(txtTableName.Text))
            {
                btnExport.Enabled = true;
            }
            else
            {
                btnExport.Enabled = false;
            }
        }
    }


}
