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
        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultTreatmentOptimizerFile;
        public bool bTerminateLoad = false;
        private frmFCSTreeVolumeEdit m_frmParent;
        public uc_tree_volume_export(frmFCSTreeVolumeEdit frmTreeVolume)
        {
            InitializeComponent();
            this.m_oEnv = new env();
            this.m_frmParent = frmTreeVolume;
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
            string strConn = dataMgr.GetConnectionString(strFile);
            if (System.IO.File.Exists(strFile))
            {
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
            // First create the copied table
            string strSourceTable = m_frmParent.GridTableSource;
            string strCreateTableSql = "";
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dataMgr.GetConnectionString(m_frmParent.GridDatabaseSource)))
            {
                conn.Open();
                dataMgr.SqlQueryReader(conn, $"SELECT sql FROM sqlite_master WHERE type='table' AND name='{strSourceTable}'");
                if (dataMgr.m_DataReader.HasRows)
                {
                    while (dataMgr.m_DataReader.Read())
                    {
                        string strSql = Convert.ToString(dataMgr.m_DataReader["sql"]);
                        string strRenameTableSql = strSql.Replace(strSourceTable, txtTableName.Text.Trim());
                        //@ToDo: Put this in to handle field name translation from FICS. Can remove when we disable FICS
                        strCreateTableSql = strRenameTableSql.Replace("_calc", "");
                    }
                    dataMgr.m_DataReader.Close();
                }
            }
            frmMain.g_oFrmMain.ActivateStandByAnimation(
                frmMain.g_oFrmMain.WindowState,
                frmMain.g_oFrmMain.Left,
                frmMain.g_oFrmMain.Height,
                frmMain.g_oFrmMain.Width,
                frmMain.g_oFrmMain.Top);
            frmMain.g_sbpInfo.Text = "Exporting Tree Data...Stand by";
            btnExport.Enabled = false;
            BtnCancel.Enabled = false;

            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                conn.Open();
                if (!string.IsNullOrEmpty(strCreateTableSql))
                {
                    dataMgr.SqlNonQuery(conn, strCreateTableSql);
                }

                // 1. Query the destination table to get its structure
               string selectQuery = $"SELECT * FROM {txtTableName.Text.Trim()} WHERE 1=0"; // WHERE 1=0 returns schema, no rows
               using (System.Data.SQLite.SQLiteDataAdapter adapter = new System.Data.SQLite.SQLiteDataAdapter(selectQuery, conn))
                {
                    using (System.Data.SQLite.SQLiteCommandBuilder commandBuilder = new System.Data.SQLite.SQLiteCommandBuilder(adapter))
                    {
                        // Ensure proper quoting for SQLite table/column names
                        commandBuilder.QuotePrefix = "[";
                        commandBuilder.QuoteSuffix = "]";

                        // Get the auto-generated INSERT command
                        //adapter.GetInsertCommand();

                        // 2. Create an empty "destination" DataTable
                        DataTable destinationTable = new DataTable();
                        adapter.Fill(destinationTable);

                        // 3. Adapt the destination table with data from your source DataTable
                        IList<object> lstItem = new List<object>();
                        IList<string> sourceColumnNames = new List<string>();
                        for (int i = 0; i < m_frmParent.GridDataTable.Columns.Count; i++)
                        {
                            sourceColumnNames.Add(m_frmParent.GridDataTable.Columns[i].ColumnName.ToUpper());
                        }
                        
                        System.Data.SQLite.SQLiteTransaction destTransaction = conn.BeginTransaction();
                        foreach (DataRow row in m_frmParent.GridDataTable.Rows)
                        {
                            DataRow destinationRow = destinationTable.NewRow();
                            // Add field values to a list from which we will create a new object array
                            lstItem.Clear();
                            for (int i = 0; i < destinationTable.Columns.Count; i++)
                            {
                                if (sourceColumnNames.Contains(destinationTable.Columns[i].ColumnName.ToUpper()))
                                {
                                    lstItem.Add(row[destinationTable.Columns[i].ColumnName]);
                                }
                                else
                                {
                                    lstItem.Add(null);
                                }
                            }
                            object[] rowArray = lstItem.ToArray();
                            destinationRow.ItemArray = rowArray;
                            destinationTable.Rows.Add(destinationRow);
                        }

                        // 4. Wrap the update in a transaction for performance and integrity
                        // The Update() method handles individual row inserts in a batch within an implicit transaction
                        // Alternatively, you can manage an explicit transaction using connection.BeginTransaction()
                        adapter.Update(destinationTable);
                        destTransaction.Commit();
                        destTransaction = null;
                        // The connection can be closed after the update is complete
                    }
                }

            }
            frmMain.g_sbpInfo.Text = "Ready";
            frmMain.g_oFrmMain.DeactivateStandByAnimation();
            MessageBox.Show("Grid data successfully exported to table!", "FIA Biosum");
            btnExport.Enabled = true;
            BtnCancel.Enabled = true;

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
