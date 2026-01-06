namespace FIA_Biosum_Manager
{
    partial class uc_tree_volume_export
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_tree_volume_export));
            this.lblTitle = new System.Windows.Forms.Label();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.btnHelp = new System.Windows.Forms.Button();
            this.btnExportDirectory = new System.Windows.Forms.Button();
            this.lblRootDirectory = new System.Windows.Forms.Label();
            this.txtRootDirectory = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.lblDatabaseName = new System.Windows.Forms.Label();
            this.btnExport = new System.Windows.Forms.Button();
            this.lblTableName = new System.Windows.Forms.Label();
            this.txtTableName = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(0, 0);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(750, 32);
            this.lblTitle.TabIndex = 27;
            this.lblTitle.Text = "Export Biomass and Volume Values";
            // 
            // BtnCancel
            // 
            this.BtnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnCancel.Location = new System.Drawing.Point(292, 335);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(165, 33);
            this.BtnCancel.TabIndex = 66;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(18, 331);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(115, 37);
            this.btnHelp.TabIndex = 67;
            this.btnHelp.Text = "Help";
            this.btnHelp.Visible = false;
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // btnExportDirectory
            // 
            this.btnExportDirectory.Image = ((System.Drawing.Image)(resources.GetObject("btnExportDirectory.Image")));
            this.btnExportDirectory.Location = new System.Drawing.Point(638, 68);
            this.btnExportDirectory.Name = "btnExportDirectory";
            this.btnExportDirectory.Size = new System.Drawing.Size(32, 32);
            this.btnExportDirectory.TabIndex = 71;
            this.btnExportDirectory.Text = "Export to Folder";
            this.btnExportDirectory.Click += new System.EventHandler(this.btnExportDirectory_Click);
            // 
            // lblRootDirectory
            // 
            this.lblRootDirectory.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRootDirectory.Location = new System.Drawing.Point(5, 76);
            this.lblRootDirectory.Name = "lblRootDirectory";
            this.lblRootDirectory.Size = new System.Drawing.Size(150, 20);
            this.lblRootDirectory.TabIndex = 69;
            this.lblRootDirectory.Text = "Export to Folder:";
            // 
            // txtRootDirectory
            // 
            this.txtRootDirectory.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtRootDirectory.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRootDirectory.Location = new System.Drawing.Point(206, 70);
            this.txtRootDirectory.Name = "txtRootDirectory";
            this.txtRootDirectory.Size = new System.Drawing.Size(416, 26);
            this.txtRootDirectory.TabIndex = 70;
            // 
            // lblDatabaseName
            // 
            this.lblDatabaseName.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDatabaseName.Location = new System.Drawing.Point(5, 110);
            this.lblDatabaseName.Name = "lblDatabaseName";
            this.lblDatabaseName.Size = new System.Drawing.Size(400, 20);
            this.lblDatabaseName.TabIndex = 72;
            this.lblDatabaseName.Text = "Database Name:";
            // 
            // btnExport
            // 
            this.btnExport.Enabled = false;
            this.btnExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExport.Location = new System.Drawing.Point(479, 335);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(165, 33);
            this.btnExport.TabIndex = 73;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.BtnExport_Click);
            // 
            // lblTableName
            // 
            this.lblTableName.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTableName.Location = new System.Drawing.Point(5, 144);
            this.lblTableName.Name = "lblTableName";
            this.lblTableName.Size = new System.Drawing.Size(150, 20);
            this.lblTableName.TabIndex = 74;
            this.lblTableName.Text = "Table Name:";
            // 
            // txtTableName
            // 
            this.txtTableName.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtTableName.Location = new System.Drawing.Point(206, 144);
            this.txtTableName.MaxLength = 20;
            this.txtTableName.Name = "txtTableName";
            this.txtTableName.Size = new System.Drawing.Size(166, 26);
            this.txtTableName.TabIndex = 75;
            this.txtTableName.TextChanged += new System.EventHandler(this.txtTableName_TextChanged);
            // 
            // uc_tree_volume_export
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.txtTableName);
            this.Controls.Add(this.lblTableName);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.lblDatabaseName);
            this.Controls.Add(this.btnExportDirectory);
            this.Controls.Add(this.lblRootDirectory);
            this.Controls.Add(this.txtRootDirectory);
            this.Controls.Add(this.btnHelp);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.lblTitle);
            this.Name = "uc_tree_volume_export";
            this.Size = new System.Drawing.Size(750, 457);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.Button btnHelp;
        private System.Windows.Forms.Button btnExportDirectory;
        private System.Windows.Forms.Label lblRootDirectory;
        public System.Windows.Forms.TextBox txtRootDirectory;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label lblDatabaseName;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Label lblTableName;
        public System.Windows.Forms.TextBox txtTableName;
    }
}
