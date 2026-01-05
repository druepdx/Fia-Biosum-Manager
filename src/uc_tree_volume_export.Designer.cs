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
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(0, 0);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(715, 32);
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
            this.BtnCancel.Visible = false;
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
            this.btnExportDirectory.Location = new System.Drawing.Point(575, 68);
            this.btnExportDirectory.Name = "btnExportDirectory";
            this.btnExportDirectory.Size = new System.Drawing.Size(32, 32);
            this.btnExportDirectory.TabIndex = 71;
            this.btnExportDirectory.Text = "Export to Folder";
            this.btnExportDirectory.Click += new System.EventHandler(this.btnExportDirectory_Click);
            // 
            // lblRootDirectory
            // 
            this.lblRootDirectory.Location = new System.Drawing.Point(5, 76);
            this.lblRootDirectory.Name = "lblRootDirectory";
            this.lblRootDirectory.Size = new System.Drawing.Size(120, 16);
            this.lblRootDirectory.TabIndex = 69;
            this.lblRootDirectory.Text = "Export to Folder";
            // 
            // txtRootDirectory
            // 
            this.txtRootDirectory.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.txtRootDirectory.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRootDirectory.Location = new System.Drawing.Point(125, 70);
            this.txtRootDirectory.Name = "txtRootDirectory";
            this.txtRootDirectory.Size = new System.Drawing.Size(416, 26);
            this.txtRootDirectory.TabIndex = 70;
            // 
            // uc_tree_volume_export
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnExportDirectory);
            this.Controls.Add(this.lblRootDirectory);
            this.Controls.Add(this.txtRootDirectory);
            this.Controls.Add(this.btnHelp);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.lblTitle);
            this.Name = "uc_tree_volume_export";
            this.Size = new System.Drawing.Size(715, 457);
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
    }
}
