namespace GM.XrmToolBox.UserRoleMatrix
{
    partial class MyPluginControl
    {
        private System.ComponentModel.IContainer components = null;

        private System.Windows.Forms.ToolStrip tsMain;

        private System.Windows.Forms.ToolStripButton tsbLoadUsersRoles;
        private System.Windows.Forms.ToolStripButton tsbLoadOwnerTeamsRoles;
        private System.Windows.Forms.ToolStripButton tsbAddUserRole;
        private System.Windows.Forms.ToolStripButton tsbDel;

        private System.Windows.Forms.ToolStripDropDownButton tsddExport;
        private System.Windows.Forms.ToolStripMenuItem tsmiExportCsv;
        private System.Windows.Forms.ToolStripMenuItem tsmiExportExcel;

        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;

        private System.Windows.Forms.ToolStripLabel tslBusinessUnit;
        private System.Windows.Forms.ToolStripComboBox tscBusinessUnit;

        private System.Windows.Forms.ToolStripLabel tslTeam;
        private System.Windows.Forms.ToolStripComboBox tscTeam;

        private System.Windows.Forms.ToolStripLabel tslAssignment;
        private System.Windows.Forms.ToolStripComboBox tscAssignment;

        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;

        private System.Windows.Forms.ToolStripLabel tslSearch;
        private System.Windows.Forms.ToolStripTextBox tstSearch;

        private System.Windows.Forms.ToolStripLabel tslCount;

        private System.Windows.Forms.DataGridView dgvResults;
        
        private System.Windows.Forms.ToolStripButton tsbAbout;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.tsMain = new System.Windows.Forms.ToolStrip();
            this.tsbLoadUsersRoles = new System.Windows.Forms.ToolStripButton();
            this.tsbLoadOwnerTeamsRoles = new System.Windows.Forms.ToolStripButton();
            this.tsbAddUserRole = new System.Windows.Forms.ToolStripButton();
            this.tsbDel = new System.Windows.Forms.ToolStripButton();
            this.tsddExport = new System.Windows.Forms.ToolStripDropDownButton();
            this.tsmiExportCsv = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiExportExcel = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.tslBusinessUnit = new System.Windows.Forms.ToolStripLabel();
            this.tscBusinessUnit = new System.Windows.Forms.ToolStripComboBox();
            this.tslTeam = new System.Windows.Forms.ToolStripLabel();
            this.tscTeam = new System.Windows.Forms.ToolStripComboBox();
            this.tslAssignment = new System.Windows.Forms.ToolStripLabel();
            this.tscAssignment = new System.Windows.Forms.ToolStripComboBox();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.tslSearch = new System.Windows.Forms.ToolStripLabel();
            this.tstSearch = new System.Windows.Forms.ToolStripTextBox();
            this.tslCount = new System.Windows.Forms.ToolStripLabel();
            this.dgvResults = new System.Windows.Forms.DataGridView();
            this.tsbAbout = new System.Windows.Forms.ToolStripButton();

            this.tsMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvResults)).BeginInit();
            this.SuspendLayout();
            // 
            // tsMain
            // 
            this.tsMain.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.tsMain.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.tsMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsbLoadUsersRoles,
            this.tsbLoadOwnerTeamsRoles,
            this.tsbAddUserRole,
            this.tsbDel,
            this.tsddExport,
            this.toolStripSeparator1,
            this.tslBusinessUnit,
            this.tscBusinessUnit,
            this.tslTeam,
            this.tscTeam,
            this.tslAssignment,
            this.tscAssignment,
            this.toolStripSeparator2,
            this.tslSearch,
            this.tstSearch,
            this.tslCount,
            this.tsbAbout});
            this.tsMain.Location = new System.Drawing.Point(0, 0);
            this.tsMain.Name = "tsMain";
            this.tsMain.Size = new System.Drawing.Size(1200, 32);
            this.tsMain.TabIndex = 0;
            this.tsMain.Text = "Main";
            // 
            // tsbLoadUsersRoles
            // 
            this.tsbLoadUsersRoles.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbLoadUsersRoles.Name = "tsbLoadUsersRoles";
            this.tsbLoadUsersRoles.Size = new System.Drawing.Size(129, 29);
            this.tsbLoadUsersRoles.Text = "Load Users & Roles";
            // 
            // tsbLoadOwnerTeamsRoles
            // 
            this.tsbLoadOwnerTeamsRoles.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbLoadOwnerTeamsRoles.Name = "tsbLoadOwnerTeamsRoles";
            this.tsbLoadOwnerTeamsRoles.Size = new System.Drawing.Size(179, 29);
            this.tsbLoadOwnerTeamsRoles.Text = "Load Owner Teams Roles";
            // 
            // tsbAddUserRole
            // 
            this.tsbAddUserRole.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbAddUserRole.Name = "tsbAddUserRole";
            this.tsbAddUserRole.Size = new System.Drawing.Size(108, 29);
            this.tsbAddUserRole.Text = "Add User Role";
            // 
            // tsbDel
            // 
            this.tsbDel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbDel.Name = "tsbDel";
            this.tsbDel.Size = new System.Drawing.Size(36, 29);
            this.tsbDel.Text = "Del";
            // 
            // tsddExport
            // 
            this.tsddExport.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsddExport.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
                this.tsmiExportCsv,
                this.tsmiExportExcel
            });
            this.tsddExport.Name = "tsddExport";
            this.tsddExport.Size = new System.Drawing.Size(66, 29);
            this.tsddExport.Text = "Export";
            // 
            // tsmiExportCsv
            // 
            this.tsmiExportCsv.Name = "tsmiExportCsv";
            this.tsmiExportCsv.Size = new System.Drawing.Size(135, 26);
            this.tsmiExportCsv.Text = "CSV...";
            // 
            // tsmiExportExcel
            // 
            this.tsmiExportExcel.Name = "tsmiExportExcel";
            this.tsmiExportExcel.Size = new System.Drawing.Size(135, 26);
            this.tsmiExportExcel.Text = "Excel...";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 32);
            // 
            // tslBusinessUnit
            // 
            this.tslBusinessUnit.Name = "tslBusinessUnit";
            this.tslBusinessUnit.Size = new System.Drawing.Size(98, 29);
            this.tslBusinessUnit.Text = "Business Unit:";
            // 
            // tscBusinessUnit
            // 
            this.tscBusinessUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.tscBusinessUnit.Name = "tscBusinessUnit";
            this.tscBusinessUnit.Size = new System.Drawing.Size(170, 32);
            // 
            // tslTeam
            // 
            this.tslTeam.Name = "tslTeam";
            this.tslTeam.Size = new System.Drawing.Size(48, 29);
            this.tslTeam.Text = "Team:";
            // 
            // tscTeam
            // 
            this.tscTeam.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.tscTeam.Name = "tscTeam";
            this.tscTeam.Size = new System.Drawing.Size(170, 32);
            // 
            // tslAssignment
            // 
            this.tslAssignment.Name = "tslAssignment";
            this.tslAssignment.Size = new System.Drawing.Size(89, 29);
            this.tslAssignment.Text = "Assignment:";
            // 
            // tscAssignment
            // 
            this.tscAssignment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.tscAssignment.Name = "tscAssignment";
            this.tscAssignment.Size = new System.Drawing.Size(95, 28);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // tslSearch
            // 
            this.tslSearch.Name = "tslSearch";
            this.tslSearch.Size = new System.Drawing.Size(56, 20);
            this.tslSearch.Text = "Search:";
            // 
            // tstSearch
            // 
            this.tstSearch.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.tstSearch.Name = "tstSearch";
            this.tstSearch.Size = new System.Drawing.Size(220, 27);
            // 
            // tslCount
            // 
            this.tslCount.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.tslCount.Name = "tslCount";
            this.tslCount.Size = new System.Drawing.Size(59, 20);
            this.tslCount.Text = "Rows: 0";
            // 
            // dgvResults
            // 
            this.dgvResults.ColumnHeadersHeight = 29;
            this.dgvResults.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvResults.Location = new System.Drawing.Point(0, 32);
            this.dgvResults.Name = "dgvResults";
            this.dgvResults.RowHeadersWidth = 51;
            this.dgvResults.Size = new System.Drawing.Size(1200, 668);
            this.dgvResults.TabIndex = 1;
            // 
            // tsbAbout
            // 
            this.tsbAbout.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.tsbAbout.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsbAbout.Image = ImageFromBase64(AboutIconBase64);
            this.tsbAbout.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbAbout.Name = "tsbAbout";
            this.tsbAbout.Size = new System.Drawing.Size(29, 28);
            this.tsbAbout.Text = "About";
            this.tsbAbout.ToolTipText = "About";

            // 
            // MyPluginControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.dgvResults);
            this.Controls.Add(this.tsMain);
            this.Name = "MyPluginControl";
            this.Size = new System.Drawing.Size(1200, 700);
            this.tsMain.ResumeLayout(false);
            this.tsMain.PerformLayout();

            ((System.ComponentModel.ISupportInitialize)(this.dgvResults)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
