namespace GM.XrmToolBox.UserRoleMatrix
{
    partial class MyPluginControl
    {
        private System.ComponentModel.IContainer components = null;

        private System.Windows.Forms.ToolStrip tsMain;
        private System.Windows.Forms.ToolStrip tsFilters;

        private System.Windows.Forms.ToolStripButton tsbLoadUsersRoles;
        private System.Windows.Forms.ToolStripButton tsbLoadOwnerTeamsRoles;
        private System.Windows.Forms.ToolStripButton tsbSelectAllRecords;
        private System.Windows.Forms.ToolStripButton tsbAddUserRole;
        private System.Windows.Forms.ToolStripButton tsbAddUserToTeam;
        private System.Windows.Forms.ToolStripButton tsbDel;

        private System.Windows.Forms.ToolStripDropDownButton tsddExport;
        private System.Windows.Forms.ToolStripMenuItem tsmiExportCsv;
        private System.Windows.Forms.ToolStripMenuItem tsmiExportExcel;

        private System.Windows.Forms.ToolStripLabel tslView;
        private System.Windows.Forms.ToolStripComboBox tscView;

        private System.Windows.Forms.ToolStripLabel tslBusinessUnit;
        private System.Windows.Forms.ToolStripComboBox tscBusinessUnit;

        private System.Windows.Forms.ToolStripLabel tslTeam;
        private System.Windows.Forms.ToolStripComboBox tscTeam;

        private System.Windows.Forms.ToolStripLabel tslAssignment;
        private System.Windows.Forms.ToolStripComboBox tscAssignment;

        private System.Windows.Forms.ToolStripLabel tslTeamBU;
        private System.Windows.Forms.ToolStripComboBox tscTeamBU;

        private System.Windows.Forms.ToolStripLabel tslRoleBU;
        private System.Windows.Forms.ToolStripComboBox tscRoleBU;

        private System.Windows.Forms.ToolStripLabel tslSearch;
        private System.Windows.Forms.ToolStripTextBox tstSearch;

        private System.Windows.Forms.ToolStripButton tsbAbout;
        private System.Windows.Forms.ToolStripLabel tslCount;

        private System.Windows.Forms.DataGridView dgvResults;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.tsMain = new System.Windows.Forms.ToolStrip();
            this.tsFilters = new System.Windows.Forms.ToolStrip();
            this.tsbLoadUsersRoles = new System.Windows.Forms.ToolStripButton();
            this.tsbLoadOwnerTeamsRoles = new System.Windows.Forms.ToolStripButton();
            this.tsbSelectAllRecords = new System.Windows.Forms.ToolStripButton();
            this.tsbAddUserRole = new System.Windows.Forms.ToolStripButton();
            this.tsbAddUserToTeam = new System.Windows.Forms.ToolStripButton();
            this.tsbDel = new System.Windows.Forms.ToolStripButton();
            this.tsddExport = new System.Windows.Forms.ToolStripDropDownButton();
            this.tsmiExportCsv = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiExportExcel = new System.Windows.Forms.ToolStripMenuItem();
            this.tslView = new System.Windows.Forms.ToolStripLabel();
            this.tscView = new System.Windows.Forms.ToolStripComboBox();
            this.tslBusinessUnit = new System.Windows.Forms.ToolStripLabel();
            this.tscBusinessUnit = new System.Windows.Forms.ToolStripComboBox();
            this.tslTeam = new System.Windows.Forms.ToolStripLabel();
            this.tscTeam = new System.Windows.Forms.ToolStripComboBox();
            this.tslAssignment = new System.Windows.Forms.ToolStripLabel();
            this.tscAssignment = new System.Windows.Forms.ToolStripComboBox();
            this.tslTeamBU = new System.Windows.Forms.ToolStripLabel();
            this.tscTeamBU = new System.Windows.Forms.ToolStripComboBox();
            this.tslRoleBU = new System.Windows.Forms.ToolStripLabel();
            this.tscRoleBU = new System.Windows.Forms.ToolStripComboBox();
            this.tslSearch = new System.Windows.Forms.ToolStripLabel();
            this.tstSearch = new System.Windows.Forms.ToolStripTextBox();
            this.tsbAbout = new System.Windows.Forms.ToolStripButton();
            this.tslCount = new System.Windows.Forms.ToolStripLabel();
            this.dgvResults = new System.Windows.Forms.DataGridView();
            this.tsMain.SuspendLayout();
            this.tsFilters.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvResults)).BeginInit();
            this.SuspendLayout();
            // 
            // tsMain – Level 1: Actions + right-aligned info
            // 
            this.tsMain.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.tsMain.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.tsMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsbAbout,
            this.tslCount,
            this.tsbLoadUsersRoles,
            this.tsbLoadOwnerTeamsRoles,
            this.tsbSelectAllRecords,
            this.tsbAddUserRole,
            this.tsbAddUserToTeam,
            this.tsbDel,
            this.tsddExport});
            this.tsMain.Location = new System.Drawing.Point(0, 0);
            this.tsMain.Name = "tsMain";
            this.tsMain.Size = new System.Drawing.Size(1200, 27);
            this.tsMain.TabIndex = 0;
            this.tsMain.Text = "Main";
            // 
            // tsFilters – Level 2: Filters
            // 
            this.tsFilters.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.tsFilters.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.tsFilters.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tslView,
            this.tscView,
            this.tslBusinessUnit,
            this.tscBusinessUnit,
            this.tslTeam,
            this.tscTeam,
            this.tslAssignment,
            this.tscAssignment,
            this.tslTeamBU,
            this.tscTeamBU,
            this.tslRoleBU,
            this.tscRoleBU,
            this.tslSearch,
            this.tstSearch});
            this.tsFilters.Location = new System.Drawing.Point(0, 27);
            this.tsFilters.Name = "tsFilters";
            this.tsFilters.Size = new System.Drawing.Size(1200, 27);
            this.tsFilters.TabIndex = 2;
            this.tsFilters.Text = "Filters";
            // 
            // tsbLoadUsersRoles
            // 
            this.tsbLoadUsersRoles.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbLoadUsersRoles.Name = "tsbLoadUsersRoles";
            this.tsbLoadUsersRoles.Size = new System.Drawing.Size(141, 24);
            this.tsbLoadUsersRoles.Text = "Load Users && Roles";
            // 
            // tsbLoadOwnerTeamsRoles
            // 
            this.tsbLoadOwnerTeamsRoles.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbLoadOwnerTeamsRoles.Name = "tsbLoadOwnerTeamsRoles";
            this.tsbLoadOwnerTeamsRoles.Size = new System.Drawing.Size(179, 24);
            this.tsbLoadOwnerTeamsRoles.Text = "Load Owner Teams Roles";
            // 
            // tsbSelectAllRecords
            // 
            this.tsbSelectAllRecords.CheckOnClick = true;
            this.tsbSelectAllRecords.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbSelectAllRecords.Name = "tsbSelectAllRecords";
            this.tsbSelectAllRecords.Size = new System.Drawing.Size(137, 24);
            this.tsbSelectAllRecords.Text = "Select/Unselect All";
            this.tsbSelectAllRecords.ToolTipText = "Select/Deselect all visible records";
            // 
            // tsbAddUserRole
            // 
            this.tsbAddUserRole.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbAddUserRole.Name = "tsbAddUserRole";
            this.tsbAddUserRole.Size = new System.Drawing.Size(108, 24);
            this.tsbAddUserRole.Text = "Add User Role";
            // 
            // tsbAddUserToTeam
            // 
            this.tsbAddUserToTeam.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbAddUserToTeam.Name = "tsbAddUserToTeam";
            this.tsbAddUserToTeam.Size = new System.Drawing.Size(152, 24);
            this.tsbAddUserToTeam.Text = "Add User to Owner Team";
            // 
            // tsbDel
            // 
            this.tsbDel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsbDel.Name = "tsbDel";
            this.tsbDel.Size = new System.Drawing.Size(36, 24);
            this.tsbDel.Text = "Del";
            // 
            // tsddExport
            // 
            this.tsddExport.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.tsddExport.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiExportCsv,
            this.tsmiExportExcel});
            this.tsddExport.Name = "tsddExport";
            this.tsddExport.Size = new System.Drawing.Size(66, 24);
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
            // tslView
            // 
            this.tslView.Name = "tslView";
            this.tslView.Size = new System.Drawing.Size(44, 24);
            this.tslView.Text = "View:";
            // 
            // tscView
            // 
            this.tscView.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.tscView.Name = "tscView";
            this.tscView.Size = new System.Drawing.Size(220, 27);
            // 
            // tslBusinessUnit
            // 
            this.tslBusinessUnit.Name = "tslBusinessUnit";
            this.tslBusinessUnit.Size = new System.Drawing.Size(98, 24);
            this.tslBusinessUnit.Text = "Business Unit:";
            // 
            // tscBusinessUnit
            // 
            this.tscBusinessUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.tscBusinessUnit.Name = "tscBusinessUnit";
            this.tscBusinessUnit.Size = new System.Drawing.Size(190, 27);
            // 
            // tslTeam – hidden, only used in Team mode
            // 
            this.tslTeam.Name = "tslTeam";
            this.tslTeam.Size = new System.Drawing.Size(48, 24);
            this.tslTeam.Text = "Team:";
            this.tslTeam.Visible = false;
            // 
            // tscTeam – hidden, only used in Team mode
            // 
            this.tscTeam.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.tscTeam.Name = "tscTeam";
            this.tscTeam.Size = new System.Drawing.Size(190, 27);
            this.tscTeam.Visible = false;
            // 
            // tslAssignment
            // 
            this.tslAssignment.Name = "tslAssignment";
            this.tslAssignment.Size = new System.Drawing.Size(89, 24);
            this.tslAssignment.Text = "Assignment:";
            // 
            // tscAssignment
            // 
            this.tscAssignment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.tscAssignment.Name = "tscAssignment";
            this.tscAssignment.Size = new System.Drawing.Size(110, 27);
            // 
            // tslTeamBU
            // 
            this.tslTeamBU.Name = "tslTeamBU";
            this.tslTeamBU.Size = new System.Drawing.Size(98, 24);
            this.tslTeamBU.Text = "Team BU:";
            // 
            // tscTeamBU
            // 
            this.tscTeamBU.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.tscTeamBU.Name = "tscTeamBU";
            this.tscTeamBU.Size = new System.Drawing.Size(160, 27);
            // 
            // tslRoleBU
            // 
            this.tslRoleBU.Name = "tslRoleBU";
            this.tslRoleBU.Size = new System.Drawing.Size(98, 24);
            this.tslRoleBU.Text = "Role BU:";
            // 
            // tscRoleBU
            // 
            this.tscRoleBU.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.tscRoleBU.Name = "tscRoleBU";
            this.tscRoleBU.Size = new System.Drawing.Size(160, 27);
            // 
            // tslSearch
            // 
            this.tslSearch.Name = "tslSearch";
            this.tslSearch.Size = new System.Drawing.Size(56, 24);
            this.tslSearch.Text = "Search:";
            // 
            // tstSearch
            // 
            this.tstSearch.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.tstSearch.Name = "tstSearch";
            this.tstSearch.Size = new System.Drawing.Size(220, 27);
            // 
            // tsbAbout
            // 
            this.tsbAbout.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.tsbAbout.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.tsbAbout.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.tsbAbout.Name = "tsbAbout";
            this.tsbAbout.Size = new System.Drawing.Size(29, 24);
            this.tsbAbout.Text = "About";
            this.tsbAbout.ToolTipText = "About";
            // 
            // tslCount
            // 
            this.tslCount.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.tslCount.Name = "tslCount";
            this.tslCount.Size = new System.Drawing.Size(59, 24);
            this.tslCount.Text = "Rows: 0";
            // 
            // dgvResults
            // 
            this.dgvResults.ColumnHeadersHeight = 29;
            this.dgvResults.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvResults.Location = new System.Drawing.Point(0, 54);
            this.dgvResults.Name = "dgvResults";
            this.dgvResults.RowHeadersWidth = 51;
            this.dgvResults.Size = new System.Drawing.Size(1200, 646);
            this.dgvResults.TabIndex = 1;
            // 
            // MyPluginControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.dgvResults);
            this.Controls.Add(this.tsFilters);
            this.Controls.Add(this.tsMain);
            this.Name = "MyPluginControl";
            this.Size = new System.Drawing.Size(1200, 700);
            this.tsMain.ResumeLayout(false);
            this.tsMain.PerformLayout();
            this.tsFilters.ResumeLayout(false);
            this.tsFilters.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvResults)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}
