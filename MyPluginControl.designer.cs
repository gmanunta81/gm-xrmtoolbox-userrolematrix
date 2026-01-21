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
            this.tslCount});
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
            this.tsbAddUserRole.Size = new System.Drawing.Size(115, 29);
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
            this.tsmiExportExcel});
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


//namespace GM.XrmToolBox.UserRoleMatrix
//{
//    partial class MyPluginControl
//    {
//        private System.ComponentModel.IContainer components = null;

//        private System.Windows.Forms.ToolStrip tsMain;
//        private System.Windows.Forms.ToolStripButton tsbLoad;

//        private System.Windows.Forms.ToolStripDropDownButton tsddExport;
//        private System.Windows.Forms.ToolStripMenuItem tsmiExportCsv;
//        private System.Windows.Forms.ToolStripMenuItem tsmiExportExcel;

//        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;

//        private System.Windows.Forms.ToolStripLabel tslBusinessUnit;
//        private System.Windows.Forms.ToolStripComboBox tscBusinessUnit;

//        private System.Windows.Forms.ToolStripLabel tslTeam;
//        private System.Windows.Forms.ToolStripComboBox tscTeam;

//        private System.Windows.Forms.ToolStripLabel tslAssignment;
//        private System.Windows.Forms.ToolStripComboBox tscAssignment;

//        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;

//        private System.Windows.Forms.ToolStripLabel tslSearch;
//        private System.Windows.Forms.ToolStripTextBox tstSearch;

//        private System.Windows.Forms.ToolStripLabel tslCount;

//        private System.Windows.Forms.DataGridView dgvResults;

//        protected override void Dispose(bool disposing)
//        {
//            if (disposing && (components != null))
//                components.Dispose();

//            base.Dispose(disposing);
//        }

//        private void InitializeComponent()
//        {
//            this.tsMain = new System.Windows.Forms.ToolStrip();
//            this.tsbLoad = new System.Windows.Forms.ToolStripButton();
//            this.tsddExport = new System.Windows.Forms.ToolStripDropDownButton();
//            this.tsmiExportCsv = new System.Windows.Forms.ToolStripMenuItem();
//            this.tsmiExportExcel = new System.Windows.Forms.ToolStripMenuItem();
//            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
//            this.tslBusinessUnit = new System.Windows.Forms.ToolStripLabel();
//            this.tscBusinessUnit = new System.Windows.Forms.ToolStripComboBox();
//            this.tslTeam = new System.Windows.Forms.ToolStripLabel();
//            this.tscTeam = new System.Windows.Forms.ToolStripComboBox();
//            this.tslAssignment = new System.Windows.Forms.ToolStripLabel();
//            this.tscAssignment = new System.Windows.Forms.ToolStripComboBox();
//            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
//            this.tslSearch = new System.Windows.Forms.ToolStripLabel();
//            this.tstSearch = new System.Windows.Forms.ToolStripTextBox();
//            this.tslCount = new System.Windows.Forms.ToolStripLabel();
//            this.dgvResults = new System.Windows.Forms.DataGridView();
//            this.tsMain.SuspendLayout();
//            ((System.ComponentModel.ISupportInitialize)(this.dgvResults)).BeginInit();
//            this.SuspendLayout();
//            // 
//            // tsMain
//            // 
//            this.tsMain.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
//            this.tsMain.ImageScalingSize = new System.Drawing.Size(20, 20);
//            this.tsMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
//            this.tsbLoad,
//            this.tsddExport,
//            this.toolStripSeparator1,
//            this.tslBusinessUnit,
//            this.tscBusinessUnit,
//            this.tslTeam,
//            this.tscTeam,
//            this.tslAssignment,
//            this.tscAssignment,
//            this.toolStripSeparator2,
//            this.tslSearch,
//            this.tstSearch,
//            this.tslCount});
//            this.tsMain.Location = new System.Drawing.Point(0, 0);
//            this.tsMain.Name = "tsMain";
//            this.tsMain.Size = new System.Drawing.Size(1100, 31);
//            this.tsMain.TabIndex = 0;
//            this.tsMain.Text = "Main";
//            // 
//            // tsbLoad
//            // 
//            this.tsbLoad.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
//            this.tsbLoad.Name = "tsbLoad";
//            this.tsbLoad.Size = new System.Drawing.Size(129, 28);
//            this.tsbLoad.Text = "Load Users & Roles";
//            // 
//            // tsddExport
//            // 
//            this.tsddExport.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
//            this.tsddExport.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
//            this.tsmiExportCsv,
//            this.tsmiExportExcel});
//            this.tsddExport.Name = "tsddExport";
//            this.tsddExport.Size = new System.Drawing.Size(66, 28);
//            this.tsddExport.Text = "Export";
//            // 
//            // tsmiExportCsv
//            // 
//            this.tsmiExportCsv.Name = "tsmiExportCsv";
//            this.tsmiExportCsv.Size = new System.Drawing.Size(135, 26);
//            this.tsmiExportCsv.Text = "CSV...";
//            // 
//            // tsmiExportExcel
//            // 
//            this.tsmiExportExcel.Name = "tsmiExportExcel";
//            this.tsmiExportExcel.Size = new System.Drawing.Size(135, 26);
//            this.tsmiExportExcel.Text = "Excel...";
//            // 
//            // toolStripSeparator1
//            // 
//            this.toolStripSeparator1.Name = "toolStripSeparator1";
//            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 31);
//            // 
//            // tslBusinessUnit
//            // 
//            this.tslBusinessUnit.Name = "tslBusinessUnit";
//            this.tslBusinessUnit.Size = new System.Drawing.Size(98, 28);
//            this.tslBusinessUnit.Text = "Business Unit:";
//            // 
//            // tscBusinessUnit
//            // 
//            this.tscBusinessUnit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
//            this.tscBusinessUnit.Name = "tscBusinessUnit";
//            this.tscBusinessUnit.Size = new System.Drawing.Size(180, 31);
//            // 
//            // tslTeam
//            // 
//            this.tslTeam.Name = "tslTeam";
//            this.tslTeam.Size = new System.Drawing.Size(48, 28);
//            this.tslTeam.Text = "Team:";
//            // 
//            // tscTeam
//            // 
//            this.tscTeam.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
//            this.tscTeam.Name = "tscTeam";
//            this.tscTeam.Size = new System.Drawing.Size(180, 31);
//            // 
//            // tslAssignment
//            // 
//            this.tslAssignment.Name = "tslAssignment";
//            this.tslAssignment.Size = new System.Drawing.Size(89, 28);
//            this.tslAssignment.Text = "Assignment:";
//            // 
//            // tscAssignment
//            // 
//            this.tscAssignment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
//            this.tscAssignment.Name = "tscAssignment";
//            this.tscAssignment.Size = new System.Drawing.Size(95, 31);
//            // 
//            // toolStripSeparator2
//            // 
//            this.toolStripSeparator2.Name = "toolStripSeparator2";
//            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 31);
//            // 
//            // tslSearch
//            // 
//            this.tslSearch.Name = "tslSearch";
//            this.tslSearch.Size = new System.Drawing.Size(56, 28);
//            this.tslSearch.Text = "Search:";
//            // 
//            // tstSearch
//            // 
//            this.tstSearch.Font = new System.Drawing.Font("Segoe UI", 9F);
//            this.tstSearch.Name = "tstSearch";
//            this.tstSearch.Size = new System.Drawing.Size(220, 27);
//            // 
//            // tslCount
//            // 
//            this.tslCount.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
//            this.tslCount.Name = "tslCount";
//            this.tslCount.Size = new System.Drawing.Size(59, 20);
//            this.tslCount.Text = "Rows: 0";
//            // 
//            // dgvResults
//            // 
//            this.dgvResults.ColumnHeadersHeight = 29;
//            this.dgvResults.Dock = System.Windows.Forms.DockStyle.Fill;
//            this.dgvResults.Location = new System.Drawing.Point(0, 31);
//            this.dgvResults.Name = "dgvResults";
//            this.dgvResults.RowHeadersWidth = 51;
//            this.dgvResults.Size = new System.Drawing.Size(1100, 669);
//            this.dgvResults.TabIndex = 1;
//            // 
//            // MyPluginControl
//            // 
//            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
//            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
//            this.Controls.Add(this.dgvResults);
//            this.Controls.Add(this.tsMain);
//            this.Name = "MyPluginControl";
//            this.Size = new System.Drawing.Size(1100, 700);
//            this.tsMain.ResumeLayout(false);
//            this.tsMain.PerformLayout();
//            ((System.ComponentModel.ISupportInitialize)(this.dgvResults)).EndInit();
//            this.ResumeLayout(false);
//            this.PerformLayout();

//        }
//    }
//}

////namespace GM.UserRolesMatrix
////{
////    partial class MyPluginControl
////    {
////        /// <summary> 
////        /// Variable nécessaire au concepteur.
////        /// </summary>
////        private System.ComponentModel.IContainer components = null;

////        /// <summary> 
////        /// Nettoyage des ressources utilisées.
////        /// </summary>
////        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
////        protected override void Dispose(bool disposing)
////        {
////            if (disposing && (components != null))
////            {
////                components.Dispose();
////            }
////            base.Dispose(disposing);
////        }

////        #region Code généré par le Concepteur de composants

////        /// <summary> 
////        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas 
////        /// le contenu de cette méthode avec l'éditeur de code.
////        /// </summary>
////        private void InitializeComponent()
////        {
////            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MyPluginControl));
////            this.toolStripMenu = new System.Windows.Forms.ToolStrip();
////            this.tsbClose = new System.Windows.Forms.ToolStripButton();
////            this.tsbSample = new System.Windows.Forms.ToolStripButton();
////            this.tssSeparator1 = new System.Windows.Forms.ToolStripSeparator();
////            this.toolStripMenu.SuspendLayout();
////            this.SuspendLayout();
////            // 
////            // toolStripMenu
////            // 
////            this.toolStripMenu.ImageScalingSize = new System.Drawing.Size(24, 24);
////            this.toolStripMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
////                this.tsbClose,
////                this.tssSeparator1,
////                this.tsbSample});
////            this.toolStripMenu.Location = new System.Drawing.Point(0, 0);
////            this.toolStripMenu.Name = "toolStripMenu";
////            this.toolStripMenu.Padding = new System.Windows.Forms.Padding(0, 0, 2, 0);
////            this.toolStripMenu.Size = new System.Drawing.Size(839, 31);
////            this.toolStripMenu.TabIndex = 4;
////            this.toolStripMenu.Text = "toolStrip1";
////            // 
////            // tsbClose
////            // 
////            this.tsbClose.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
////            this.tsbClose.Name = "tsbClose";
////            this.tsbClose.Size = new System.Drawing.Size(28, 28);
////            this.tsbClose.Text = "Close this tool";
////            this.tsbClose.Click += new System.EventHandler(this.tsbClose_Click);
////            // 
////            // tssSeparator1
////            // 
////            this.tssSeparator1.Name = "tssSeparator1";
////            this.tssSeparator1.Size = new System.Drawing.Size(6, 31);
////            // 
////            // tsbSample
////            // 
////            this.tsbSample.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
////            this.tsbSample.Name = "tsbSample";
////            this.tsbSample.Size = new System.Drawing.Size(28, 28);
////            this.tsbSample.Text = "Try me";
////            this.tsbSample.Click += new System.EventHandler(this.tsbSample_Click);
////            // 
////            // MyPluginControl
////            // 
////            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
////            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
////            this.Controls.Add(this.toolStripMenu);
////            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
////            this.Name = "SampleTool";
////            this.Size = new System.Drawing.Size(839, 462);
////            this.Load += new System.EventHandler(this.MyPluginControl_Load);
////            this.toolStripMenu.ResumeLayout(false);
////            this.toolStripMenu.PerformLayout();
////            this.ResumeLayout(false);
////            this.PerformLayout();

////        }

////        #endregion
////        private System.Windows.Forms.ToolStrip toolStripMenu;
////        private System.Windows.Forms.ToolStripButton tsbClose;
////        private System.Windows.Forms.ToolStripButton tsbSample;
////        private System.Windows.Forms.ToolStripSeparator tssSeparator1;
////    }
//}