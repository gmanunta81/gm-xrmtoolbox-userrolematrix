using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Xml.Linq;
using XrmToolBox.Extensibility;
using Button = System.Windows.Forms.Button;
using ComboBox = System.Windows.Forms.ComboBox;
using Image = System.Drawing.Image;
using Label = System.Windows.Forms.Label;
using ToolTip = System.Windows.Forms.ToolTip;

namespace GM.XrmToolBox.UserRoleMatrix
{
    public partial class MyPluginControl : PluginControlBase
    {
        // -----------------------
        // MODE (global hidden variable)
        // -----------------------
        private const string ModeNone = "";
        private const string ModeSystemUser = "systemuser";
        private const string ModeTeam = "team";
        private string _currentMode = ModeNone;
        private const string AboutIconBase64 =
  "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAADf0lEQVR42u2Xa0hTYRjH/2eXc3brzNXUuUJDaxOnrVgXukmEZZkfihYliRRBYWQWRCn1qaB1nUVhsT5UUISaIFnR5UPQfUZmMaNsoEsj03ltc3q2efrQWrNpTmquDz1wvpyH9/3/zvM+z3OeF4iwESN6iizsX1U6kkqEBlBkYVmDJiEsX1tcb/sVhDecOFFcbwsHAGvQJBCwsIEQnPES/xEB1qBJCDxePwnLsvHhFA+KBEF8/BmBv51woZhPkxPpMvwPEHEA3kgOikcgf95ErEujoYmlICY5cDCDaHd48d4+gJxrLXAwg7izOQGZKgkAoMHOINn4AawvpUkugaZ9KsRN+C5zta4HuWUto0dALubCvD0RJdkK8Fwd3cv1m0olsuhNM1NU+cUFW05L7PUN1O19RjB9/YHrVHISWY3Gc/AhbNBK/eIAgPqbD1FdVDJqBExrlNDGCeDsczHZ6brCdp4iFmvO7rfFqKfaHPbuSuNzK9pau+Bl3AAEAMAwjIckSd6uLTkLbpke1SIpXbdr4SQE+kI6glgJD6tTaABARXnZ4/ZuB4P80gMQyb6/lCqjIVVGQ71sfuC6urq6Rqk0SpyRkaHVHL9YGpO4UjdLKYC59s3HKDHFVavVk0NKwpQYCoSvP1qt1s9IzlwAkYy+sn4KWIPG/+xZLB/a3ViWPX2p3AwAhfol03fPFbgAoOTY4QrwKDLkKiCIIZsCsvg4AMgta4FQt/aM3/ngxGWYL1UFrr1cUf2yy+Fy5+XlLVmVFito/tzWW1lZ+QwCWhIywNu2AX8WJyUlKYYQZR0q+F1J9bn63Rde9HAoiuJzOARx5tTJKo9ILgMpFoYM0PrVg2pLhxsA9Hr9/ChPR9dY6vqsuZfr8Q6yTqez/4LJdA+6javG3Ii23rDz31ptdpqmRVWGbYu0E70ukksgTSEYFaC5xw3+zgcdElqa2+0c8EKrXzbmRvTF4cGc8838Amn5jXVZSzWPt0/jCoRCtrd/kGho/eq2PL3/sqam5gO4c2cPuyutkGPv6+t/1An7+DLp0Z6lmUf3XL2Fd3fPobPpEzwDbpASIUQyGtHT4jE7dfqKizagbNtBND55BeUM1XB7JRutgCl7BzqbPkGTnR48E47DNBQ0kPjmw/9/w38EYIRLQ1jNp8kJGpnHKwGDjuBIKhFuiMDs/2euZhG/nEbcvgEhmWjf6ekl5gAAAABJRU5ErkJggg==";


        // -----------------------
        // DATA
        // -----------------------
        private DataTable _table;
        private DataView _view;
        private readonly BindingSource _bindingSource = new BindingSource();
        private bool _updatingFilters;
        private List<string> _allBusinessUnitNames = new List<string>();


        // Org setting status
        private OrgSettingStatus _orgSettingStatus = OrgSettingStatus.Unknown();

        // Export columns (do NOT export technical columns or Selected)
        private static readonly string[] ExportColumns = new[]
        {
            "User",
            "Email",
            "User Business Unit",
            "Assignment Type",
            "Team",
            "Is Default Team",
            "Team Business Unit",
            "Role",
            "Role Business Unit",
            "Duplicated Role in same BU"
        };

        public MyPluginControl()
        {
            DependencyResolver.Register();
            InitializeComponent();
            tsbAbout.Image = ImageFromBase64(AboutIconBase64);
            tsbAbout.Click += (s, e) => ShowAboutDialog();

            // Grid config
            dgvResults.AutoGenerateColumns = true;
            dgvResults.AllowUserToAddRows = false;
            dgvResults.AllowUserToDeleteRows = false;
            dgvResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvResults.MultiSelect = true;
            dgvResults.EditMode = DataGridViewEditMode.EditOnEnter;
            dgvResults.DataSource = _bindingSource;

            // Toolstrip events
            tsbLoadUsersRoles.Click += (s, e) => ExecuteMethod(LoadUsersAndRoles);
            tsbLoadOwnerTeamsRoles.Click += (s, e) => ExecuteMethod(LoadOwnerTeamsRoles);
            tsbDel.Click += (s, e) => ExecuteMethod(DeleteSelectedAssignments);
            tsbAddUserRole.Click += (s, e) => ExecuteMethod(OpenAddUserRoleDialog);

            tsmiExportCsv.Click += (s, e) => ExportCsv();
            tsmiExportExcel.Click += (s, e) => ExportExcel();

            // Filters/search
            tstSearch.TextChanged += (s, e) => ApplyAllFilters();
            tscBusinessUnit.SelectedIndexChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };
            tscTeam.SelectedIndexChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };
            tscAssignment.SelectedIndexChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };
            if (tscIsDefaultTeam != null)
                tscIsDefaultTeam.SelectedIndexChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };

            dgvResults.DataBindingComplete += (s, e) =>
            {
                ConfigureGridColumns();
                ApplyDuplicateRowHighlight();
            };

            InitializeStaticFilters();
            UpdateCountLabel(0, 0);
            UpdateActionButtonsEnabledState();
        }
        private void ShowAboutDialog()
        {
            using (var f = new AboutForm())
            {
                f.StartPosition = FormStartPosition.CenterParent;
                f.ShowDialog(this);
            }
        }

        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);

            ClearResults();
            RefreshOrgSettingStatus(); // async in background
        }

        // -----------------------
        // UI helpers
        // -----------------------
        private void InitializeStaticFilters()
        {
            _updatingFilters = true;
            try
            {
                tscAssignment.Items.Clear();
                tscAssignment.Items.Add("All");
                tscAssignment.Items.Add("Direct");
                tscAssignment.Items.Add("Team");
                tscAssignment.SelectedIndex = 0;

                tscBusinessUnit.Items.Clear();
                tscBusinessUnit.Items.Add("All");
                tscBusinessUnit.SelectedIndex = 0;

                tscTeam.Items.Clear();
                tscTeam.Items.Add("All");
                tscTeam.SelectedIndex = 0;

                if (tscIsDefaultTeam != null)
                {
                    tscIsDefaultTeam.Items.Clear();
                    tscIsDefaultTeam.Items.Add("All");
                    tscIsDefaultTeam.Items.Add("Yes");
                    tscIsDefaultTeam.Items.Add("No");
                    tscIsDefaultTeam.SelectedIndex = 0;
                }
            }
            finally
            {
                _updatingFilters = false;
            }
        }
        private static Image ImageFromBase64(string base64)
        {
            var bytes = Convert.FromBase64String(base64);
            using (var ms = new MemoryStream(bytes))
            using (var img = Image.FromStream(ms))
            {
                return new Bitmap(img);
            }
        }

        private void ClearResults()
        {
            _table = null;
            _view = null;
            _bindingSource.DataSource = null;

            _currentMode = ModeNone;

            InitializeStaticFilters();
            UpdateCountLabel(0, 0);
            UpdateActionButtonsEnabledState();
        }

        private void UpdateActionButtonsEnabledState()
        {
            var hasData = _table != null && _table.Rows.Count > 0;

            tsbDel.Enabled = hasData && (_currentMode == ModeSystemUser || _currentMode == ModeTeam);
            tsbAddUserRole.Enabled = true; // always available, it reloads user view after add
        }

        private void UpdateCountLabel(int filtered, int total)
        {
            var recomputeText = _orgSettingStatus.IsKnown
                ? (_orgSettingStatus.EnableOwnershipAcrossBusinessUnits ? "Active" : "Inactive")
                : "Unknown";

            // Required feature #1: status near record count
            tslCount.Text = $"Rows: {filtered:n0} / {total:n0} | EnableOwnershipAccrossBUs: {recomputeText}";
        }

        // -----------------------
        // Feature #1: Read Org Setting
        // -----------------------
        private void RefreshOrgSettingStatus()
        {
            if (Service == null) return;

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Reading organization DB settings...",
                Work = (w, e) =>
                {
                    e.Result = TryReadOrgSettingStatus(Service);
                },
                PostWorkCallBack = e =>
                {
                    if (e.Error != null)
                    {
                        // non-blocking: keep Unknown
                        _orgSettingStatus = OrgSettingStatus.Unknown();
                    }
                    else
                    {
                        _orgSettingStatus = (OrgSettingStatus)e.Result;
                    }

                    // refresh count label with current view counts
                    var filtered = _view?.Count ?? 0;
                    var total = _table?.Rows.Count ?? 0;
                    UpdateCountLabel(filtered, total);
                }
            });
        }

        private static OrgSettingStatus TryReadOrgSettingStatus(IOrganizationService svc)
        {
            // Under the hood this is stored in organization.orgdborgsettings as XML. :contentReference[oaicite:1]{index=1}
            try
            {
                var qe = new QueryExpression("organization")
                {
                    ColumnSet = new ColumnSet("orgdborgsettings"),
                    NoLock = true
                };

                var ec = RetrieveAll(svc, qe, 5000);
                var org = ec.Entities.FirstOrDefault();
                var xml = org?.GetAttributeValue<string>("orgdborgsettings");

                if (string.IsNullOrWhiteSpace(xml))
                    return OrgSettingStatus.Known(false);

                // Parse XML and get the RecomputeOwnershipAcrossBusinessUnits element
                var recompute = TryParseOrgSettingBool(xml, "EnableOwnershipAcrossBusinessUnits");

                // If element not found, treat as false but still "Known"
                return OrgSettingStatus.Known(recompute ?? false);
            }
            catch
            {
                return OrgSettingStatus.Unknown();
            }
        }

        private static bool? TryParseOrgSettingBool(string xml, string elementName)
        {
            try
            {
                var doc = XDocument.Parse(xml);
                var node = doc.Descendants()
                    .FirstOrDefault(x => string.Equals(x.Name.LocalName, elementName, StringComparison.OrdinalIgnoreCase));

                if (node == null) return null;

                var raw = (node.Value ?? "").Trim();
                if (bool.TryParse(raw, out var b)) return b;

                // sometimes can be 0/1
                if (raw == "1") return true;
                if (raw == "0") return false;

                return null;
            }
            catch
            {
                return null;
            }
        }

        private readonly struct OrgSettingStatus
        {
            public bool IsKnown { get; }
            public bool EnableOwnershipAcrossBusinessUnits { get; }

            private OrgSettingStatus(bool isKnown, bool recompute)
            {
                IsKnown = isKnown;
                EnableOwnershipAcrossBusinessUnits = recompute;
            }

            public static OrgSettingStatus Unknown() => new OrgSettingStatus(false, false);
            public static OrgSettingStatus Known(bool recompute) => new OrgSettingStatus(true, recompute);
        }

        // -----------------------
        // Feature #2: Load Users & Roles (systemuser mode)
        // -----------------------
        private void LoadUsersAndRoles()
        {
            tsbLoadUsersRoles.Enabled = false;
            tsbLoadOwnerTeamsRoles.Enabled = false;

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading users and roles...",
                Work = (worker, args) =>
                {
                    worker.ReportProgress(0, "Retrieving users...");
                    var users = RetrieveAllUsers(Service);

                    worker.ReportProgress(0, "Retrieving direct user roles...");
                    var directRoles = RetrieveDirectUserRoles(Service);

                    worker.ReportProgress(0, "Retrieving roles via Owner Teams...");
                    var teamRoles = RetrieveTeamUserRoles(Service);

                    worker.ReportProgress(0, "Building rows (duplicates)...");
                    var rows = BuildUserRows(users, directRoles, teamRoles);

                    worker.ReportProgress(0, "Retrieving all Business Units...");
                    var allBus = RetrieveAllBusinessUnitNames(Service);


                    args.Result = new LoadResult
                    {
                        Rows = rows,
                        BusinessUnits = allBus
                    };

                },
                ProgressChanged = e =>
                {
                    if (e.UserState != null)
                        SetWorkingMessage(e.UserState.ToString());
                },
                PostWorkCallBack = e =>
                {
                    tsbLoadUsersRoles.Enabled = true;
                    tsbLoadOwnerTeamsRoles.Enabled = true;

                    if (e.Error != null)
                    {
                        ShowErrorDialog(e.Error);
                        return;
                    }

                    _currentMode = ModeSystemUser;

                    var result = (LoadResult)e.Result;
                    // Persist business unit names for filters
                    _allBusinessUnitNames = result.BusinessUnits ?? new List<string>();

                    BindRows(result.Rows ?? new List<MatrixRow>());
                    PopulateDropdownFiltersFromTable();
                    ApplyAllFilters();

                    UpdateActionButtonsEnabledState();
                }
            });
        }

        // -----------------------
        // Feature #2: Load Owner Teams Roles (team mode)
        // -----------------------
        private void LoadOwnerTeamsRoles()
        {
            tsbLoadUsersRoles.Enabled = false;
            tsbLoadOwnerTeamsRoles.Enabled = false;

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading Owner Teams and their roles...",
                Work = (worker, args) =>
                {
                    worker.ReportProgress(0, "Retrieving Owner Teams...");
                    var teams = RetrieveOwnerTeams(Service);

                    worker.ReportProgress(0, "Retrieving Owner Team roles...");
                    var teamRoles = RetrieveOwnerTeamRoles(Service);

                    worker.ReportProgress(0, "Building rows...");
                    var rows = BuildOwnerTeamRows(teams, teamRoles);

                    worker.ReportProgress(0, "Retrieving all Business Units...");
                    var allBus = RetrieveAllBusinessUnitNames(Service);

                    args.Result = new LoadResult
                    {
                        Rows = rows,
                        BusinessUnits = allBus
                    };

                },
                ProgressChanged = e =>
                {
                    if (e.UserState != null)
                        SetWorkingMessage(e.UserState.ToString());
                },
                PostWorkCallBack = e =>
                {
                    tsbLoadUsersRoles.Enabled = true;
                    tsbLoadOwnerTeamsRoles.Enabled = true;

                    if (e.Error != null)
                    {
                        ShowErrorDialog(e.Error);
                        return;
                    }

                    _currentMode = ModeTeam;

                    var result = (LoadResult)e.Result;
                    _allBusinessUnitNames = result.BusinessUnits ?? new List<string>();

                    BindRows(result.Rows);
                    PopulateDropdownFiltersFromTable();
                    ApplyAllFilters();

                    UpdateActionButtonsEnabledState();
                }
            });
        }

        // -----------------------
        // Feature #3: Delete Selected (DisassociateRequest)
        // -----------------------
        private void DeleteSelectedAssignments()
        {
            if (_table == null || _table.Rows.Count == 0 || dgvResults.Rows.Count == 0)
            {
                MessageBox.Show("No records selected/found.", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (_currentMode != ModeSystemUser && _currentMode != ModeTeam)
            {
                MessageBox.Show("No records selected/found.", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Commit checkbox edits
            dgvResults.EndEdit();
            // Ensure the in-progress cell edit is committed to the underlying data source
            dgvResults.CommitEdit(DataGridViewDataErrorContexts.Commit);
            _bindingSource.EndEdit();

            // Gather selected rows (visible)
            var ops = new List<DeleteOp>();
            int skippedTeamAssignmentsInUserMode = 0;
            int skippedMissingIds = 0;

            foreach (DataGridViewRow gridRow in dgvResults.Rows)
            {
                if (!(gridRow.DataBoundItem is DataRowView drv)) continue;

                var row = drv.Row;

                var selected = row.Field<bool?>("Selected") ?? false;
                if (!selected) continue;

                var userId = row.Field<Guid?>("UserId") ?? Guid.Empty;
                var teamId = row.Field<Guid?>("TeamId") ?? Guid.Empty;
                var roleId = row.Field<Guid?>("RoleId") ?? Guid.Empty;
                var assignmentType = (row.Field<string>("Assignment Type") ?? "").Trim();

                if (_currentMode == ModeSystemUser)
                {
                    // Safety: In systemuser mode we only remove DIRECT user-role assignments.
                    // Team assignments are skipped to avoid accidentally impacting a whole team.
                    if (!string.Equals(assignmentType, "Direct", StringComparison.OrdinalIgnoreCase))
                    {
                        skippedTeamAssignmentsInUserMode++;
                        continue;
                    }

                    if (userId == Guid.Empty || roleId == Guid.Empty)
                    {
                        skippedMissingIds++;
                        continue;
                    }

                    ops.Add(DeleteOp.ForUserRole(userId, roleId));
                }
                else if (_currentMode == ModeTeam)
                {
                    if (teamId == Guid.Empty || roleId == Guid.Empty)
                    {
                        skippedMissingIds++;
                        continue;
                    }

                    ops.Add(DeleteOp.ForTeamRole(teamId, roleId));
                }
            }

            // Deduplicate operations
            ops = ops
                .GroupBy(o => o.Key)
                .Select(g => g.First())
                .ToList();

            if (ops.Count == 0)
            {
                MessageBox.Show("No records selected/found.", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            tsbDel.Enabled = false;

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Deleting role assignments...",
                Work = (w, e) =>
                {
                    int success = 0;
                    var errors = new List<string>();

                    foreach (var op in ops)
                    {
                        try
                        {
                            var req = new DisassociateRequest
                            {
                                Target = op.Target,
                                Relationship = new Relationship(op.RelationshipName),
                                RelatedEntities = new EntityReferenceCollection
                                {
                                    op.Related
                                }
                            };

                            Service.Execute(req);
                            success++;
                        }
                        catch (Exception ex)
                        {
                            errors.Add($"{op.Key}: {ex.Message}");
                        }
                    }

                    e.Result = new DeleteResult
                    {
                        Total = ops.Count,
                        Success = success,
                        Errors = errors,
                        SkippedTeamAssignmentsInUserMode = skippedTeamAssignmentsInUserMode,
                        SkippedMissingIds = skippedMissingIds
                    };
                },
                PostWorkCallBack = e =>
                {
                    tsbDel.Enabled = true;

                    if (e.Error != null)
                    {
                        ShowErrorDialog(e.Error);
                        return;
                    }

                    var res = (DeleteResult)e.Result;

                    var sb = new StringBuilder();
                    sb.AppendLine($"Deleted: {res.Success}/{res.Total}");

                    if (res.SkippedTeamAssignmentsInUserMode > 0)
                        sb.AppendLine($"Skipped (Team assignments in user mode): {res.SkippedTeamAssignmentsInUserMode}");

                    if (res.SkippedMissingIds > 0)
                        sb.AppendLine($"Skipped (Missing IDs): {res.SkippedMissingIds}");

                    if (res.Errors.Count > 0)
                    {
                        sb.AppendLine();
                        sb.AppendLine("Errors:");
                        foreach (var err in res.Errors.Take(10))
                            sb.AppendLine("- " + err);

                        if (res.Errors.Count > 10)
                            sb.AppendLine($"... ({res.Errors.Count - 10} more)");
                    }

                    MessageBox.Show(sb.ToString(), "Delete", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Reload current view
                    if (_currentMode == ModeSystemUser)
                        LoadUsersAndRoles();
                    else if (_currentMode == ModeTeam)
                        LoadOwnerTeamsRoles();
                }
            });
        }

        private sealed class DeleteOp
        {
            public string RelationshipName { get; private set; }
            public EntityReference Target { get; private set; }
            public EntityReference Related { get; private set; }
            public string Key { get; private set; }

            public static DeleteOp ForUserRole(Guid userId, Guid roleId)
            {
                return new DeleteOp
                {
                    RelationshipName = "systemuserroles_association",
                    Target = new EntityReference("systemuser", userId),
                    Related = new EntityReference("role", roleId),
                    Key = $"systemuser:{userId:N}|role:{roleId:N}"
                };
            }

            public static DeleteOp ForTeamRole(Guid teamId, Guid roleId)
            {
                return new DeleteOp
                {
                    RelationshipName = "teamroles_association",
                    Target = new EntityReference("team", teamId),
                    Related = new EntityReference("role", roleId),
                    Key = $"team:{teamId:N}|role:{roleId:N}"
                };
            }
        }

        private sealed class DeleteResult
        {
            public int Total { get; set; }
            public int Success { get; set; }
            public List<string> Errors { get; set; } = new List<string>();

            public int SkippedTeamAssignmentsInUserMode { get; set; }
            public int SkippedMissingIds { get; set; }
        }

        // -----------------------
        // Feature #4: Add User Role (popup + AssociateRequest)
        // -----------------------
        private void OpenAddUserRoleDialog()
        {
            if (Service == null)
            {
                MessageBox.Show("Not connected.", "Add User Role", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Ensure org setting is known (best effort)
            if (!_orgSettingStatus.IsKnown)
            {
                _orgSettingStatus = TryReadOrgSettingStatus(Service);
                var filtered = _view?.Count ?? 0;
                var total = _table?.Rows.Count ?? 0;
                UpdateCountLabel(filtered, total);
            }

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading users and business units...",
                Work = (w, e) =>
                {
                    var users = RetrieveAllUsers(Service).Values
                        .Where(u => !u.IsDisabled)
                        .OrderBy(u => u.FullName, StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    var bus = RetrieveAllBusinessUnits(Service);
                    var root = bus.FirstOrDefault(b => b.ParentBusinessUnitId == Guid.Empty) ?? bus.FirstOrDefault();

                    e.Result = new AddDialogData
                    {
                        Users = users,
                        BusinessUnits = bus,
                        RootBusinessUnit = root
                    };
                },
                PostWorkCallBack = e =>
                {
                    if (e.Error != null)
                    {
                        ShowErrorDialog(e.Error);
                        return;
                    }

                    var data = (AddDialogData)e.Result;

                    using (var dlg = new AddUserRoleForm(
                        service: Service,
                        enableAcrossBuEnabled: _orgSettingStatus.EnableOwnershipAcrossBusinessUnits,
                        users: data.Users,
                        businessUnits: data.BusinessUnits,
                        rootBusinessUnit: data.RootBusinessUnit))
                    {
                        var result = dlg.ShowDialog(this);
                        if (result == DialogResult.OK)
                        {
                            // Reload user view
                            LoadUsersAndRoles();
                        }
                    }
                }
            });
        }

        private sealed class AddDialogData
        {
            public List<UserInfo> Users { get; set; }
            public List<BusinessUnitInfo> BusinessUnits { get; set; }
            public BusinessUnitInfo RootBusinessUnit { get; set; }
        }

        private sealed class AddUserRoleForm : Form
        {
            private readonly IOrganizationService _service;
            private readonly bool _enableAcrossBuEnabled;
            private readonly List<UserInfo> _users;
            private readonly List<BusinessUnitInfo> _businessUnits;
            private readonly BusinessUnitInfo _rootBusinessUnit;

            private readonly ComboBox _cbUser = new ComboBox();
            private readonly ComboBox _cbBusinessUnit = new ComboBox();
            private readonly ComboBox _cbRole = new ComboBox();
            private readonly Button _btnAdd = new Button();
            private readonly Button _btnCancel = new Button();
            private readonly ToolTip _toolTip = new ToolTip();

            public AddUserRoleForm(
                IOrganizationService service,
                bool enableAcrossBuEnabled,
                List<UserInfo> users,
                List<BusinessUnitInfo> businessUnits,
                BusinessUnitInfo rootBusinessUnit)
            {
                _service = service ?? throw new ArgumentNullException(nameof(service));
                _enableAcrossBuEnabled = enableAcrossBuEnabled;
                _users = users ?? new List<UserInfo>();
                _businessUnits = businessUnits ?? new List<BusinessUnitInfo>();
                _rootBusinessUnit = rootBusinessUnit;

                Text = "Add User Role";
                StartPosition = FormStartPosition.CenterParent;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                ShowInTaskbar = false;
                Width = 560;
                Height = 260;

                BuildUi();
                LoadInitialData();

              

            }

            private void BuildUi()
            {
                var lblUser = new Label { Text = "User:", Left = 16, Top = 20, Width = 120 };
                var lblBu = new Label { Text = "Business Unit:", Left = 16, Top = 70, Width = 120 };
                var lblRole = new Label { Text = "Security Role:", Left = 16, Top = 120, Width = 120 };

                ConfigureCombo(_cbUser);
                _cbUser.Left = 150;
                _cbUser.Top = 16;
                _cbUser.Width = 370;

                ConfigureCombo(_cbBusinessUnit);
                _cbBusinessUnit.Left = 150;
                _cbBusinessUnit.Top = 66;
                _cbBusinessUnit.Width = 370;
                _cbBusinessUnit.SelectedIndexChanged += async (s, e) => await LoadRolesForSelectedBuAsync();

                ConfigureCombo(_cbRole);
                _cbRole.Left = 150;
                _cbRole.Top = 116;
                _cbRole.Width = 370;

                _btnAdd.Text = "Add";
                _btnAdd.Left = 330;
                _btnAdd.Top = 170;
                _btnAdd.Width = 90;
                _btnAdd.Click += async (s, e) => await AddClickedAsync();

                _btnCancel.Text = "Cancel";
                _btnCancel.Left = 430;
                _btnCancel.Top = 170;
                _btnCancel.Width = 90;
                _btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

                Controls.Add(lblUser);
                Controls.Add(_cbUser);
                Controls.Add(lblBu);
                Controls.Add(_cbBusinessUnit);
                Controls.Add(lblRole);
                Controls.Add(_cbRole);
                Controls.Add(_btnAdd);
                Controls.Add(_btnCancel);
            }

            private static void ConfigureCombo(ComboBox cb)
            {
                cb.DropDownStyle = ComboBoxStyle.DropDown;
                cb.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                cb.AutoCompleteSource = AutoCompleteSource.ListItems;
            }

            private void LoadInitialData()
            {
                // Users
                _cbUser.Items.Clear();
                foreach (var u in _users)
                {
                    var text = string.IsNullOrWhiteSpace(u.Email)
                        ? u.FullName
                        : $"{u.FullName} ({u.Email})";

                    _cbUser.Items.Add(new ComboItem(u.UserId, text));
                }

                // Business Units
                _cbBusinessUnit.Items.Clear();

                if (!_enableAcrossBuEnabled)
                {
                    // Requirement: BU dropdown not selectable, tooltip "dborgsetting inactive", show only root BU.
                    _cbBusinessUnit.Enabled = false;
                    _toolTip.SetToolTip(_cbBusinessUnit, "dborgsetting inactive");

                    if (_rootBusinessUnit != null)
                        _cbBusinessUnit.Items.Add(new ComboItem(_rootBusinessUnit.BusinessUnitId, _rootBusinessUnit.Name));
                }
                else
                {
                    _cbBusinessUnit.Enabled = true;

                    foreach (var bu in _businessUnits.OrderBy(b => b.Name, StringComparer.OrdinalIgnoreCase))
                        _cbBusinessUnit.Items.Add(new ComboItem(bu.BusinessUnitId, bu.Name));
                }

                // Select root BU by default and auto-load root roles
                if (_rootBusinessUnit != null)
                {
                    var rootItem = _cbBusinessUnit.Items
                        .Cast<object>()
                        .OfType<ComboItem>()
                        .FirstOrDefault(x => x.Id == _rootBusinessUnit.BusinessUnitId);

                    if (rootItem != null)
                        _cbBusinessUnit.SelectedItem = rootItem;
                }
                AttachCloseDropDownOnTyping(_cbUser);
                AttachCloseDropDownOnTyping(_cbBusinessUnit);
                AttachCloseDropDownOnTyping(_cbRole);
            }

            private async System.Threading.Tasks.Task LoadRolesForSelectedBuAsync()
            {
                var buItem = _cbBusinessUnit.SelectedItem as ComboItem;
                if (buItem == null || buItem.Id == Guid.Empty)
                {
                    _cbRole.Items.Clear();
                    return;
                }

                _cbRole.Enabled = false;
                _cbRole.Items.Clear();
                _cbRole.Items.Add("Loading...");
                _cbRole.SelectedIndex = 0;

                try
                {
                    var roles = await System.Threading.Tasks.Task.Run(() => RetrieveRolesByBusinessUnit(_service, buItem.Id));

                    if (IsDisposed) return;

                    BeginInvoke(new Action(() =>
                    {
                        _cbRole.Items.Clear();
                        foreach (var r in roles.OrderBy(r => r.Name, StringComparer.OrdinalIgnoreCase))
                            _cbRole.Items.Add(new ComboItem(r.RoleId, r.Name));

                        // If we loaded items, select the first one so the "Loading..." text is removed.
                        if (_cbRole.Items.Count > 0)
                        {
                            _cbRole.SelectedIndex = 0;
                        }
                        else
                        {
                            _cbRole.SelectedIndex = -1;
                            _cbRole.Text = string.Empty;
                        }

                        _cbRole.Enabled = true;
                    }));
                }
                catch
                {
                    if (IsDisposed) return;

                    BeginInvoke(new Action(() =>
                    {
                        _cbRole.Items.Clear();
                        _cbRole.Enabled = true;
                        _cbRole.SelectedIndex = -1;
                        _cbRole.Text = string.Empty;
                    }));
                }
            }

            private async System.Threading.Tasks.Task AddClickedAsync()
            {
                var userItem = _cbUser.SelectedItem as ComboItem;
                var buItem = _cbBusinessUnit.SelectedItem as ComboItem;
                var roleItem = _cbRole.SelectedItem as ComboItem;

                TrySelectByText(_cbUser);
                TrySelectByText(_cbBusinessUnit);
                TrySelectByText(_cbRole);

                // Required validation message
                if (userItem == null || userItem.Id == Guid.Empty ||
                    roleItem == null || roleItem.Id == Guid.Empty ||
                    (_enableAcrossBuEnabled && (buItem == null || buItem.Id == Guid.Empty)))
                {
                    MessageBox.Show("Please set properly the sec data.", "Add User Role", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var req = new AssociateRequest
                {
                    Target = new EntityReference("systemuser", userItem.Id),
                    Relationship = new Relationship("systemuserroles_association"),
                    RelatedEntities = new EntityReferenceCollection
                    {
                        new EntityReference("role", roleItem.Id)
                    }
                };

                // Show a small modal wait dialog while the operation completes.
                using (var wait = new Form()
                {
                    StartPosition = FormStartPosition.CenterParent,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    ControlBox = false,
                    Width = 300,
                    Height = 100
                })
                {
                    var lbl = new Label { Text = "Please wait...", Left = 16, Top = 16, Width = 260 };
                    var pb = new System.Windows.Forms.ProgressBar { Style = ProgressBarStyle.Marquee, Left = 16, Top = 40, Width = 260, Height = 16 };
                    wait.Controls.Add(lbl);
                    wait.Controls.Add(pb);

                    Exception exception = null;

                    // Run the associate request on a background thread and close the wait dialog when done
                    var background = System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            _service.Execute(req);
                        }
                        catch (Exception ex)
                        {
                            exception = ex;
                        }

                        try
                        {
                            if (!wait.IsDisposed)
                                wait.Invoke(new Action(() => wait.Close()));
                        }
                        catch
                        {
                            // ignore invoke/close errors
                        }
                    });

                    // Show modal dialog; it will be closed by the background task via Invoke
                    try
                    {
                        wait.ShowDialog(this);
                    }
                    catch
                    {
                        // ignore show dialog errors
                    }

                    // Ensure background finished and exception is observed
                    try { await background; } catch { }

                    if (exception == null)
                    {
                        DialogResult = DialogResult.OK;
                        Close();
                    }
                    else
                    {
                        MessageBox.Show(exception.Message, "Add User Role", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            private sealed class ComboItem
            {
                public Guid Id { get; }
                public string Text { get; }

                public ComboItem(Guid id, string text)
                {
                    Id = id;
                    Text = text ?? "";
                }

                public override string ToString() => Text;
            }

            private static void AttachCloseDropDownOnTyping(ComboBox cb)
            {
                cb.TextUpdate += (s, e) =>
                {
                    if (cb.DroppedDown)
                    {
                        cb.DroppedDown = false;
                        cb.SelectionStart = cb.Text.Length;
                        cb.SelectionLength = 0;
                    }
                };
            }

            private static void TrySelectByText(ComboBox cb)
            {
                if (cb.SelectedItem != null) return;

                var text = (cb.Text ?? "").Trim();
                if (text.Length == 0) return;

                for (int i = 0; i < cb.Items.Count; i++)
                {
                    var itemText = cb.GetItemText(cb.Items[i]);
                    if (string.Equals(itemText?.Trim(), text, StringComparison.OrdinalIgnoreCase))
                    {
                        cb.SelectedIndex = i;
                        return;
                    }
                }
            }


        }

        private sealed class AboutForm : Form
        {
            public AboutForm()
            {
                Text = "About - User Roles Matrix";
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                Width = 520;
                Height = 220;

                var lblName = new Label { Left = 16, Top = 20, Width = 470, Text = "Giovanni Manunta" };
                var linkLinkedIn = new LinkLabel { Left = 16, Top = 80, Width = 470, Text = "LinkedIn profile" };
                var linkGitHub = new LinkLabel { Left = 16, Top = 110, Width = 470, Text = "GitHub profile" };

                var btnOk = new Button { Text = "OK", Left = 400, Top = 140, Width = 80, DialogResult = DialogResult.OK };
                AcceptButton = btnOk;

                
                linkLinkedIn.LinkClicked += (s, e) => OpenUrl("https://www.linkedin.com/in/giovanni-manunta-3555868/");
                linkGitHub.LinkClicked += (s, e) => OpenUrl("https://github.com/gmanunta81");

                Controls.Add(lblName);
                Controls.Add(linkLinkedIn);
                Controls.Add(linkGitHub);
                Controls.Add(btnOk);
            }

            private static void OpenUrl(string url)
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url)
                    {
                        UseShellExecute = true
                    });
                }
                catch
                {
                    // optional: ignore
                }
            }
        }

        private sealed class LoadResult
        {
            public List<MatrixRow> Rows { get; set; }
            public List<string> BusinessUnits { get; set; }
        }


        // -----------------------
        // Table binding
        // -----------------------
        private void BindRows(List<MatrixRow> rows)
        {
            _table = CreateSchema();

            foreach (var r in rows)
            {
                var dr = _table.NewRow();

                dr["Selected"] = false;

                dr["UserId"] = r.UserId;
                dr["TeamId"] = r.TeamId;
                dr["RoleId"] = r.RoleId;

                dr["User"] = r.UserFullName ?? "";
                dr["Email"] = r.UserEmail ?? "";
                dr["User Business Unit"] = r.UserBusinessUnit ?? "";

                dr["Assignment Type"] = r.AssignmentType ?? "";

                dr["Team"] = r.TeamName ?? "";
                dr["Team Business Unit"] = r.TeamBusinessUnit ?? "";
                dr["Is Default Team"] = r.IsDefaultTeam ?? "";

                dr["Role"] = r.RoleName ?? "";
                dr["Role Business Unit"] = r.RoleBusinessUnit ?? "";

                dr["Duplicated Role via Teams"] = r.Duplicate;

                _table.Rows.Add(dr);
            }

            _view = new DataView(_table);
            _bindingSource.DataSource = _view;

            ConfigureGridColumns();
            ApplyDuplicateRowHighlight();
        }

        private static DataTable CreateSchema()
        {
            var dt = new DataTable("Matrix");

            // UI selection checkbox
            dt.Columns.Add("Selected", typeof(bool));

            // Technical (hidden)
            dt.Columns.Add("UserId", typeof(Guid));
            dt.Columns.Add("TeamId", typeof(Guid));
            dt.Columns.Add("RoleId", typeof(Guid));

            // Display
            dt.Columns.Add("User", typeof(string));
            dt.Columns.Add("Email", typeof(string));
            dt.Columns.Add("User Business Unit", typeof(string));
            dt.Columns.Add("Assignment Type", typeof(string)); // Direct / Team / None
            dt.Columns.Add("Team", typeof(string));
            dt.Columns.Add("Is Default Team", typeof(string));
            dt.Columns.Add("Team Business Unit", typeof(string));
            dt.Columns.Add("Role", typeof(string));
            dt.Columns.Add("Role Business Unit", typeof(string));
            dt.Columns.Add("Duplicated Role via Teams", typeof(bool));

            return dt;
        }

        private void ConfigureGridColumns()
        {
            if (dgvResults.Columns.Count == 0) return;

            // Hide technical columns
            HideColumn("UserId");
            HideColumn("TeamId");
            HideColumn("RoleId");

            // Enlarge Business Unit columns (avoid header wrapping)
            SetFixedColumnWidth("User Business Unit", 180);
            SetFixedColumnWidth("Team Business Unit", 180);
            SetFixedColumnWidth("Role Business Unit", 180);

            // If we are in "team" mode (Load Owner Teams Roles), hide user-related columns
            bool isTeamMode = string.Equals(_currentMode, ModeTeam, StringComparison.OrdinalIgnoreCase);

            SetColumnVisible("User", !isTeamMode);
            SetColumnVisible("Email", !isTeamMode);
            SetColumnVisible("User Business Unit", !isTeamMode);
            SetColumnVisible("Assignment Type", !isTeamMode);

            // Make all columns read-only except Selected checkbox
            foreach (DataGridViewColumn col in dgvResults.Columns)
            {
                if (string.Equals(col.Name, "Selected", StringComparison.OrdinalIgnoreCase))
                {
                    col.ReadOnly = false;
                    col.Width = 60;
                    col.DisplayIndex = 0;
                }
                else
                {
                    col.ReadOnly = true;
                }
            }
        }

        private void ApplyDuplicateRowHighlight()
        {
            if (dgvResults.Rows.Count == 0 || dgvResults.Columns["Duplicated Role via Teams"] == null)
                return;

            var normalBack = dgvResults.RowsDefaultCellStyle.BackColor;
            if (normalBack == System.Drawing.Color.Empty)
                normalBack = SystemColors.Window;

            foreach (DataGridViewRow row in dgvResults.Rows)
            {
                var value = row.Cells["Duplicated Role via Teams"]?.Value;
                var isDup = value is bool b && b;
                row.DefaultCellStyle.BackColor = isDup ? System.Drawing.Color.LightYellow : normalBack;
            }
        }

        // -----------------------
        // Filters
        // -----------------------
        private void PopulateDropdownFiltersFromTable()
        {
            if (_table == null) return;

            _updatingFilters = true;
            try
            {
                string buColumn = (_currentMode == ModeTeam) ? "Team Business Unit" : "User Business Unit";

                var businessUnits = (_allBusinessUnitNames ?? new List<string>())
                 .Where(s => !string.IsNullOrWhiteSpace(s))
                 .Distinct(StringComparer.OrdinalIgnoreCase)
                 .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                 .ToList();


                var teams = _table.AsEnumerable()
                    .Select(r => r.Field<string>("Team"))
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                tscBusinessUnit.Items.Clear();
                tscBusinessUnit.Items.Add("All");
                foreach (var bu in businessUnits) tscBusinessUnit.Items.Add(bu);
                tscBusinessUnit.SelectedIndex = 0;

                tscTeam.Items.Clear();
                tscTeam.Items.Add("All");
                foreach (var t in teams) tscTeam.Items.Add(t);
                tscTeam.SelectedIndex = 0;
            }
            finally
            {
                _updatingFilters = false;
            }
        }

        private void ApplyAllFilters()
        {
            if (_view == null) return;

            var filters = new List<string>();

            string buColumn = (_currentMode == ModeTeam) ? "Team Business Unit" : "User Business Unit";

            var bu = (tscBusinessUnit.SelectedItem?.ToString() ?? "All").Trim();
            if (!string.Equals(bu, "All", StringComparison.OrdinalIgnoreCase))
                filters.Add($"[{buColumn}] = '{EscapeRowFilterValue(bu)}'");

            var team = (tscTeam.SelectedItem?.ToString() ?? "All").Trim();
            if (!string.Equals(team, "All", StringComparison.OrdinalIgnoreCase))
                filters.Add($"[Team] = '{EscapeRowFilterValue(team)}'");

            var assignment = (tscAssignment.SelectedItem?.ToString() ?? "All").Trim();
            if (string.Equals(assignment, "Direct", StringComparison.OrdinalIgnoreCase))
                filters.Add($"[Assignment Type] = 'Direct'");
            else if (string.Equals(assignment, "Team", StringComparison.OrdinalIgnoreCase))
                filters.Add($"[Assignment Type] = 'Team'");

            // Is Default Team filter (Only applicable when the column exists)
            if (tscIsDefaultTeam != null)
            {
                var isDefault = (tscIsDefaultTeam.SelectedItem?.ToString() ?? "All").Trim();
                if (string.Equals(isDefault, "Yes", StringComparison.OrdinalIgnoreCase))
                    filters.Add("[Is Default Team] = 'Yes'");
                else if (string.Equals(isDefault, "No", StringComparison.OrdinalIgnoreCase))
                    filters.Add("[Is Default Team] = 'No'");
            }

            var search = (tstSearch.Text ?? "").Trim();
            if (!string.IsNullOrWhiteSpace(search))
            {
                var s = EscapeRowFilterValue(search);

                var searchClause =
                    $"[User] LIKE '%{s}%'" +
                    $" OR [Email] LIKE '%{s}%'" +
                    $" OR [User Business Unit] LIKE '%{s}%'" +
                    $" OR [Team] LIKE '%{s}%'" +
                    $" OR [Team Business Unit] LIKE '%{s}%'" +
                    $" OR [Role] LIKE '%{s}%'" +
                    $" OR [Is Default Team] LIKE '%{s}%'" +
                    $" OR [Role Business Unit] LIKE '%{s}%'" +
                    $" OR [Assignment Type] LIKE '%{s}%'";
                filters.Add("(" + searchClause + ")");
            }

            _view.RowFilter = string.Join(" AND ", filters);

            UpdateCountLabel(_view.Count, _table?.Rows.Count ?? 0);
            ApplyDuplicateRowHighlight();
        }

        private static string EscapeRowFilterValue(string value) => (value ?? "").Replace("'", "''");

        // -----------------------
        // EXPORT
        // -----------------------
        private DataTable GetCurrentViewForExport()
        {
            if (_view == null) return null;
            return _view.ToTable(false, ExportColumns);
        }

        private void ExportCsv()
        {
            if (_view == null)
            {
                MessageBox.Show("No data to export. Please load data first.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var sfd = new SaveFileDialog
            {
                Title = "Export to CSV",
                Filter = "CSV (*.csv)|*.csv",
                FileName = $"UserRoleMatrix_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;

                var path = sfd.FileName;
                var snapshot = GetCurrentViewForExport();

                WorkAsync(new WorkAsyncInfo
                {
                    Message = "Exporting CSV...",
                    Work = (w, e) => WriteCsv(path, snapshot),
                    PostWorkCallBack = e =>
                    {
                        if (e.Error != null)
                        {
                            ShowErrorDialog(e.Error);
                            return;
                        }

                        MessageBox.Show("CSV export completed.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                });
            }
        }

        private static void WriteCsv(string path, DataTable dt)
        {
            using (var sw = new StreamWriter(path, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true)))
            {
                var headers = dt.Columns.Cast<DataColumn>().Select(c => CsvEscape(c.ColumnName));
                sw.WriteLine(string.Join(",", headers));

                foreach (DataRow row in dt.Rows)
                {
                    var fields = dt.Columns.Cast<DataColumn>().Select(c => CsvEscape(row[c]?.ToString() ?? ""));
                    sw.WriteLine(string.Join(",", fields));
                }
            }
        }

        private static string CsvEscape(string value)
        {
            if (value == null) return "";
            var mustQuote = value.Contains(",") || value.Contains("\"") || value.Contains("\r") || value.Contains("\n");
            if (!mustQuote) return value;
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private void ExportExcel()
        {
            if (_view == null)
            {
                MessageBox.Show("No data to export. Please load data first.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var sfd = new SaveFileDialog
            {
                Title = "Export to Excel",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = $"UserRoleMatrix_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;

                var path = sfd.FileName;
                var snapshot = GetCurrentViewForExport();

                WorkAsync(new WorkAsyncInfo
                {
                    Message = "Exporting Excel...",
                    Work = (w, e) => WriteExcel(path, snapshot),
                    PostWorkCallBack = e =>
                    {
                        if (e.Error != null)
                        {
                            ShowErrorDialog(e.Error);
                            return;
                        }

                        MessageBox.Show("Excel export completed.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                });
            }
        }

        private static void WriteExcel(string path, DataTable dt)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("UserRoleMatrix");
                ws.Cell(1, 1).InsertTable(dt, "UserRoleMatrixTable", true);
                ws.Columns().AdjustToContents();
                wb.SaveAs(path);
            }
        }

        // -----------------------
        // DATA RETRIEVAL (QueryExpression only)
        // -----------------------

        private static Dictionary<Guid, UserInfo> RetrieveAllUsers(IOrganizationService svc)
        {
            var qe = new QueryExpression("systemuser")
            {
                ColumnSet = new ColumnSet("systemuserid", "fullname", "internalemailaddress", "isdisabled", "businessunitid"),
                NoLock = true
            };

            var buLink = qe.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            buLink.EntityAlias = "bu";
            buLink.Columns = new ColumnSet("name");

            var ec = RetrieveAll(svc, qe);

            var users = new Dictionary<Guid, UserInfo>();
            foreach (var e in ec.Entities)
            {
                users[e.Id] = new UserInfo
                {
                    UserId = e.Id,
                    FullName = e.GetAttributeValue<string>("fullname"),
                    Email = e.GetAttributeValue<string>("internalemailaddress"),
                    BusinessUnitName = GetAliasedString(e, "bu", "name"),
                    IsDisabled = e.GetAttributeValue<bool?>("isdisabled") ?? false
                };
            }

            return users;
        }

        private static Dictionary<Guid, List<RoleInfo>> RetrieveDirectUserRoles(IOrganizationService svc)
        {
            var qe = new QueryExpression("systemuserroles")
            {
                ColumnSet = new ColumnSet("systemuserid", "roleid"),
                NoLock = true
            };

            var roleLink = qe.AddLink("role", "roleid", "roleid", JoinOperator.Inner);
            roleLink.EntityAlias = "role";
            roleLink.Columns = new ColumnSet("roleid", "name", "businessunitid");

            var roleBuLink = roleLink.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            roleBuLink.EntityAlias = "rolebu";
            roleBuLink.Columns = new ColumnSet("name");

            var ec = RetrieveAll(svc, qe);

            var map = new Dictionary<Guid, List<RoleInfo>>();
            var dedupe = new Dictionary<Guid, HashSet<Guid>>();

            foreach (var e in ec.Entities)
            {
                var userId = GetIdFromAttribute(e, "systemuserid");
                var roleId = GetIdFromAttribute(e, "roleid");
                if (userId == Guid.Empty || roleId == Guid.Empty) continue;

                if (!map.TryGetValue(userId, out var list))
                {
                    list = new List<RoleInfo>();
                    map[userId] = list;
                    dedupe[userId] = new HashSet<Guid>();
                }

                if (dedupe[userId].Add(roleId))
                {
                    list.Add(new RoleInfo
                    {
                        RoleId = roleId,
                        Name = GetAliasedString(e, "role", "name"),
                        BusinessUnitName = GetAliasedString(e, "rolebu", "name")
                    });
                }
            }

            return map;
        }

        private static Dictionary<Guid, List<TeamRoleInfo>> RetrieveTeamUserRoles(IOrganizationService svc)
        {
            var qe = new QueryExpression("teammembership")
            {
                ColumnSet = new ColumnSet("systemuserid", "teamid"),
                NoLock = true
            };

            var teamLink = qe.AddLink("team", "teamid", "teamid", JoinOperator.Inner);
            teamLink.EntityAlias = "team";
            // include isdefault so we can surface it in team-role rows
            teamLink.Columns = new ColumnSet("teamid", "name", "teamtype", "businessunitid", "isdefault");
            teamLink.LinkCriteria.AddCondition("teamtype", ConditionOperator.Equal, 0); // Owner team

            var teamBuLink = teamLink.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            teamBuLink.EntityAlias = "teambu";
            teamBuLink.Columns = new ColumnSet("name");

            var teamRolesLink = teamLink.AddLink("teamroles", "teamid", "teamid", JoinOperator.Inner);

            var roleLink = teamRolesLink.AddLink("role", "roleid", "roleid", JoinOperator.Inner);
            roleLink.EntityAlias = "role";
            roleLink.Columns = new ColumnSet("roleid", "name", "businessunitid");

            var roleBuLink = roleLink.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            roleBuLink.EntityAlias = "rolebu";
            roleBuLink.Columns = new ColumnSet("name");

            var ec = RetrieveAll(svc, qe);

            var map = new Dictionary<Guid, List<TeamRoleInfo>>();
            var dedupe = new Dictionary<Guid, HashSet<string>>();

            foreach (var e in ec.Entities)
            {
                var userId = GetIdFromAttribute(e, "systemuserid");
                var teamId = GetIdFromAttribute(e, "teamid");
                var roleId = GetAliasedGuid(e, "role", "roleid");

                if (userId == Guid.Empty || teamId == Guid.Empty || roleId == Guid.Empty) continue;

                var key = $"{teamId:N}|{roleId:N}";

                if (!map.TryGetValue(userId, out var list))
                {
                    list = new List<TeamRoleInfo>();
                    map[userId] = list;
                    dedupe[userId] = new HashSet<string>();
                }

                if (dedupe[userId].Add(key))
                {
                    list.Add(new TeamRoleInfo
                    {
                        TeamId = teamId,
                        TeamName = GetAliasedString(e, "team", "name"),
                        TeamBusinessUnitName = GetAliasedString(e, "teambu", "name"),
                        RoleId = roleId,
                        RoleName = GetAliasedString(e, "role", "name"),
                        RoleBusinessUnitName = GetAliasedString(e, "rolebu", "name"),
                        IsDefaultTeam = GetAliasedBoolean(e, "team", "isdefault") ?? false,

                    });
                }
            }

            return map;
        }

        private static Dictionary<Guid, OwnerTeamInfo> RetrieveOwnerTeams(IOrganizationService svc)
        {
            var qe = new QueryExpression("team")
            {
                // include isdefault so we can surface it in owner-team rows
                ColumnSet = new ColumnSet("teamid", "name", "teamtype", "businessunitid", "isdefault"),
                NoLock = true
            };
            qe.Criteria.AddCondition("teamtype", ConditionOperator.Equal, 0); // Owner

            var buLink = qe.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            buLink.EntityAlias = "teambu";
            buLink.Columns = new ColumnSet("name");

            var ec = RetrieveAll(svc, qe);

            var dict = new Dictionary<Guid, OwnerTeamInfo>();
            foreach (var e in ec.Entities)
            {
                dict[e.Id] = new OwnerTeamInfo
                {
                    TeamId = e.Id,
                    Name = e.GetAttributeValue<string>("name"),
                    BusinessUnitName = GetAliasedString(e, "teambu", "name"),
                    IsDefaultTeam = e.GetAttributeValue<bool?>("isdefault")
                };
            }

            return dict;
        }

        private static List<TeamRoleInfo> RetrieveOwnerTeamRoles(IOrganizationService svc)
        {
            var qe = new QueryExpression("teamroles")
            {
                ColumnSet = new ColumnSet("teamid", "roleid"),
                NoLock = true
            };

            // Filter Owner Teams via join
            var teamLink = qe.AddLink("team", "teamid", "teamid", JoinOperator.Inner);
            teamLink.EntityAlias = "team";
            teamLink.Columns = new ColumnSet("teamid", "name", "teamtype", "businessunitid");
            teamLink.LinkCriteria.AddCondition("teamtype", ConditionOperator.Equal, 0);

            var teamBuLink = teamLink.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            teamBuLink.EntityAlias = "teambu";
            teamBuLink.Columns = new ColumnSet("name");

            var roleLink = qe.AddLink("role", "roleid", "roleid", JoinOperator.Inner);
            roleLink.EntityAlias = "role";
            roleLink.Columns = new ColumnSet("roleid", "name", "businessunitid");

            var roleBuLink = roleLink.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            roleBuLink.EntityAlias = "rolebu";
            roleBuLink.Columns = new ColumnSet("name");

            var ec = RetrieveAll(svc, qe);

            var list = new List<TeamRoleInfo>();
            var dedupe = new HashSet<string>();

            foreach (var e in ec.Entities)
            {
                var teamId = GetIdFromAttribute(e, "teamid");
                var roleId = GetIdFromAttribute(e, "roleid");
                if (teamId == Guid.Empty || roleId == Guid.Empty) continue;

                var key = $"{teamId:N}|{roleId:N}";
                if (!dedupe.Add(key)) continue;

                list.Add(new TeamRoleInfo
                {
                    TeamId = teamId,
                    TeamName = GetAliasedString(e, "team", "name"),
                    TeamBusinessUnitName = GetAliasedString(e, "teambu", "name"),
                    RoleId = roleId,
                    RoleName = GetAliasedString(e, "role", "name"),
                    RoleBusinessUnitName = GetAliasedString(e, "rolebu", "name")
                });
            }

            return list;
        }

        private static List<BusinessUnitInfo> RetrieveAllBusinessUnits(IOrganizationService svc)
        {
            var qe = new QueryExpression("businessunit")
            {
                ColumnSet = new ColumnSet("businessunitid", "name", "parentbusinessunitid"),
                NoLock = true
            };

            var ec = RetrieveAll(svc, qe);

            var list = new List<BusinessUnitInfo>();
            foreach (var e in ec.Entities)
            {
                list.Add(new BusinessUnitInfo
                {
                    BusinessUnitId = e.Id,
                    Name = e.GetAttributeValue<string>("name"),
                    ParentBusinessUnitId = e.GetAttributeValue<EntityReference>("parentbusinessunitid")?.Id ?? Guid.Empty
                });
            }

            return list;
        }
        private static List<string> RetrieveAllBusinessUnitNames(IOrganizationService svc)
        {
            var qe = new QueryExpression("businessunit")
            {
                ColumnSet = new ColumnSet("name"),
                NoLock = true
            };
            qe.Orders.Add(new OrderExpression("name", OrderType.Ascending));

            var ec = RetrieveAll(svc, qe);

            return ec.Entities
                .Select(e => e.GetAttributeValue<string>("name"))
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<RoleInfo> RetrieveRolesByBusinessUnit(IOrganizationService svc, Guid businessUnitId)
        {
            var qe = new QueryExpression("role")
            {
                ColumnSet = new ColumnSet("roleid", "name", "businessunitid"),
                NoLock = true
            };

            qe.Criteria.AddCondition("businessunitid", ConditionOperator.Equal, businessUnitId);

            var ec = RetrieveAll(svc, qe);

            return ec.Entities.Select(e => new RoleInfo
            {
                RoleId = e.Id,
                Name = e.GetAttributeValue<string>("name"),
                BusinessUnitName = null
            }).ToList();
        }

        // -----------------------
        // Build rows
        // -----------------------
        private static List<MatrixRow> BuildUserRows(
            Dictionary<Guid, UserInfo> users,
            Dictionary<Guid, List<RoleInfo>> directRoles,
            Dictionary<Guid, List<TeamRoleInfo>> teamRoles)
        {
            var directPairs = new HashSet<(Guid UserId, Guid RoleId)>();
            foreach (var kvp in directRoles)
                foreach (var r in kvp.Value)
                    directPairs.Add((kvp.Key, r.RoleId));

            var teamPairs = new HashSet<(Guid UserId, Guid RoleId)>();
            foreach (var kvp in teamRoles)
                foreach (var r in kvp.Value)
                    teamPairs.Add((kvp.Key, r.RoleId));

            var duplicates = new HashSet<(Guid UserId, Guid RoleId)>(directPairs);
            duplicates.IntersectWith(teamPairs);

            var rows = new List<MatrixRow>(users.Count * 2);

            foreach (var u in users.Values.OrderBy(x => x.FullName, StringComparer.OrdinalIgnoreCase))
            {
                var any = false;

                if (directRoles.TryGetValue(u.UserId, out var dr))
                {
                    foreach (var r in dr.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase))
                    {
                        rows.Add(new MatrixRow
                        {
                            UserId = u.UserId,
                            TeamId = Guid.Empty,
                            RoleId = r.RoleId,

                            UserFullName = u.FullName,
                            UserEmail = u.Email,
                            UserBusinessUnit = u.BusinessUnitName,

                            AssignmentType = "Direct",
                            IsDefaultTeam = "",

                            TeamName = "",
                            TeamBusinessUnit = "",

                            RoleName = r.Name,
                            RoleBusinessUnit = r.BusinessUnitName,

                            Duplicate = duplicates.Contains((u.UserId, r.RoleId))
                        });

                        any = true;
                    }
                }

                if (teamRoles.TryGetValue(u.UserId, out var tr))
                {
                    foreach (var t in tr.OrderBy(x => x.TeamName, StringComparer.OrdinalIgnoreCase)
                                        .ThenBy(x => x.RoleName, StringComparer.OrdinalIgnoreCase))
                    {
                        rows.Add(new MatrixRow
                        {
                            UserId = u.UserId,
                            TeamId = t.TeamId,
                            RoleId = t.RoleId,

                            UserFullName = u.FullName,
                            UserEmail = u.Email,
                            UserBusinessUnit = u.BusinessUnitName,

                            AssignmentType = "Team",
                            IsDefaultTeam = t.IsDefaultTeam ? "Yes" : "No",

                            TeamName = t.TeamName,
                            TeamBusinessUnit = t.TeamBusinessUnitName,

                            RoleName = t.RoleName,
                            RoleBusinessUnit = t.RoleBusinessUnitName,

                            Duplicate = duplicates.Contains((u.UserId, t.RoleId))
                        });

                        any = true;
                    }
                }

                if (!any)
                {
                    rows.Add(new MatrixRow
                    {
                        UserId = u.UserId,
                        TeamId = Guid.Empty,
                        RoleId = Guid.Empty,

                        UserFullName = u.FullName,
                        UserEmail = u.Email,
                        UserBusinessUnit = u.BusinessUnitName,

                        AssignmentType = "None",

                        TeamName = "",
                        TeamBusinessUnit = "",

                        RoleName = "",
                        RoleBusinessUnit = "",

                        Duplicate = false
                    });
                }
            }

            return rows;
        }

        private static List<MatrixRow> BuildOwnerTeamRows(
            Dictionary<Guid, OwnerTeamInfo> teams,
            List<TeamRoleInfo> teamRoles)
        {
            var rows = new List<MatrixRow>();
            var teamsWithRoles = new HashSet<Guid>();

            foreach (var tr in teamRoles)
            {
                teamsWithRoles.Add(tr.TeamId);

                rows.Add(new MatrixRow
                {
                    UserId = Guid.Empty,
                    TeamId = tr.TeamId,
                    RoleId = tr.RoleId,

                    UserFullName = "",
                    UserEmail = "",
                    UserBusinessUnit = "",

                    AssignmentType = "Team",

                    TeamName = tr.TeamName,
                    TeamBusinessUnit = tr.TeamBusinessUnitName,

                    RoleName = tr.RoleName,
                    RoleBusinessUnit = tr.RoleBusinessUnitName,

                    Duplicate = false,
                    IsDefaultTeam = tr.IsDefaultTeam ? "Yes" : "No"
                });
            }

            // Include Owner Teams with no roles
            foreach (var t in teams.Values.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase))
            {
                if (teamsWithRoles.Contains(t.TeamId)) continue;

                var isDefaultText = t.IsDefaultTeam.HasValue ? (t.IsDefaultTeam.Value ? "Yes" : "No") : "";

                rows.Add(new MatrixRow
                {
                    UserId = Guid.Empty,
                    TeamId = t.TeamId,
                    RoleId = Guid.Empty,

                    UserFullName = "",
                    UserEmail = "",
                    UserBusinessUnit = "",

                    AssignmentType = "Team",

                    TeamName = t.Name,
                    TeamBusinessUnit = t.BusinessUnitName,

                    RoleName = "",
                    RoleBusinessUnit = "",

                    Duplicate = false,
                    IsDefaultTeam = isDefaultText
                });
            }

            return rows;
        }

        // -----------------------
        // Paging helper
        // -----------------------
        private static EntityCollection RetrieveAll(IOrganizationService svc, QueryExpression qe, int pageSize = 5000)
        {
            var result = new EntityCollection();

            qe.PageInfo = new PagingInfo
            {
                PageNumber = 1,
                Count = pageSize,
                ReturnTotalRecordCount = false
            };

            while (true)
            {
                var ec = svc.RetrieveMultiple(qe);
                result.Entities.AddRange(ec.Entities);

                if (!ec.MoreRecords) break;

                qe.PageInfo.PageNumber++;
                qe.PageInfo.PagingCookie = ec.PagingCookie;
            }

            return result;
        }

        // -----------------------
        // Aliased helpers
        // -----------------------
        private void HideColumn(string columnName)
        {
            if (dgvResults.Columns[columnName] != null)
                dgvResults.Columns[columnName].Visible = false;
        }

        private void SetColumnVisible(string columnName, bool visible)
        {
            if (dgvResults.Columns[columnName] != null)
                dgvResults.Columns[columnName].Visible = visible;
        }

        private void SetFixedColumnWidth(string columnName, int width)
        {
            if (dgvResults.Columns[columnName] == null) return;

            var col = dgvResults.Columns[columnName];
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            col.Width = width;
            col.MinimumWidth = width;

            // Prevent header wrapping for these columns
            col.HeaderCell.Style.WrapMode = DataGridViewTriState.False;
        }


        private static string GetAliasedString(Entity e, string alias, string attribute)
        {
            var key = $"{alias}.{attribute}";
            if (!e.Attributes.TryGetValue(key, out var obj) || obj == null)
                return null;

            if (obj is AliasedValue av)
                return av.Value?.ToString();

            return obj.ToString();
        }

        private static Guid GetAliasedGuid(Entity e, string alias, string attribute)
        {
            var key = $"{alias}.{attribute}";
            if (!e.Attributes.TryGetValue(key, out var obj) || obj == null)
                return Guid.Empty;

            if (obj is AliasedValue av)
            {
                if (av.Value is Guid g) return g;
                if (av.Value is EntityReference er) return er.Id;
            }

            return Guid.Empty;
        }

        private static Guid GetIdFromAttribute(Entity e, string attributeLogicalName)
        {
            if (!e.Attributes.TryGetValue(attributeLogicalName, out var obj) || obj == null)
                return Guid.Empty;

            if (obj is Guid g) return g;
            if (obj is EntityReference er) return er.Id;

            return Guid.Empty;
        }
        private static bool? GetAliasedBoolean(Entity e, string alias, string attribute)
        {
            var key = $"{alias}.{attribute}";
            if (!e.Attributes.TryGetValue(key, out var obj) || obj == null)
                return null;

            if (obj is AliasedValue av)
                obj = av.Value;

            return obj is bool b ? b : (bool?)null;
        }

        // -----------------------
        // Models
        // -----------------------
        private sealed class MatrixRow
        {
            public Guid UserId { get; set; }
            public Guid TeamId { get; set; }
            public Guid RoleId { get; set; }

            public string UserFullName { get; set; }
            public string UserEmail { get; set; }
            public string UserBusinessUnit { get; set; }

            public string AssignmentType { get; set; }

            public string TeamName { get; set; }
            public string TeamBusinessUnit { get; set; }

            public string RoleName { get; set; }
            public string RoleBusinessUnit { get; set; }

            public bool Duplicate { get; set; }
            public string IsDefaultTeam { get; set; }

        }

        private sealed class UserInfo
        {
            public Guid UserId { get; set; }
            public string FullName { get; set; }
            public string Email { get; set; }
            public string BusinessUnitName { get; set; }
            public bool IsDisabled { get; set; }
        }

        private sealed class BusinessUnitInfo
        {
            public Guid BusinessUnitId { get; set; }
            public string Name { get; set; }
            public Guid ParentBusinessUnitId { get; set; }
        }

        private sealed class RoleInfo
        {
            public Guid RoleId { get; set; }
            public string Name { get; set; }
            public string BusinessUnitName { get; set; }
        }

        private sealed class OwnerTeamInfo
        {
            public Guid TeamId { get; set; }
            public string Name { get; set; }
            public string BusinessUnitName { get; set; }
            public bool? IsDefaultTeam { get; set; }
        }

        private sealed class TeamRoleInfo
        {
            public Guid TeamId { get; set; }
            public string TeamName { get; set; }
            public string TeamBusinessUnitName { get; set; }
            public Guid RoleId { get; set; }
            public string RoleName { get; set; }
            public string RoleBusinessUnitName { get; set; }
            public bool IsDefaultTeam { get; set; }

        }

        private sealed class DeleteOpComparer : IEqualityComparer<DeleteOp>
        {
            public bool Equals(DeleteOp x, DeleteOp y) => x?.Key == y?.Key;
            public int GetHashCode(DeleteOp obj) => obj.Key?.GetHashCode() ?? 0;
        }
    }

    static class DependencyResolver
    {
        private static bool _registered;

        public static void Register()
        {
            if (_registered) return;
            _registered = true;

            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
        }

        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                // Requested simple name (e.g. ClosedXML)
                var requestedName = new AssemblyName(args.Name).Name + ".dll";

                // Folder where the plugin dll is located (Plugins\)
                var pluginDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                // Dedicated dependency folder under Plugins\
                var depsDir = Path.Combine(pluginDir, "GM.XrmToolBox.UserRoleMatrix");

                var candidatePath = Path.Combine(depsDir, requestedName);

                if (File.Exists(candidatePath))
                {
                    return Assembly.LoadFrom(candidatePath);
                }
            }
            catch
            {
                // ignore and let default loader continue
            }

            return null;
        }
    }
    // ...
}





