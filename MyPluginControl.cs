using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Xml.Linq;
using XrmToolBox.Extensibility;
using Button = System.Windows.Forms.Button;
using Color = System.Drawing.Color;
using ComboBox = System.Windows.Forms.ComboBox;
using Font = System.Drawing.Font;
using Image = System.Drawing.Image;
using Label = System.Windows.Forms.Label;
using ToolTip = System.Windows.Forms.ToolTip;

namespace GM.XrmToolBox.UserRoleMatrix
{
    public partial class MyPluginControl : PluginControlBase
    {
        private readonly BindingSource _bindingSource = new BindingSource();
        private DataTable _table;
        private DataView _view;

        // Global mode: "systemuser" or "team" or ""
        private const string ModeNone = "";
        private const string ModeSystemUser = "systemuser";
        private const string ModeTeam = "team";
        private string _currentMode = ModeNone;

        private bool _updatingFilters;

        private List<string> _allBusinessUnitNames = new List<string>();
        private List<string> _allTeamNames = new List<string>();
        private bool _internalSelectAllChange;

        // Persist all owner team names (so Team dropdown can show all teams, not just those in grid)
        private OrgSettingStatus _orgSettingStatus = OrgSettingStatus.Unknown();

        // Icons
        private const string AboutIconBase64 =
            "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAADf0lEQVR42u2Xa0hTYRjH/2eXc3brzNXUuUJDaxOnrVgXukmEZZkfihYliRRBYWQWRCn1qaB1nUVhsT5UUISaIFnR5UPQfUZmMaNsoEsj03ltc3q2efrQWrNpTmquDz1wvpyH9/3/zvM+z3OeF4iwESN6iizsX1U6kkqEBlBkYVmDJiEsX1tcb/sVhDecOFFcbwsHAGvQJBCwsIEQnPES/xEB1qBJCDxePwnLsvHhFA+KBEF8/BmBv51woZhPkxPpMvwPEHEA3kgOikcgf95ErEujoYmlICY5cDCDaHd48d4+gJxrLXAwg7izOQGZKgkAoMHOINn4AawvpUkugaZ9KsRN+C5zta4HuWUto0dALubCvD0RJdkK8Fwd3cv1m0olsuhNM1NU+cUFW05L7PUN1O19RjB9/YHrVHISWY3Gc/AhbNBK/eIAgPqbD1FdVDJqBExrlNDGCeDsczHZ6brCdp4iFmvO7rfFqKfaHPbuSuNzK9pau+Bl3AAEAMAwjIckSd6uLTkLbpke1SIpXbdr4SQE+kI6glgJD6tTaABARXnZ4/ZuB4P80gMQyb6/lCqjIVVGQ71sfuC6urq6Rqk0SpyRkaHVHL9YGpO4UjdLKYC59s3HKDHFVavVk0NKwpQYCoSvP1qt1s9IzlwAkYy+sn4KWIPG/+xZLB/a3ViWPX2p3AwAhfol03fPFbgAoOTY4QrwKDLkKiCIIZsCsvg4AMgta4FQt/aM3/ngxGWYL1UFrr1cUf2yy+Fy5+XlLVmVFito/tzWW1lZ+QwCWhIywNu2AX8WJyUlKYYQZR0q+F1J9bn63Rde9HAoiuJzOARx5tTJKo9ILgMpFoYM0PrVg2pLhxsA9Hr9/ChPR9dY6vqsuZfr8Q6yTqez/4LJdA+6javG3Ii23rDz31ptdpqmRVWGbYu0E70ukksgTSEYFaC5xw3+zgcdElqa2+0c8EKrXzbmRvTF4cGc8838Amn5jXVZSzWPt0/jCoRCtrd/kGho/eq2PL3/sqam5gO4c2cPuyutkGPv6+t/1An7+DLp0Z6lmUf3XL2Fd3fPobPpEzwDbpASIUQyGtHT4jE7dfqKizagbNtBND55BeUM1XB7JRutgCl7BzqbPkGTnR48E47DNBQ0kPjmw/9/w38EYIRLQ1jNp8kJGpnHKwGDjuBIKhFuiMDs/2euZhG/nEbcvgEhmWjf6ekl5gAAAABJRU5ErkJggg==";

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
            "Duplicate"
        };

        public MyPluginControl()
        {
            InitializeComponent();

            // Dependency resolver for ClosedXML + dependencies in dedicated folder
            DependencyResolver.Register();

            tsbAbout.Image = ImageFromBase64(AboutIconBase64);
            tsbAbout.Click += (s, e) => ShowAboutDialog();

            // Grid config
            dgvResults.AutoGenerateColumns = true;
            dgvResults.AllowUserToAddRows = false;
            dgvResults.AllowUserToDeleteRows = false;
            dgvResults.ReadOnly = false;
            dgvResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvResults.MultiSelect = true;

            dgvResults.DataSource = _bindingSource;
            dgvResults.DataBindingComplete += (s, e) =>
            {
                ConfigureGridColumns();
                ApplyDuplicateRowHighlight();
            };

            // Toolstrip events
            tsbLoadUsersRoles.Click += (s, e) => ExecuteMethod(LoadUsersAndRoles);
            tsbLoadOwnerTeamsRoles.Click += (s, e) => ExecuteMethod(LoadOwnerTeamsRoles);

            tsbSelectAllRecords.CheckedChanged += (s, e) =>
            {
                if (_internalSelectAllChange) return;
                ToggleSelectAllRecords(tsbSelectAllRecords.Checked);
            };

            tsbAddUserRole.Click += (s, e) => ExecuteMethod(OpenAddUserRoleDialog);
            tsbDel.Click += (s, e) => ExecuteMethod(DeleteSelectedAssignments);

            tsmiExportCsv.Click += (s, e) => ExportCsv();
            tsmiExportExcel.Click += (s, e) => ExportExcel();

            // Filters/search
            tscBusinessUnit.SelectedIndexChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };
            tscTeam.SelectedIndexChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };
            tscAssignment.SelectedIndexChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };
            tstSearch.TextChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };

            InitializeStaticFilters();
            UpdateCountLabel(0, 0);

            UpdateConnectionControls(false);
            UpdateActionButtonsEnabledState();
        }

        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail connectionDetail, string actionName = "", object parameter = null)
        {
            base.UpdateConnection(newService, connectionDetail, actionName, parameter);

            if (newService == null)
            {
                ClearResults();
                UpdateConnectionControls(false);
                return;
            }

            // When we connect, enable buttons and read org setting status asynchronously
            UpdateConnectionControls(true);
            ReadOrgSettingsAsync();
        }

        private void UpdateConnectionControls(bool connected)
        {
            tsbLoadUsersRoles.Enabled = connected;
            tsbLoadOwnerTeamsRoles.Enabled = connected;
            tsbAddUserRole.Enabled = connected;
            tsbDel.Enabled = connected;
            tsddExport.Enabled = connected;
        }

        private void UpdateActionButtonsEnabledState()
        {
            var hasData = _table != null && _table.Rows.Count > 0;

            tsbSelectAllRecords.Enabled = hasData;
            if (!hasData)
                SetSelectAllRecordsChecked(false);

            tsbDel.Enabled = hasData && (_currentMode == ModeSystemUser || _currentMode == ModeTeam);
            tsbAddUserRole.Enabled = true; // always available, it reloads user view after add
        }

        private void InitializeStaticFilters()
        {
            _updatingFilters = true;
            try
            {
                tscBusinessUnit.Items.Clear();
                tscBusinessUnit.Items.Add("All");
                tscBusinessUnit.SelectedIndex = 0;

                tscTeam.Items.Clear();
                tscTeam.Items.Add("All");
                tscTeam.SelectedIndex = 0;

                tscAssignment.Items.Clear();
                tscAssignment.Items.Add("All");
                tscAssignment.Items.Add("Direct");
                tscAssignment.Items.Add("Team");
                tscAssignment.SelectedIndex = 0;

                tstSearch.Text = "";
            }
            finally
            {
                _updatingFilters = false;
            }
        }

        private void UpdateCountLabel(int filtered, int total)
        {
            var settingText = _orgSettingStatus.IsKnown
                ? (_orgSettingStatus.EnableOwnershipAcrossBusinessUnits ? "EnableOwnershipAcrossBUs: ON" : "EnableOwnershipAcrossBUs: OFF")
                : "EnableOwnershipAcrossBUs: Unknown";

            tslCount.Text = $"Rows: {filtered}/{total} | {settingText}";
        }

        private void SetWorkingMessage(string message)
        {
            // XrmToolBox already shows "Working..." message; we can optionally extend. No-op here.
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

        private void SetSelectAllRecordsChecked(bool isChecked)
        {
            if (tsbSelectAllRecords == null) return;
            _internalSelectAllChange = true;
            try
            {
                tsbSelectAllRecords.Checked = isChecked;
            }
            finally
            {
                _internalSelectAllChange = false;
            }
        }

        private void ToggleSelectAllRecords(bool select)
        {
            if (_view == null || dgvResults.Rows.Count == 0) return;

            // Commit any pending edits in the grid
            dgvResults.EndEdit();

            _bindingSource.RaiseListChangedEvents = false;
            try
            {
                foreach (DataGridViewRow gridRow in dgvResults.Rows)
                {
                    if (gridRow.DataBoundItem is DataRowView drv)
                    {
                        drv.Row["Selected"] = select;
                    }
                }
            }
            finally
            {
                _bindingSource.RaiseListChangedEvents = true;
                _bindingSource.ResetBindings(false);
            }
        }

        private void ShowAboutDialog()
        {
            using (var f = new AboutForm())
            {
                f.StartPosition = FormStartPosition.CenterParent;
                f.ShowDialog(this);
            }
        }

        private void ClearResults()
        {
            _table = null;
            _view = null;
            _bindingSource.DataSource = null;

            _currentMode = ModeNone;

            SetSelectAllRecordsChecked(false);
            InitializeStaticFilters();
            UpdateCountLabel(0, 0);
            UpdateActionButtonsEnabledState();
        }

        // -----------------------
        // Feature #1: Org setting label (EnableOwnershipAcrossBusinessUnits)
        // -----------------------
        private void ReadOrgSettingsAsync()
        {
            _orgSettingStatus = OrgSettingStatus.Unknown();

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
            // Under the hood this is stored in organization.orgdborgsettings as XML.
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

                // Parse XML and get the EnableOwnershipAcrossBusinessUnits element
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

                    worker.ReportProgress(0, "Retrieving all Owner Teams...");
                    var allTeams = RetrieveAllOwnerTeamNames(Service);

                    args.Result = new LoadResult
                    {
                        Rows = rows,
                        BusinessUnits = allBus,
                        Teams = allTeams
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
                    // Persist team names for Team filter
                    _allTeamNames = result.Teams ?? new List<string>();

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
                        BusinessUnits = allBus,
                        Teams = teams?.Values
                                    .Select(t => t.Name)
                                    .Where(n => !string.IsNullOrWhiteSpace(n))
                                    .Distinct(StringComparer.OrdinalIgnoreCase)
                                    .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                                    .ToList()
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
                    _allTeamNames = result.Teams ?? new List<string>();

                    BindRows(result.Rows);
                    PopulateDropdownFiltersFromTable();
                    ApplyAllFilters();

                    UpdateActionButtonsEnabledState();
                }
            });
        }

        // -----------------------
        // Feature #3: Delete Selected (DisassociateRequest) - now fast via ExecuteMultiple
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

                    const int batchSize = 500;

                    for (int i = 0; i < ops.Count; i += batchSize)
                    {
                        var batchOps = ops.Skip(i).Take(batchSize).ToList();

                        var emReq = new ExecuteMultipleRequest
                        {
                            Settings = new ExecuteMultipleSettings
                            {
                                ContinueOnError = true,
                                ReturnResponses = true
                            },
                            Requests = new OrganizationRequestCollection()
                        };

                        foreach (var op in batchOps)
                        {
                            emReq.Requests.Add(new DisassociateRequest
                            {
                                Target = op.Target,
                                Relationship = new Relationship(op.RelationshipName),
                                RelatedEntities = new EntityReferenceCollection
                                {
                                    op.Related
                                }
                            });
                        }

                        try
                        {
                            var emResp = (ExecuteMultipleResponse)Service.Execute(emReq);

                            var failed = new HashSet<int>();
                            foreach (var respItem in emResp.Responses)
                            {
                                if (respItem.Fault != null)
                                {
                                    failed.Add(respItem.RequestIndex);

                                    var key = (respItem.RequestIndex >= 0 && respItem.RequestIndex < batchOps.Count)
                                        ? batchOps[respItem.RequestIndex].Key
                                        : $"batchIndex:{i}+{respItem.RequestIndex}";

                                    errors.Add($"{key}: {respItem.Fault.Message}");
                                }
                            }

                            success += batchOps.Count - failed.Count;
                        }
                        catch
                        {
                            // Fallback: execute individually for this batch
                            foreach (var op in batchOps)
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
                                catch (Exception ex2)
                                {
                                    errors.Add($"{op.Key}: {ex2.Message}");
                                }
                            }
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
            private readonly System.Windows.Forms.Button _btnAdd = new Button();
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
                _btnAdd.Click += (s, e) => AddClicked();

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
                    var baseText = string.IsNullOrWhiteSpace(u.Email)
                        ? u.FullName
                        : $"{u.FullName} ({u.Email})";

                    var buSuffix = string.IsNullOrWhiteSpace(u.BusinessUnitName)
                        ? string.Empty
                        : $" - {u.BusinessUnitName}";

                    _cbUser.Items.Add(new ComboItem(u.UserId, baseText + buSuffix));
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
                // Keep the combobox text empty while loading to avoid a lingering "Loading..." placeholder
                _cbRole.Text = string.Empty;

                try
                {
                    var roles = await System.Threading.Tasks.Task.Run(() => RetrieveRolesByBusinessUnit(_service, buItem.Id));

                    if (IsDisposed) return;

                    BeginInvoke(new Action(() =>
                    {
                        _cbRole.Items.Clear();
                        foreach (var r in roles.OrderBy(r => r.Name, StringComparer.OrdinalIgnoreCase))
                            _cbRole.Items.Add(new ComboItem(r.RoleId, r.Name));

                        _cbRole.Enabled = true;
                        // Select first item only if we have entries; otherwise leave blank
                        if (_cbRole.Items.Count > 0)
                            _cbRole.SelectedIndex = 0;
                        else
                            _cbRole.Text = string.Empty;
                    }));
                }
                catch
                {
                    if (IsDisposed) return;

                    BeginInvoke(new Action(() =>
                    {
                        _cbRole.Items.Clear();
                        _cbRole.Text = string.Empty;
                        _cbRole.Enabled = true;
                    }));
                }
            }

            private void AddClicked()
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

                Exception executeException = null;

                using (var waiting = new WaitingForm("waiting adding record"))
                {
                    // Execute on background thread and close waiting dialog when done
                    System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            _service.Execute(req);
                        }
                        catch (Exception ex)
                        {
                            executeException = ex;
                        }
                        finally
                        {
                            try
                            {
                                if (!waiting.IsDisposed)
                                    waiting.BeginInvoke(new Action(() => { try { waiting.Close(); } catch { } }));
                            }
                            catch
                            {
                                // ignore
                            }
                        }
                    });

                    // Show modal waiting dialog; it will be closed by background task
                    waiting.ShowDialog(this);
                }

                if (executeException != null)
                {
                    MessageBox.Show(executeException.Message, "Add User Role", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                DialogResult = DialogResult.OK;
                Close();
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

            private sealed class WaitingForm : Form
            {
                public WaitingForm(string message)
                {
                    FormBorderStyle = FormBorderStyle.FixedDialog;
                    StartPosition = FormStartPosition.CenterParent;
                    ShowInTaskbar = false;
                    ControlBox = false;
                    Width = 360;
                    Height = 110;
                    BackColor = Color.LightYellow;

                    var lbl = new Label
                    {
                        Text = message ?? "Please wait...",
                        AutoSize = false,
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 10, FontStyle.Bold)
                    };

                    Controls.Add(lbl);

                    Shown += (s, e) => Cursor = Cursors.WaitCursor;
                    FormClosed += (s, e) => Cursor = Cursors.Default;
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
                ShowInTaskbar = false;
                Width = 560;
                Height = 260;

                var pic = new PictureBox
                {
                    Left = 16,
                    Top = 20,
                    Width = 64,
                    Height = 64,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Image = ImageFromBase64(AboutIconBase64)
                };

                var lblName = new Label
                {
                    Left = 96,
                    Top = 20,
                    Width = 440,
                    Text = "Giovanni Manunta",
                    Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold)
                };

                var linkLinkedIn = new LinkLabel { Left = 96, Top = 80, Width = 440, Text = "LinkedIn profile" };
                var linkGitHub = new LinkLabel { Left = 96, Top = 110, Width = 440, Text = "GitHub profile" };

                var btnOk = new Button { Text = "OK", Left = 440, Top = 160, Width = 80, DialogResult = DialogResult.OK };
                AcceptButton = btnOk;

                linkLinkedIn.LinkClicked += (s, e) => OpenUrl("https://www.linkedin.com/in/giovanni-manunta-3555868/");
                linkGitHub.LinkClicked += (s, e) => OpenUrl("https://github.com/gmanunta81");

                Controls.Add(pic);
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
                    // ignore
                }
            }
        }

        private sealed class LoadResult
        {
            public List<MatrixRow> Rows { get; set; }
            public List<string> BusinessUnits { get; set; }
            // added to carry team names back to UI thread
            public List<string> Teams { get; set; }
        }

        // -----------------------
        // Table binding
        // -----------------------
        private void BindRows(List<MatrixRow> rows)
        {
            SetSelectAllRecordsChecked(false);

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

                dr["Duplicate"] = r.Duplicate;

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
            dt.Columns.Add("Duplicate", typeof(bool));

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
            if (dgvResults.Rows.Count == 0 || dgvResults.Columns["Duplicate"] == null)
                return;

            var normalBack = dgvResults.RowsDefaultCellStyle.BackColor;
            if (normalBack == Color.Empty)
                normalBack = SystemColors.Window;

            foreach (DataGridViewRow row in dgvResults.Rows)
            {
                var value = row.Cells["Duplicate"]?.Value;
                var isDup = value is bool b && b;
                row.DefaultCellStyle.BackColor = isDup ? Color.LightYellow : normalBack;
            }
        }

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
            if (dgvResults.Columns[columnName] != null)
            {
                dgvResults.Columns[columnName].Width = width;
                dgvResults.Columns[columnName].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
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
                var businessUnits = (_allBusinessUnitNames ?? new List<string>())
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var teamsInTable = _table.AsEnumerable()
                    .Select(r => r.Field<string>("Team"))
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var teams = (_allTeamNames ?? new List<string>())
                    .Union(teamsInTable, StringComparer.OrdinalIgnoreCase)
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

                        var openNow = MessageBox.Show("Excel export completed.\n\nOpen the file now?", "Export", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        if (openNow == DialogResult.Yes)
                        {
                            try
                            {
                                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path)
                                {
                                    UseShellExecute = true
                                });
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message, "Open file", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
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

        private static List<string> RetrieveAllOwnerTeamNames(IOrganizationService svc)
        {
            var qe = new QueryExpression("team")
            {
                ColumnSet = new ColumnSet("name"),
                NoLock = true
            };
            qe.Criteria.AddCondition("teamtype", ConditionOperator.Equal, 0); // Owner Teams
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

            var buLink = qe.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            buLink.EntityAlias = "rolebu";
            buLink.Columns = new ColumnSet("name");

            var ec = RetrieveAll(svc, qe);

            var roles = new List<RoleInfo>();
            foreach (var e in ec.Entities)
            {
                roles.Add(new RoleInfo
                {
                    RoleId = e.Id,
                    Name = e.GetAttributeValue<string>("name"),
                    BusinessUnitName = GetAliasedString(e, "rolebu", "name")
                });
            }

            return roles;
        }

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
            // systemuserroles: intersect entity systemuser <-> role
            var qe = new QueryExpression("systemuserroles")
            {
                ColumnSet = new ColumnSet("systemuserid", "roleid"),
                NoLock = true
            };

            var roleLink = qe.AddLink("role", "roleid", "roleid", JoinOperator.Inner);
            roleLink.EntityAlias = "role";
            roleLink.Columns = new ColumnSet("name", "businessunitid");

            var roleBuLink = roleLink.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            roleBuLink.EntityAlias = "rolebu";
            roleBuLink.Columns = new ColumnSet("name");

            var ec = RetrieveAll(svc, qe);

            var dict = new Dictionary<Guid, List<RoleInfo>>();

            foreach (var e in ec.Entities)
            {
                var userId = GetIdFromAttribute(e, "systemuserid");
                var roleId = GetIdFromAttribute(e, "roleid");

                if (userId == Guid.Empty || roleId == Guid.Empty) continue;

                var role = new RoleInfo
                {
                    RoleId = roleId,
                    Name = GetAliasedString(e, "role", "name"),
                    BusinessUnitName = GetAliasedString(e, "rolebu", "name")
                };

                if (!dict.TryGetValue(userId, out var list))
                {
                    list = new List<RoleInfo>();
                    dict[userId] = list;
                }

                // de-dupe by role id
                if (!list.Any(r => r.RoleId == role.RoleId))
                    list.Add(role);
            }

            return dict;
        }

        private static Dictionary<Guid, List<TeamRoleInfo>> RetrieveTeamUserRoles(IOrganizationService svc)
        {
            // For each user: their teams -> those teams' roles
            // We retrieve team membership via teammembership and then teamroles.

            // 1) teammembership: systemuser <-> team
            var qeMember = new QueryExpression("teammembership")
            {
                ColumnSet = new ColumnSet("systemuserid", "teamid"),
                NoLock = true
            };

            // join team to filter owner teams only
            var teamLink = qeMember.AddLink("team", "teamid", "teamid", JoinOperator.Inner);
            teamLink.EntityAlias = "team";
            teamLink.Columns = new ColumnSet("name", "teamtype", "businessunitid", "isdefault");
            teamLink.LinkCriteria.AddCondition("teamtype", ConditionOperator.Equal, 0); // Owner

            var teamBuLink = teamLink.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            teamBuLink.EntityAlias = "teambu";
            teamBuLink.Columns = new ColumnSet("name");

            var memberEc = RetrieveAll(svc, qeMember);

            // Build membership list
            var membership = new List<(Guid UserId, Guid TeamId, string TeamName, string TeamBu, bool IsDefault)>();
            foreach (var e in memberEc.Entities)
            {
                var userId = GetIdFromAttribute(e, "systemuserid");
                var teamId = GetIdFromAttribute(e, "teamid");
                if (userId == Guid.Empty || teamId == Guid.Empty) continue;

                membership.Add((userId, teamId,
                    GetAliasedString(e, "team", "name"),
                    GetAliasedString(e, "teambu", "name"),
                    GetAliasedBoolean(e, "team", "isdefault") ?? false));
            }

            // 2) teamroles: team <-> role
            var qeTeamRoles = new QueryExpression("teamroles")
            {
                ColumnSet = new ColumnSet("teamid", "roleid"),
                NoLock = true
            };

            // join role for name and BU
            var roleLink = qeTeamRoles.AddLink("role", "roleid", "roleid", JoinOperator.Inner);
            roleLink.EntityAlias = "role";
            roleLink.Columns = new ColumnSet("name", "businessunitid");

            var roleBuLink = roleLink.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            roleBuLink.EntityAlias = "rolebu";
            roleBuLink.Columns = new ColumnSet("name");

            // join team to filter owner teams and get BU
            var teamLink2 = qeTeamRoles.AddLink("team", "teamid", "teamid", JoinOperator.Inner);
            teamLink2.EntityAlias = "team";
            teamLink2.Columns = new ColumnSet("name", "teamtype", "businessunitid", "isdefault");
            teamLink2.LinkCriteria.AddCondition("teamtype", ConditionOperator.Equal, 0); // Owner

            var teamBuLink2 = teamLink2.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            teamBuLink2.EntityAlias = "teambu";
            teamBuLink2.Columns = new ColumnSet("name");

            var teamRoleEc = RetrieveAll(svc, qeTeamRoles);

            // Map: teamId -> roles
            var teamRolesMap = new Dictionary<Guid, List<TeamRoleInfo>>();
            foreach (var e in teamRoleEc.Entities)
            {
                var teamId = GetIdFromAttribute(e, "teamid");
                var roleId = GetIdFromAttribute(e, "roleid");
                if (teamId == Guid.Empty || roleId == Guid.Empty) continue;

                var tri = new TeamRoleInfo
                {
                    TeamId = teamId,
                    TeamName = GetAliasedString(e, "team", "name"),
                    TeamBusinessUnitName = GetAliasedString(e, "teambu", "name"),
                    RoleId = roleId,
                    RoleName = GetAliasedString(e, "role", "name"),
                    RoleBusinessUnitName = GetAliasedString(e, "rolebu", "name"),
                    IsDefaultTeam = GetAliasedBoolean(e, "team", "isdefault") ?? false
                };

                if (!teamRolesMap.TryGetValue(teamId, out var list))
                {
                    list = new List<TeamRoleInfo>();
                    teamRolesMap[teamId] = list;
                }

                if (!list.Any(x => x.RoleId == tri.RoleId))
                    list.Add(tri);
            }

            // 3) Build userId -> list of team-role rows
            var userMap = new Dictionary<Guid, List<TeamRoleInfo>>();

            foreach (var m in membership)
            {
                if (!teamRolesMap.TryGetValue(m.TeamId, out var roles))
                    continue;

                if (!userMap.TryGetValue(m.UserId, out var list))
                {
                    list = new List<TeamRoleInfo>();
                    userMap[m.UserId] = list;
                }

                foreach (var r in roles)
                {
                    // Ensure default-team info from membership is reflected
                    var copy = new TeamRoleInfo
                    {
                        TeamId = r.TeamId,
                        TeamName = r.TeamName ?? m.TeamName,
                        TeamBusinessUnitName = r.TeamBusinessUnitName ?? m.TeamBu,
                        RoleId = r.RoleId,
                        RoleName = r.RoleName,
                        RoleBusinessUnitName = r.RoleBusinessUnitName,
                        IsDefaultTeam = m.IsDefault
                    };

                    list.Add(copy);
                }
            }

            // De-dupe by TeamId+RoleId per user
            foreach (var k in userMap.Keys.ToList())
            {
                userMap[k] = userMap[k]
                    .GroupBy(x => $"{x.TeamId:N}|{x.RoleId:N}")
                    .Select(g => g.First())
                    .ToList();
            }

            return userMap;
        }

        private static Dictionary<Guid, OwnerTeamInfo> RetrieveOwnerTeams(IOrganizationService svc)
        {
            var qe = new QueryExpression("team")
            {
                ColumnSet = new ColumnSet("teamid", "name", "teamtype", "businessunitid"),
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
                    BusinessUnitName = GetAliasedString(e, "teambu", "name")
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

            // Filter Owner Teams via join (include isdefault so we can detect default teams)
            var teamLink = qe.AddLink("team", "teamid", "teamid", JoinOperator.Inner);
            teamLink.EntityAlias = "team";
            teamLink.Columns = new ColumnSet("teamid", "name", "teamtype", "businessunitid", "isdefault");
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
                    RoleBusinessUnitName = GetAliasedString(e, "rolebu", "name"),
                    IsDefaultTeam = GetAliasedBoolean(e, "team", "isdefault") ?? false
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
            qe.Orders.Add(new OrderExpression("name", OrderType.Ascending));

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

        private static List<MatrixRow> BuildUserRows(
            Dictionary<Guid, UserInfo> users,
            Dictionary<Guid, List<RoleInfo>> directRoles,
            Dictionary<Guid, List<TeamRoleInfo>> teamRoles)
        {
            var rows = new List<MatrixRow>();

            // Build a duplicate detector: same user + role name assigned both direct and via team
            var directKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kvp in directRoles)
            {
                var userId = kvp.Key;
                foreach (var r in kvp.Value)
                {
                    directKeys.Add($"{userId:N}|{r.RoleId:N}");
                }
            }

            // Direct assignments
            foreach (var kvp in directRoles)
            {
                var userId = kvp.Key;
                users.TryGetValue(userId, out var u);

                foreach (var role in kvp.Value)
                {
                    rows.Add(new MatrixRow
                    {
                        UserId = userId,
                        RoleId = role.RoleId,
                        TeamId = Guid.Empty,

                        UserFullName = u?.FullName,
                        UserEmail = u?.Email,
                        UserBusinessUnit = u?.BusinessUnitName,

                        AssignmentType = "Direct",

                        TeamName = "",
                        TeamBusinessUnit = "",
                        IsDefaultTeam = "",

                        RoleName = role.Name,
                        RoleBusinessUnit = role.BusinessUnitName,

                        Duplicate = false
                    });
                }
            }

            // Team-based assignments (via membership)
            foreach (var kvp in teamRoles)
            {
                var userId = kvp.Key;
                users.TryGetValue(userId, out var u);

                foreach (var tr in kvp.Value)
                {
                    var isDup = directKeys.Contains($"{userId:N}|{tr.RoleId:N}");

                    rows.Add(new MatrixRow
                    {
                        UserId = userId,
                        RoleId = tr.RoleId,
                        TeamId = tr.TeamId,

                        UserFullName = u?.FullName,
                        UserEmail = u?.Email,
                        UserBusinessUnit = u?.BusinessUnitName,

                        AssignmentType = "Team",

                        TeamName = tr.TeamName,
                        TeamBusinessUnit = tr.TeamBusinessUnitName,
                        IsDefaultTeam = tr.IsDefaultTeam ? "Yes" : "No",

                        RoleName = tr.RoleName,
                        RoleBusinessUnit = tr.RoleBusinessUnitName,

                        Duplicate = isDup
                    });
                }
            }

            // Mark duplicates on both rows (direct and team)
            // For direct rows, if a corresponding team assignment exists, set Duplicate = true.
            var dupKeys = rows
                .Where(r => r.AssignmentType == "Team" && r.Duplicate)
                .Select(r => $"{r.UserId:N}|{r.RoleId:N}")
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var r in rows)
            {
                if (r.AssignmentType == "Direct" && dupKeys.Contains($"{r.UserId:N}|{r.RoleId:N}"))
                    r.Duplicate = true;
            }

            return rows;
        }

        private static List<MatrixRow> BuildOwnerTeamRows(
            Dictionary<Guid, OwnerTeamInfo> teams,
            List<TeamRoleInfo> teamRoles)
        {
            var rows = new List<MatrixRow>();

            foreach (var tr in teamRoles)
            {
                teams.TryGetValue(tr.TeamId, out var t);

                rows.Add(new MatrixRow
                {
                    UserId = Guid.Empty,
                    RoleId = tr.RoleId,
                    TeamId = tr.TeamId,

                    UserFullName = "",
                    UserEmail = "",
                    UserBusinessUnit = "",

                    AssignmentType = "",

                    TeamName = tr.TeamName ?? t?.Name,
                    TeamBusinessUnit = tr.TeamBusinessUnitName ?? t?.BusinessUnitName,
                    IsDefaultTeam = (tr.IsDefaultTeam ? "Yes" : "No"),

                    RoleName = tr.RoleName,
                    RoleBusinessUnit = tr.RoleBusinessUnitName,

                    Duplicate = false
                });
            }

            return rows;
        }

        // -----------------------
        // Helpers
        // -----------------------
        private static string GetAliasedString(Entity e, string alias, string attributeName)
        {
            var key = $"{alias}.{attributeName}";
            if (!e.Attributes.ContainsKey(key)) return null;
            var val = e.Attributes[key] as AliasedValue;
            return val?.Value as string;
        }

        private static bool? GetAliasedBoolean(Entity e, string alias, string attributeName)
        {
            var key = $"{alias}.{attributeName}";
            if (!e.Attributes.ContainsKey(key)) return null;
            var val = e.Attributes[key] as AliasedValue;
            return val?.Value as bool?;
        }

        private static Guid GetIdFromAttribute(Entity e, string attributeName)
        {
            if (e == null || string.IsNullOrWhiteSpace(attributeName))
                return Guid.Empty;

            // Prefer direct attribute access to correctly handle multiple underlying types
            if (e.Attributes != null && e.Attributes.TryGetValue(attributeName, out var raw))
            {
                // Intersect/relationship entities often return Guid (not EntityReference)
                if (raw is Guid g) return g;

                if (raw is EntityReference er) return er.Id;

                if (raw is AliasedValue av)
                {
                    if (av.Value is Guid ag) return ag;
                    if (av.Value is EntityReference aer) return aer.Id;
                }
            }

            // Fallbacks using typed getters
            var maybeEr = e.GetAttributeValue<EntityReference>(attributeName);
            if (maybeEr != null) return maybeEr.Id;

            var maybeGuid = e.GetAttributeValue<Guid?>(attributeName);
            if (maybeGuid.HasValue) return maybeGuid.Value;

            return Guid.Empty;
        }

        private static EntityCollection RetrieveAll(IOrganizationService svc, QueryExpression qe, int pageSize = 5000)
        {
            if (qe.PageInfo == null)
                qe.PageInfo = new PagingInfo();

            qe.PageInfo.Count = pageSize;
            qe.PageInfo.PageNumber = 1;
            qe.PageInfo.PagingCookie = null;

            var all = new EntityCollection();
            bool more = true;

            while (more)
            {
                var resp = svc.RetrieveMultiple(qe);
                all.Entities.AddRange(resp.Entities);

                more = resp.MoreRecords;
                if (more)
                {
                    qe.PageInfo.PageNumber++;
                    qe.PageInfo.PagingCookie = resp.PagingCookie;
                }
            }

            return all;
        }

        private void ShowErrorDialog(Exception ex)
        {
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

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
}
