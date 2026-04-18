using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using McTools.Xrm.Connection;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
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


        // -----------------------
        // Views filter (SystemUser / Team)
        // -----------------------
        private enum ViewSource
        {
            System,
            Personal
        }

        private readonly struct ViewKey : IEquatable<ViewKey>
        {
            public ViewSource Source { get; }
            public Guid Id { get; }

            public bool IsAll => Id == Guid.Empty;

            public ViewKey(ViewSource source, Guid id)
            {
                Source = source;
                Id = id;
            }

            public static ViewKey All => new ViewKey(ViewSource.System, Guid.Empty);

            public bool Equals(ViewKey other)
            {
                if (IsAll && other.IsAll) return true;
                if (IsAll || other.IsAll) return false;
                return Source == other.Source && Id == other.Id;
            }

            public override bool Equals(object obj) => obj is ViewKey other && Equals(other);

            public override int GetHashCode()
            {
                if (IsAll) return 0;
                unchecked
                {
                    int hash = 17;
                    hash = (hash * 23) + Source.GetHashCode();
                    hash = (hash * 23) + Id.GetHashCode();
                    return hash;
                }
            }

            public override string ToString() => IsAll ? "All" : $"{Source}:{Id:N}";
        }

        private sealed class ViewDefinition
        {
            public ViewSource Source { get; set; }
            public Guid Id { get; set; }
            public string Name { get; set; }
            public string EntityLogicalName { get; set; }
            public string FetchXml { get; set; }

            public ViewKey Key => Id == Guid.Empty ? ViewKey.All : new ViewKey(Source, Id);

            public string DisplayName
                => Source == ViewSource.System ? $"[System] {Name}" : $"[Personal] {Name}";
        }

        private sealed class ViewComboItem
        {
            public ViewKey Key { get; }
            public string Text { get; }

            public ViewComboItem(ViewKey key, string text)
            {
                Key = key;
                Text = text;
            }

            public override string ToString() => Text;

            public static ViewComboItem All => new ViewComboItem(ViewKey.All, "All");
        }

        private string _currentViewEntityLogicalName = null; // "systemuser" or "team"
        private ViewKey _selectedSystemUserView = ViewKey.All;
        private ViewKey _selectedTeamView = ViewKey.All;

        // Cached entity metadata items for Privilege Discovery entity dropdown
        private List<EntityMetadataItem> _cachedEntityMetadataItems;

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
            tsbPrivilegeDiscovery.Click += (s, e) => ExecuteMethod(OpenPrivilegeDiscoveryForSelectedRow);
            tsbAddUserToTeam.Click += (s, e) => ExecuteMethod(OpenAddUserToTeamDialog);
            tsbDel.Click += (s, e) => ExecuteMethod(DeleteSelectedAssignments);

            tsmiExportCsv.Click += (s, e) => ExportCsv();
            tsmiExportExcel.Click += (s, e) => ExportExcel();

            // Views filter (reload on selection)
            tscView.SelectedIndexChanged += (s, e) =>
            {
                if (_updatingFilters) return;

                PersistSelectedViewKeyFromUI();

                if (Service == null) return;

                if (_currentMode == ModeSystemUser)
                    ExecuteMethod(LoadUsersAndRoles);
                else if (_currentMode == ModeTeam)
                    ExecuteMethod(LoadOwnerTeamsRoles);
            };

            // Filters/search
            tscBusinessUnit.SelectedIndexChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };
            tscTeam.SelectedIndexChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };
            tscAssignment.SelectedIndexChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };
            tscTeamBU.SelectedIndexChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };
            tscRoleBU.SelectedIndexChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };
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
            tsbAddUserToTeam.Enabled = connected;
            tsbDel.Enabled = connected;
            tsddExport.Enabled = connected;
            tslView.Enabled = connected;
            tscView.Enabled = connected;
        }

        private void UpdateActionButtonsEnabledState()
        {
            var hasData = _table != null && _table.Rows.Count > 0;

            tsbSelectAllRecords.Enabled = hasData;
            if (!hasData)
                SetSelectAllRecordsChecked(false);

            tsbDel.Enabled = hasData && (_currentMode == ModeSystemUser || _currentMode == ModeTeam);
            tsbAddUserRole.Enabled = true; // always available, it reloads user view after add
            tsbAddUserToTeam.Enabled = true; // always available
        }

        private void InitializeStaticFilters()
        {
            _updatingFilters = true;
            try
            {
                tscView.Items.Clear();
                tscView.Items.Add(ViewComboItem.All);
                tscView.SelectedIndex = 0;

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

                tscTeamBU.Items.Clear();
                tscTeamBU.Items.Add("All");
                tscTeamBU.SelectedIndex = 0;

                tscRoleBU.Items.Clear();
                tscRoleBU.Items.Add("All");
                tscRoleBU.SelectedIndex = 0;

                tstSearch.Text = "";
            }
            finally
            {
                _updatingFilters = false;
            }
        }

        
        // -----------------------
        // Views (SavedQuery/UserQuery) dropdown
        // -----------------------
        private void PersistSelectedViewKeyFromUI()
        {
            var item = tscView.SelectedItem as ViewComboItem ?? ViewComboItem.All;
            var key = item.Key;

            // Prefer the currently loaded view entity; fallback to current mode.
            var entity = _currentViewEntityLogicalName;
            if (string.IsNullOrWhiteSpace(entity))
                entity = _currentMode;

            if (string.Equals(entity, ModeSystemUser, StringComparison.OrdinalIgnoreCase))
            {
                _selectedSystemUserView = key;
            }
            else if (string.Equals(entity, ModeTeam, StringComparison.OrdinalIgnoreCase))
            {
                _selectedTeamView = key;
            }
            else
            {
                // Default: store as systemuser view
                _selectedSystemUserView = key;
            }
        }

        private ViewKey GetPersistedViewKeyForEntity(string entityLogicalName)
        {
            if (string.Equals(entityLogicalName, ModeSystemUser, StringComparison.OrdinalIgnoreCase))
                return _selectedSystemUserView;

            if (string.Equals(entityLogicalName, ModeTeam, StringComparison.OrdinalIgnoreCase))
                return _selectedTeamView;

            return ViewKey.All;
        }

        private void PopulateViewsDropdown(string entityLogicalName, List<ViewDefinition> views)
        {
            _currentViewEntityLogicalName = entityLogicalName;

            _updatingFilters = true;
            try
            {
                tscView.Items.Clear();
                tscView.Items.Add(ViewComboItem.All);

                if (views != null && views.Count > 0)
                {
                    foreach (var v in views
                        .Where(v => v != null && !string.IsNullOrWhiteSpace(v.Name))
                        .OrderBy(v => v.Source)
                        .ThenBy(v => v.Name, StringComparer.OrdinalIgnoreCase))
                    {
                        tscView.Items.Add(new ViewComboItem(v.Key, v.DisplayName));
                    }
                }

                // restore selection (if present)
                var wanted = GetPersistedViewKeyForEntity(entityLogicalName);
                var match = tscView.Items
                    .OfType<ViewComboItem>()
                    .FirstOrDefault(i => i.Key.Equals(wanted));

                if (match != null)
                    tscView.SelectedItem = match;
                else
                    tscView.SelectedIndex = 0;
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

            CloseAllPrivilegeTabs();
            _cachedEntityMetadataItems = null;

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

            var selectedViewKey = _selectedSystemUserView;

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading users and roles...",
                Work = (worker, args) =>
                {
                    worker.ReportProgress(0, "Retrieving SystemUser views...");
                    var views = RetrieveViewsForEntity(Service, "systemuser");

                    IReadOnlyCollection<Guid> userIdsFilter = null;

                    if (!selectedViewKey.IsAll)
                    {
                        var selectedView = views.FirstOrDefault(v => v.Key.Equals(selectedViewKey));
                        if (selectedView != null)
                        {
                            worker.ReportProgress(0, $"Executing view: {selectedView.Name}...");
                            userIdsFilter = RetrieveRecordIdsFromView(Service, "systemuser", selectedView);
                        }
                    }

                    worker.ReportProgress(0, "Retrieving users...");
                    var users = RetrieveAllUsers(Service, userIdsFilter);

                    worker.ReportProgress(0, "Retrieving direct user roles...");
                    var directRoles = RetrieveDirectUserRoles(Service, userIdsFilter);

                    worker.ReportProgress(0, "Retrieving roles via Owner Teams...");
                    var teamRoles = RetrieveTeamUserRoles(Service, userIdsFilter);

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
                        Teams = allTeams,
                        Views = views,
                        ViewsEntityLogicalName = "systemuser"
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

                    // Views dropdown (SystemUser)
                    PopulateViewsDropdown(result.ViewsEntityLogicalName, result.Views);

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

            var selectedViewKey = _selectedTeamView;

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading Owner Teams and their roles...",
                Work = (worker, args) =>
                {
                    worker.ReportProgress(0, "Retrieving Team views...");
                    var views = RetrieveViewsForEntity(Service, "team");

                    IReadOnlyCollection<Guid> teamIdsFilter = null;

                    if (!selectedViewKey.IsAll)
                    {
                        var selectedView = views.FirstOrDefault(v => v.Key.Equals(selectedViewKey));
                        if (selectedView != null)
                        {
                            worker.ReportProgress(0, $"Executing view: {selectedView.Name}...");
                            teamIdsFilter = RetrieveRecordIdsFromView(Service, "team", selectedView);
                        }
                    }

                    worker.ReportProgress(0, "Retrieving Owner Teams...");
                    var teams = RetrieveOwnerTeams(Service, teamIdsFilter);

                    worker.ReportProgress(0, "Retrieving Owner Team roles...");
                    var teamRoles = RetrieveOwnerTeamRoles(Service, teamIdsFilter);

                    worker.ReportProgress(0, "Building rows...");
                    var rows = BuildOwnerTeamRows(teams, teamRoles);

                    worker.ReportProgress(0, "Retrieving all Business Units...");
                    var allBus = RetrieveAllBusinessUnitNames(Service);

                    worker.ReportProgress(0, "Retrieving all Owner Teams...");
                    var allTeams = RetrieveAllOwnerTeamNames(Service);

                    args.Result = new LoadResult
                    {
                        Rows = rows,
                        BusinessUnits = allBus,
                        Teams = allTeams,
                        Views = views,
                        ViewsEntityLogicalName = "team"
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

                    // Views dropdown (Team)
                    PopulateViewsDropdown(result.ViewsEntityLogicalName, result.Views);

                    _allBusinessUnitNames = result.BusinessUnits ?? new List<string>();
                    _allTeamNames = result.Teams ?? new List<string>();

                    BindRows(result.Rows ?? new List<MatrixRow>());
                    PopulateDropdownFiltersFromTable();
                    ApplyAllFilters();

                    UpdateActionButtonsEnabledState();
                }
            });
        }

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
            var teamRemovalOps = new List<TeamMemberRemoveOp>();
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
                    if (string.Equals(assignmentType, "Direct", StringComparison.OrdinalIgnoreCase))
                    {
                        if (userId == Guid.Empty || roleId == Guid.Empty)
                        {
                            skippedMissingIds++;
                            continue;
                        }

                        ops.Add(DeleteOp.ForUserRole(userId, roleId));
                    }
                    else if (string.Equals(assignmentType, "Team", StringComparison.OrdinalIgnoreCase))
                    {
                        if (userId == Guid.Empty || teamId == Guid.Empty)
                        {
                            skippedMissingIds++;
                            continue;
                        }

                        var teamName = row.Field<string>("Team") ?? "";
                        teamRemovalOps.Add(new TeamMemberRemoveOp
                        {
                            UserId = userId,
                            TeamId = teamId,
                            TeamName = teamName,
                            UserName = row.Field<string>("User") ?? ""
                        });
                    }
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

            // Deduplicate team removal ops (same user + team)
            teamRemovalOps = teamRemovalOps
                .GroupBy(o => $"{o.UserId:N}|{o.TeamId:N}")
                .Select(g => g.First())
                .ToList();

            // If there are team-inherited role removals, ask for confirmation
            if (teamRemovalOps.Count > 0)
            {
                var msg = "The following selected records are roles inherited via Owner Teams.\n" +
                          "Removing a user from an Owner Team will remove ALL roles inherited from that team.\n\n";

                foreach (var tr in teamRemovalOps.Take(10))
                    msg += $"  • {tr.UserName} → {tr.TeamName}\n";

                if (teamRemovalOps.Count > 10)
                    msg += $"  ... and {teamRemovalOps.Count - 10} more.\n";

                msg += "\nDo you want to proceed?";

                var confirm = MessageBox.Show(msg, "Remove User from Owner Team",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

                if (confirm != DialogResult.OK)
                    teamRemovalOps.Clear();
            }

            // Deduplicate operations
            ops = ops
                .GroupBy(o => o.Key)
                .Select(g => g.First())
                .ToList();

            if (ops.Count == 0 && teamRemovalOps.Count == 0)
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
                    int teamRemoveSuccess = 0;

                    // Process direct role disassociations
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

                    // Process team membership removals (RemoveMembersTeamRequest)
                    foreach (var tmOp in teamRemovalOps)
                    {
                        try
                        {
                            var removeReq = new RemoveMembersTeamRequest
                            {
                                TeamId = tmOp.TeamId,
                                MemberIds = new[] { tmOp.UserId }
                            };
                            Service.Execute(removeReq);
                            teamRemoveSuccess++;
                        }
                        catch (Exception ex2)
                        {
                            errors.Add($"Remove {tmOp.UserName} from {tmOp.TeamName}: {ex2.Message}");
                        }
                    }

                    e.Result = new DeleteResult
                    {
                        Total = ops.Count + teamRemovalOps.Count,
                        Success = success + teamRemoveSuccess,
                        Errors = errors,
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

            public int SkippedMissingIds { get; set; }
        }

        private sealed class TeamMemberRemoveOp
        {
            public Guid UserId { get; set; }
            public Guid TeamId { get; set; }
            public string TeamName { get; set; }
            public string UserName { get; set; }
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
                _cbUser.SelectedIndexChanged += CbUser_SelectedIndexChanged;

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

            private void CbUser_SelectedIndexChanged(object sender, EventArgs e)
            {
                if (_enableAcrossBuEnabled) return;

                var userItem = _cbUser.SelectedItem as ComboItem;
                if (userItem == null || userItem.Id == Guid.Empty) return;

                var user = _users.FirstOrDefault(u => u.UserId == userItem.Id);
                if (user == null || user.BusinessUnitId == Guid.Empty) return;

                _cbBusinessUnit.Items.Clear();
                _cbBusinessUnit.Items.Add(new ComboItem(user.BusinessUnitId, user.BusinessUnitName ?? ""));
                _cbBusinessUnit.SelectedIndex = 0;
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

        // -----------------------
        // Feature: Add User to Owner Team
        // -----------------------
        private void OpenAddUserToTeamDialog()
        {
            if (Service == null)
            {
                MessageBox.Show("Not connected.", "Add User to Owner Team", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading users and owner teams...",
                Work = (w, e) =>
                {
                    var users = RetrieveAllUsers(Service).Values
                        .Where(u => !u.IsDisabled)
                        .OrderBy(u => u.FullName, StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    var teams = RetrieveOwnerTeams(Service).Values
                        .OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    e.Result = new AddUserToTeamDialogData
                    {
                        Users = users,
                        Teams = teams
                    };
                },
                PostWorkCallBack = e =>
                {
                    if (e.Error != null)
                    {
                        ShowErrorDialog(e.Error);
                        return;
                    }

                    var data = (AddUserToTeamDialogData)e.Result;

                    using (var dlg = new AddUserToTeamForm(Service, data.Users, data.Teams))
                    {
                        var result = dlg.ShowDialog(this);
                        if (result == DialogResult.OK)
                        {
                            if (_currentMode == ModeSystemUser)
                                LoadUsersAndRoles();
                            else if (_currentMode == ModeTeam)
                                LoadOwnerTeamsRoles();
                        }
                    }
                }
            });
        }

        private sealed class AddUserToTeamDialogData
        {
            public List<UserInfo> Users { get; set; }
            public List<OwnerTeamInfo> Teams { get; set; }
        }

        private sealed class AddUserToTeamForm : Form
        {
            private readonly IOrganizationService _service;
            private readonly List<UserInfo> _users;
            private readonly List<OwnerTeamInfo> _teams;

            private readonly ComboBox _cbUser = new ComboBox();
            private readonly ComboBox _cbTeam = new ComboBox();
            private readonly Button _btnAdd = new Button();
            private readonly Button _btnCancel = new Button();

            public AddUserToTeamForm(
                IOrganizationService service,
                List<UserInfo> users,
                List<OwnerTeamInfo> teams)
            {
                _service = service ?? throw new ArgumentNullException(nameof(service));
                _users = users ?? new List<UserInfo>();
                _teams = teams ?? new List<OwnerTeamInfo>();

                Text = "Add User to Owner Team";
                StartPosition = FormStartPosition.CenterParent;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                ShowInTaskbar = false;
                Width = 560;
                Height = 210;

                BuildUi();
                LoadData();
            }

            private void BuildUi()
            {
                var lblUser = new Label { Text = "User:", Left = 16, Top = 20, Width = 120 };
                var lblTeam = new Label { Text = "Owner Team:", Left = 16, Top = 70, Width = 120 };

                ConfigureCombo(_cbUser);
                _cbUser.Left = 150;
                _cbUser.Top = 16;
                _cbUser.Width = 370;

                ConfigureCombo(_cbTeam);
                _cbTeam.Left = 150;
                _cbTeam.Top = 66;
                _cbTeam.Width = 370;

                _btnAdd.Text = "Add";
                _btnAdd.Left = 330;
                _btnAdd.Top = 120;
                _btnAdd.Width = 90;
                _btnAdd.Click += (s, e) => AddClicked();

                _btnCancel.Text = "Cancel";
                _btnCancel.Left = 430;
                _btnCancel.Top = 120;
                _btnCancel.Width = 90;
                _btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };

                Controls.Add(lblUser);
                Controls.Add(_cbUser);
                Controls.Add(lblTeam);
                Controls.Add(_cbTeam);
                Controls.Add(_btnAdd);
                Controls.Add(_btnCancel);
            }

            private static void ConfigureCombo(ComboBox cb)
            {
                cb.DropDownStyle = ComboBoxStyle.DropDown;
                cb.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                cb.AutoCompleteSource = AutoCompleteSource.ListItems;
            }

            private void LoadData()
            {
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

                _cbTeam.Items.Clear();
                foreach (var t in _teams)
                {
                    var buSuffix = string.IsNullOrWhiteSpace(t.BusinessUnitName)
                        ? string.Empty
                        : $" - {t.BusinessUnitName}";

                    _cbTeam.Items.Add(new ComboItem(t.TeamId, t.Name + buSuffix));
                }
            }

            private void AddClicked()
            {
                TrySelectByText(_cbUser);
                TrySelectByText(_cbTeam);

                var userItem = _cbUser.SelectedItem as ComboItem;
                var teamItem = _cbTeam.SelectedItem as ComboItem;

                if (userItem == null || userItem.Id == Guid.Empty ||
                    teamItem == null || teamItem.Id == Guid.Empty)
                {
                    MessageBox.Show("Please select a user and an owner team.", "Add User to Owner Team", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Exception executeException = null;

                using (var waiting = new WaitingForm("Adding user to team..."))
                {
                    System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            var req = new AddMembersTeamRequest
                            {
                                TeamId = teamItem.Id,
                                MemberIds = new[] { userItem.Id }
                            };
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
                            catch { }
                        }
                    });

                    waiting.ShowDialog(this);
                }

                if (executeException != null)
                {
                    MessageBox.Show(executeException.Message, "Add User to Owner Team", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        // -----------------------
        // Feature: Privileges Discovery
        // -----------------------

        private void OpenPrivilegeDiscoveryForSelectedRow()
        {
            if (Service == null)
            {
                MessageBox.Show("Not connected.", "Privileges Discovery", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // If current tab is already a Privilege Discovery tab, reload it with current entity selection
            if (tcMain.SelectedTab != null && tcMain.SelectedTab != tpMatrix && tcMain.SelectedTab.Tag is PrivilegeTabInfo existingInfo)
            {
                LoadPrivilegesForTab(tcMain.SelectedTab, existingInfo);
                return;
            }

            // Get selected record from main grid
            var selected = GetSelectedRecordFromGrid();
            if (selected == null)
            {
                MessageBox.Show("Please select a row from the grid.", "Privileges Discovery", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // If entity metadata not yet loaded, load it then create tab
            if (_cachedEntityMetadataItems == null)
            {
                WorkAsync(new WorkAsyncInfo
                {
                    Message = "Loading entity metadata...",
                    Work = (w, e) =>
                    {
                        e.Result = RetrieveAllEntityMetadataItems(Service);
                    },
                    PostWorkCallBack = e =>
                    {
                        if (e.Error != null)
                        {
                            ShowErrorDialog(e.Error);
                            return;
                        }

                        _cachedEntityMetadataItems = (List<EntityMetadataItem>)e.Result;
                        var tab = CreatePrivilegeTab(selected, _cachedEntityMetadataItems);
                        tcMain.TabPages.Add(tab);
                        tcMain.SelectedTab = tab;

                        // Auto-load with default entity (account)
                        var tabInfo = (PrivilegeTabInfo)tab.Tag;
                        LoadPrivilegesForTab(tab, tabInfo);
                    }
                });
            }
            else
            {
                var tab = CreatePrivilegeTab(selected, _cachedEntityMetadataItems);
                tcMain.TabPages.Add(tab);
                tcMain.SelectedTab = tab;

                var tabInfo = (PrivilegeTabInfo)tab.Tag;
                LoadPrivilegesForTab(tab, tabInfo);
            }
        }

        private SelectedRecordInfo GetSelectedRecordFromGrid()
        {
            if (dgvResults.CurrentRow == null) return null;

            var drv = dgvResults.CurrentRow.DataBoundItem as DataRowView;
            if (drv == null) return null;

            var row = drv.Row;

            if (_currentMode == ModeSystemUser)
            {
                var userId = row.Field<Guid?>("UserId") ?? Guid.Empty;
                if (userId == Guid.Empty) return null;

                var name = row.Field<string>("User") ?? "";
                return new SelectedRecordInfo { Mode = ModeSystemUser, RecordId = userId, RecordName = name };
            }

            if (_currentMode == ModeTeam)
            {
                var teamId = row.Field<Guid?>("TeamId") ?? Guid.Empty;
                if (teamId == Guid.Empty) return null;

                var name = row.Field<string>("Team") ?? "";
                return new SelectedRecordInfo { Mode = ModeTeam, RecordId = teamId, RecordName = name };
            }

            return null;
        }

        private TabPage CreatePrivilegeTab(SelectedRecordInfo selected, List<EntityMetadataItem> entityItems)
        {
            var modeLabel = selected.Mode == ModeSystemUser ? "User" : "Team";
            var tabText = $"Privileges - {modeLabel}: {selected.RecordName}";

            var tp = new TabPage(tabText) { UseVisualStyleBackColor = true };

            // Top panel with entity dropdown + buttons
            var panel = new Panel { Dock = DockStyle.Top, Height = 40 };

            var lblEntity = new Label
            {
                Text = "Entity:",
                Left = 8,
                Top = 10,
                Width = 50,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var cbEntity = new ComboBox
            {
                Left = 60,
                Top = 8,
                Width = 300,
                DropDownStyle = ComboBoxStyle.DropDown,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };

            foreach (var item in entityItems)
            {
                cbEntity.Items.Add(new EntityComboItem(item.LogicalName, item.DisplayName));
            }

            // Default selection: "account"
            var defaultItem = cbEntity.Items
                .OfType<EntityComboItem>()
                .FirstOrDefault(x => string.Equals(x.LogicalName, "account", StringComparison.OrdinalIgnoreCase));
            if (defaultItem != null)
                cbEntity.SelectedItem = defaultItem;
            else if (cbEntity.Items.Count > 0)
                cbEntity.SelectedIndex = 0;

            var btnLoad = new Button
            {
                Text = "Load",
                Left = 370,
                Top = 7,
                Width = 70,
                Height = 26
            };

            var btnClose = new Button
            {
                Text = "\u2715 Close",
                Left = 450,
                Top = 7,
                Width = 70,
                Height = 26
            };

            panel.Controls.Add(lblEntity);
            panel.Controls.Add(cbEntity);
            panel.Controls.Add(btnLoad);
            panel.Controls.Add(btnClose);

            // Grid for privilege matrix
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoGenerateColumns = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                ColumnHeadersHeight = 29,
                RowHeadersWidth = 51
            };

            grid.CellPainting += PrivilegeGrid_CellPainting;
            grid.ShowCellToolTips = true;
            grid.CellToolTipTextNeeded += PrivilegeGrid_CellToolTipTextNeeded;

            tp.Controls.Add(grid);
            tp.Controls.Add(panel);

            var tabInfo = new PrivilegeTabInfo
            {
                Mode = selected.Mode,
                RecordId = selected.RecordId,
                RecordName = selected.RecordName,
                EntityComboBox = cbEntity,
                PrivilegeGrid = grid
            };
            tp.Tag = tabInfo;

            btnLoad.Click += (s, e) =>
            {
                var info = (PrivilegeTabInfo)tp.Tag;
                LoadPrivilegesForTab(tp, info);
            };

            btnClose.Click += (s, e) =>
            {
                ClosePrivilegeTab(tp);
            };

            return tp;
        }

        private void LoadPrivilegesForTab(TabPage tab, PrivilegeTabInfo tabInfo)
        {
            // Resolve selected entity
            var entityItem = tabInfo.EntityComboBox.SelectedItem as EntityComboItem;
            if (entityItem == null)
            {
                // Try matching by text
                var text = (tabInfo.EntityComboBox.Text ?? "").Trim();
                if (text.Length > 0)
                {
                    entityItem = tabInfo.EntityComboBox.Items
                        .OfType<EntityComboItem>()
                        .FirstOrDefault(x => string.Equals(x.ToString(), text, StringComparison.OrdinalIgnoreCase)
                            || string.Equals(x.LogicalName, text, StringComparison.OrdinalIgnoreCase));
                }

                if (entityItem == null)
                {
                    MessageBox.Show("Please select an entity.", "Privileges Discovery", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                tabInfo.EntityComboBox.SelectedItem = entityItem;
            }

            var entityLogicalName = entityItem.LogicalName;
            var mode = tabInfo.Mode;
            var recordId = tabInfo.RecordId;
            var grid = tabInfo.PrivilegeGrid;

            grid.DataSource = null;

            WorkAsync(new WorkAsyncInfo
            {
                Message = $"Loading privileges for {entityLogicalName}...",
                Work = (w, e) =>
                {
                    // 1) Retrieve entity privileges metadata
                    w.ReportProgress(0, $"Retrieving metadata for {entityLogicalName}...");
                    var entityPrivileges = RetrieveEntityPrivileges(Service, entityLogicalName);

                    // 2) Retrieve roles for the user or team
                    w.ReportProgress(0, "Retrieving roles...");
                    List<RoleInfo> roles;
                    if (mode == ModeSystemUser)
                        roles = RetrieveAllRolesForUser(Service, recordId);
                    else
                        roles = RetrieveAllRolesForTeam(Service, recordId);

                    roles = NormalizeRolesToRoot(Service, roles);

                    // 3) Retrieve role-privilege mappings
                    w.ReportProgress(0, "Retrieving role privileges...");
                    var roleIds = roles.Select(r => r.RoleId).Distinct().ToList();
                    var privIds = entityPrivileges.Select(p => p.PrivilegeId).Distinct().ToList();

                    var rolePrivileges = new List<RolePrivilegeItem>();
                    if (roleIds.Count > 0 && privIds.Count > 0)
                        rolePrivileges = RetrieveRolePrivilegesForRoles(Service, roleIds, privIds);

                    e.Result = new PrivilegeDiscoveryResult
                    {
                        EntityLogicalName = entityLogicalName,
                        Roles = roles,
                        EntityPrivileges = entityPrivileges,
                        RolePrivileges = rolePrivileges
                    };
                },
                ProgressChanged = e =>
                {
                    if (e.UserState != null)
                        SetWorkingMessage(e.UserState.ToString());
                },
                PostWorkCallBack = e =>
                {
                    if (e.Error != null)
                    {
                        ShowErrorDialog(e.Error);
                        return;
                    }

                    var result = (PrivilegeDiscoveryResult)e.Result;

                    BuildAndBindPrivilegeMatrix(grid, result);
                }
            });
        }

        private void BuildAndBindPrivilegeMatrix(DataGridView grid, PrivilegeDiscoveryResult result)
        {
            var dt = new DataTable("PrivilegeMatrix");
            dt.Columns.Add("Role", typeof(string));

            // Determine which privilege types exist for this entity
            var privTypes = new[] { "Create", "Read", "Write", "Delete", "Append", "AppendTo", "Assign", "Share" };
            var entityPrivsByType = new Dictionary<string, PrivilegeMetadataItem>(StringComparer.OrdinalIgnoreCase);

            foreach (var ep in result.EntityPrivileges)
            {
                if (!string.IsNullOrWhiteSpace(ep.PrivilegeType) && !entityPrivsByType.ContainsKey(ep.PrivilegeType))
                    entityPrivsByType[ep.PrivilegeType] = ep;
            }

            var activePrivTypes = new List<string>();
            foreach (var pt in privTypes)
            {
                if (entityPrivsByType.ContainsKey(pt))
                {
                    activePrivTypes.Add(pt);
                    dt.Columns.Add(pt, typeof(int));
                }
            }

            // Build lookup: (roleId, privilegeId) -> depth
            var depthLookup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var rp in result.RolePrivileges)
            {
                var key = $"{rp.RoleId:N}|{rp.PrivilegeId:N}";
                if (!depthLookup.ContainsKey(key))
                    depthLookup[key] = rp.DepthMask;
                else
                {
                    // Keep highest depth
                    if (rp.DepthMask > depthLookup[key])
                        depthLookup[key] = rp.DepthMask;
                }
            }

            // One row per role
            foreach (var role in result.Roles.OrderBy(r => r.Name, StringComparer.OrdinalIgnoreCase))
            {
                var dr = dt.NewRow();
                dr["Role"] = role.Name ?? "";

                foreach (var pt in activePrivTypes)
                {
                    if (!entityPrivsByType.TryGetValue(pt, out var privMeta))
                    {
                        dr[pt] = "";
                        continue;
                    }

                    var lookupKey = $"{role.RoleId:N}|{privMeta.PrivilegeId:N}";
                    if (depthLookup.TryGetValue(lookupKey, out var depth) && depth >= 1)
                        dr[pt] = depth;
                    else
                        dr[pt] = 0;
                }

                dt.Rows.Add(dr);
            }

            grid.DataSource = dt;

            // Configure columns
            if (grid.Columns["Role"] != null)
            {
                grid.Columns["Role"].Width = 250;
                grid.Columns["Role"].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            }

            foreach (DataGridViewColumn col in grid.Columns)
            {
                if (!string.Equals(col.Name, "Role", StringComparison.OrdinalIgnoreCase))
                {
                    col.Width = 90;
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
            }
        }

        private static string DepthToText(int depthMask)
        {
            switch (depthMask)
            {
                case 1: return "User";
                case 2: return "Business Unit";
                case 4: return "Parent: Child Business Unit";
                case 8: return "Organization";
                default: return depthMask.ToString(CultureInfo.InvariantCulture);
            }
        }

        private static Color DepthToIconColor(int depthMask)
        {
            switch (depthMask)
            {
                case 1: return Color.FromArgb(90, 90, 90);     // Dark gray – User
                case 2: return Color.FromArgb(0, 120, 215);    // Blue – Business Unit
                case 4: return Color.FromArgb(200, 55, 55);    // Red – Parent:Child BU
                case 8: return Color.FromArgb(16, 137, 62);    // Green – Organization
                default: return Color.FromArgb(200, 55, 55);   // Red – None (no-entry)
            }
        }

        private void PrivilegeGrid_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            var grid = sender as DataGridView;
            if (grid == null || e.ColumnIndex < 0 || e.RowIndex < 0) return;

            var colName = grid.Columns[e.ColumnIndex].Name;
            if (string.Equals(colName, "Role", StringComparison.OrdinalIgnoreCase)) return;

            int depth = 0;
            if (e.Value != null && e.Value != DBNull.Value)
            {
                if (e.Value is int intVal)
                    depth = intVal;
                else
                    int.TryParse(e.Value.ToString(), out depth);
            }

            e.Paint(e.CellBounds, DataGridViewPaintParts.Background | DataGridViewPaintParts.Border | DataGridViewPaintParts.SelectionBackground);

            var g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            const int iconSize = 16;
            const int horizontalPadding = 3;
            const int verticalPadding = 2;
            int cx = e.CellBounds.Left + horizontalPadding + ((e.CellBounds.Width - (horizontalPadding * 2)) / 2);
            int cy = e.CellBounds.Top + verticalPadding + ((e.CellBounds.Height - (verticalPadding * 2)) / 2);

            var iconColor = DepthToIconColor(depth);

            switch (depth)
            {
                case 1:
                    DrawPersonIcon(g, iconColor, cx, cy, iconSize);
                    break;
                case 2:
                    DrawBusinessUnitIcon(g, iconColor, cx, cy, iconSize);
                    break;
                case 4:
                    DrawParentChildBusinessUnitIcon(g, iconColor, cx, cy, iconSize);
                    break;
                case 8:
                    DrawHierarchyIcon(g, iconColor, cx, cy, iconSize);
                    break;
                default:
                    DrawNoEntryIcon(g, iconColor, cx, cy, iconSize);
                    break;
            }

            e.Handled = true;
        }

        private static void DrawPersonIcon(Graphics g, Color color, int cx, int cy, int size)
        {
            var headSize = Math.Max(size / 3, 4);
            var bodyWidth = Math.Max((int)(size * 0.62f), 7);
            var bodyHeight = Math.Max((int)(size * 0.38f), 5);
            var headTop = cy - size / 2 + 1;
            var headRect = new Rectangle(cx - headSize / 2, headTop, headSize, headSize);
            var bodyRect = new Rectangle(cx - bodyWidth / 2, headRect.Bottom - 1, bodyWidth, bodyHeight);

            using (var pen = new Pen(color, 1.2f))
            {
                pen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                pen.LineJoin = System.Drawing.Drawing2D.LineJoin.Round;
                g.DrawEllipse(pen, headRect);
                g.DrawArc(pen, bodyRect, 200, 140);
            }
        }

        private static void DrawBusinessUnitIcon(Graphics g, Color color, int cx, int cy, int size)
        {
            DrawPersonIcon(g, color, cx, cy - 2, size - 1);
            DrawPersonIcon(g, color, cx - 6, cy + 3, size - 7);
            DrawPersonIcon(g, color, cx + 6, cy + 3, size - 7);
        }

        private static void DrawParentChildBusinessUnitIcon(Graphics g, Color color, int cx, int cy, int size)
        {
            using (var pen = new Pen(color, 1.1f))
            {
                pen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                var topAnchorY = cy - 2;
                var branchY = cy + 1;
                var leftX = cx - 5;
                var rightX = cx + 5;

                g.DrawLine(pen, cx, topAnchorY, cx, branchY);
                g.DrawLine(pen, leftX, branchY, rightX, branchY);
                g.DrawLine(pen, leftX, branchY, leftX, cy + 4);
                g.DrawLine(pen, rightX, branchY, rightX, cy + 4);
            }

            DrawPersonIcon(g, color, cx, cy - 5, size - 8);
            DrawPersonIcon(g, color, cx - 5, cy + 4, size - 9);
            DrawPersonIcon(g, color, cx + 5, cy + 4, size - 9);
        }

        private static void DrawHierarchyIcon(Graphics g, Color color, int cx, int cy, int size)
        {
            var boxW = 5;
            var boxH = 4;
            var topBox = new Rectangle(cx - boxW / 2, cy - size / 2 + 1, boxW, boxH);
            var leftBox = new Rectangle(cx - 7 - boxW / 2, cy + size / 2 - boxH - 1, boxW, boxH);
            var midBox = new Rectangle(cx - boxW / 2, cy + size / 2 - boxH - 1, boxW, boxH);
            var rightBox = new Rectangle(cx + 7 - boxW / 2, cy + size / 2 - boxH - 1, boxW, boxH);
            var branchY = cy - 1;

            using (var pen = new Pen(color, 1.1f))
            {
                pen.LineJoin = System.Drawing.Drawing2D.LineJoin.Round;
                g.DrawRectangle(pen, topBox);
                g.DrawRectangle(pen, leftBox);
                g.DrawRectangle(pen, midBox);
                g.DrawRectangle(pen, rightBox);

                g.DrawLine(pen, cx, topBox.Bottom, cx, branchY);
                g.DrawLine(pen, leftBox.X + leftBox.Width / 2, branchY, rightBox.X + rightBox.Width / 2, branchY);
                g.DrawLine(pen, leftBox.X + leftBox.Width / 2, branchY, leftBox.X + leftBox.Width / 2, leftBox.Y);
                g.DrawLine(pen, midBox.X + midBox.Width / 2, branchY, midBox.X + midBox.Width / 2, midBox.Y);
                g.DrawLine(pen, rightBox.X + rightBox.Width / 2, branchY, rightBox.X + rightBox.Width / 2, rightBox.Y);
            }
        }

        private static void DrawNoEntryIcon(Graphics g, Color color, int cx, int cy, int size)
        {
            int r = size / 2 - 2;
            var rect = new Rectangle(cx - r, cy - r, r * 2, r * 2);

            using (var pen = new Pen(color, 1.4f))
            {
                pen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                g.DrawEllipse(pen, rect);
                g.DrawLine(pen, cx - r + 2, cy - r + 2, cx + r - 2, cy + r - 2);
            }
        }

        private void PrivilegeGrid_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            var grid = sender as DataGridView;
            if (grid == null || e.ColumnIndex < 0 || e.RowIndex < 0) return;

            var colName = grid.Columns[e.ColumnIndex].Name;
            if (string.Equals(colName, "Role", StringComparison.OrdinalIgnoreCase)) return;

            int depth = 0;
            var cellValue = grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            if (cellValue != null && cellValue != DBNull.Value)
            {
                if (cellValue is int intVal)
                    depth = intVal;
                else
                    int.TryParse(cellValue.ToString(), out depth);
            }

            e.ToolTipText = depth >= 1 ? DepthToText(depth) : "None";
        }

        private void ClosePrivilegeTab(TabPage tab)
        {
            if (tab == null || tab == tpMatrix) return;

            tcMain.TabPages.Remove(tab);
            tab.Dispose();
        }

        private void CloseAllPrivilegeTabs()
        {
            var toRemove = new List<TabPage>();
            foreach (TabPage tp in tcMain.TabPages)
            {
                if (tp != tpMatrix)
                    toRemove.Add(tp);
            }

            foreach (var tp in toRemove)
            {
                tcMain.TabPages.Remove(tp);
                tp.Dispose();
            }
        }

        // -----------------------
        // Privileges Discovery: Data Retrieval
        // -----------------------
        private static List<EntityMetadataItem> RetrieveAllEntityMetadataItems(IOrganizationService svc)
        {
            var req = new RetrieveAllEntitiesRequest
            {
                EntityFilters = EntityFilters.Entity,
                RetrieveAsIfPublished = true
            };

            var resp = (RetrieveAllEntitiesResponse)svc.Execute(req);

            var list = new List<EntityMetadataItem>();
            foreach (var em in resp.EntityMetadata)
            {
                if (em.IsPrivate == true || em.IsIntersect == true)
                    continue;

                var displayName = em.DisplayName?.UserLocalizedLabel?.Label;
                if (string.IsNullOrWhiteSpace(displayName))
                    displayName = em.LogicalName;

                list.Add(new EntityMetadataItem
                {
                    LogicalName = em.LogicalName,
                    DisplayName = displayName
                });
            }

            return list.OrderBy(x => x.DisplayName, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static List<PrivilegeMetadataItem> RetrieveEntityPrivileges(IOrganizationService svc, string entityLogicalName)
        {
            var req = new RetrieveEntityRequest
            {
                LogicalName = entityLogicalName,
                EntityFilters = EntityFilters.Privileges,
                RetrieveAsIfPublished = true
            };

            var resp = (RetrieveEntityResponse)svc.Execute(req);
            var privileges = resp.EntityMetadata.Privileges;

            if (privileges == null || privileges.Length == 0)
                return new List<PrivilegeMetadataItem>();

            var list = new List<PrivilegeMetadataItem>();
            foreach (var sp in privileges)
            {
                var privType = sp.PrivilegeType.ToString();

                list.Add(new PrivilegeMetadataItem
                {
                    PrivilegeId = sp.PrivilegeId,
                    Name = sp.Name,
                    PrivilegeType = privType
                });
            }

            return list;
        }

        private static List<RoleInfo> RetrieveAllRolesForUser(IOrganizationService svc, Guid userId)
        {
            var roles = new Dictionary<Guid, RoleInfo>();

            // 1) Direct roles via systemuserroles
            var qeDirect = new QueryExpression("systemuserroles")
            {
                ColumnSet = new ColumnSet("systemuserid", "roleid"),
                NoLock = true
            };
            qeDirect.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, userId);

            var roleLink = qeDirect.AddLink("role", "roleid", "roleid", JoinOperator.Inner);
            roleLink.EntityAlias = "role";
            roleLink.Columns = new ColumnSet("name", "businessunitid", "parentroleid");

            var roleBuLink = roleLink.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            roleBuLink.EntityAlias = "rolebu";
            roleBuLink.Columns = new ColumnSet("name");

            var directEc = RetrieveAll(svc, qeDirect);
            foreach (var e in directEc.Entities)
            {
                var roleId = GetIdFromAttribute(e, "roleid");
                if (roleId == Guid.Empty) continue;

                if (!roles.ContainsKey(roleId))
                {
                    roles[roleId] = new RoleInfo
                    {
                        RoleId = roleId,
                        Name = GetAliasedString(e, "role", "name"),
                        BusinessUnitName = GetAliasedString(e, "rolebu", "name"),
                        ParentRoleId = GetAliasedEntityRefId(e, "role", "parentroleid")
                    };
                }
            }

            // 2) Team roles via teammembership -> teamroles
            var qeMember = new QueryExpression("teammembership")
            {
                ColumnSet = new ColumnSet("teamid"),
                NoLock = true
            };
            qeMember.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, userId);

            var teamLink = qeMember.AddLink("team", "teamid", "teamid", JoinOperator.Inner);
            teamLink.EntityAlias = "team";
            teamLink.Columns = new ColumnSet("teamtype");
            teamLink.LinkCriteria.AddCondition("teamtype", ConditionOperator.Equal, 0); // Owner Teams

            var memberEc = RetrieveAll(svc, qeMember);

            var teamIds = memberEc.Entities
                .Select(e => GetIdFromAttribute(e, "teamid"))
                .Where(id => id != Guid.Empty)
                .Distinct()
                .ToList();

            if (teamIds.Count > 0)
            {
                const int chunkSize = 500;
                foreach (var chunk in ChunkGuids(teamIds, chunkSize))
                {
                    var qeTeamRoles = new QueryExpression("teamroles")
                    {
                        ColumnSet = new ColumnSet("teamid", "roleid"),
                        NoLock = true
                    };
                    qeTeamRoles.Criteria.AddCondition("teamid", ConditionOperator.In, chunk.Cast<object>().ToArray());

                    var trRoleLink = qeTeamRoles.AddLink("role", "roleid", "roleid", JoinOperator.Inner);
                    trRoleLink.EntityAlias = "role";
                    trRoleLink.Columns = new ColumnSet("name", "businessunitid", "parentroleid");

                    var trRoleBuLink = trRoleLink.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
                    trRoleBuLink.EntityAlias = "rolebu";
                    trRoleBuLink.Columns = new ColumnSet("name");

                    var teamRoleEc = RetrieveAll(svc, qeTeamRoles);
                    foreach (var e in teamRoleEc.Entities)
                    {
                        var roleId = GetIdFromAttribute(e, "roleid");
                        if (roleId == Guid.Empty) continue;

                        if (!roles.ContainsKey(roleId))
                        {
                            roles[roleId] = new RoleInfo
                            {
                                RoleId = roleId,
                                Name = GetAliasedString(e, "role", "name"),
                                BusinessUnitName = GetAliasedString(e, "rolebu", "name"),
                                ParentRoleId = GetAliasedEntityRefId(e, "role", "parentroleid")
                            };
                        }
                    }
                }
            }

            return roles.Values.ToList();
        }

        private static List<RoleInfo> RetrieveAllRolesForTeam(IOrganizationService svc, Guid teamId)
        {
            var roles = new Dictionary<Guid, RoleInfo>();

            var qe = new QueryExpression("teamroles")
            {
                ColumnSet = new ColumnSet("teamid", "roleid"),
                NoLock = true
            };
            qe.Criteria.AddCondition("teamid", ConditionOperator.Equal, teamId);

            var roleLink = qe.AddLink("role", "roleid", "roleid", JoinOperator.Inner);
            roleLink.EntityAlias = "role";
            roleLink.Columns = new ColumnSet("name", "businessunitid", "parentroleid");

            var roleBuLink = roleLink.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            roleBuLink.EntityAlias = "rolebu";
            roleBuLink.Columns = new ColumnSet("name");

            var ec = RetrieveAll(svc, qe);
            foreach (var e in ec.Entities)
            {
                var roleId = GetIdFromAttribute(e, "roleid");
                if (roleId == Guid.Empty) continue;

                if (!roles.ContainsKey(roleId))
                {
                    roles[roleId] = new RoleInfo
                    {
                        RoleId = roleId,
                        Name = GetAliasedString(e, "role", "name"),
                        BusinessUnitName = GetAliasedString(e, "rolebu", "name"),
                        ParentRoleId = GetAliasedEntityRefId(e, "role", "parentroleid")
                    };
                }
            }

            return roles.Values.ToList();
        }

        private static List<RoleInfo> NormalizeRolesToRoot(IOrganizationService svc, List<RoleInfo> roles)
        {
            if (roles == null || roles.Count == 0)
                return new List<RoleInfo>();

            var rootIds = ResolveRootRoleIds(svc, roles
                .Select(r => r.ParentRoleId != Guid.Empty ? r.ParentRoleId : r.RoleId)
                .Where(id => id != Guid.Empty)
                .Distinct()
                .ToList());

            var roots = new Dictionary<Guid, RoleInfo>();

            foreach (var chunk in ChunkGuids(rootIds, 500))
            {
                var qe = new QueryExpression("role")
                {
                    ColumnSet = new ColumnSet("roleid", "name", "businessunitid", "parentroleid"),
                    NoLock = true
                };
                qe.Criteria.AddCondition("roleid", ConditionOperator.In, chunk.Cast<object>().ToArray());

                var roleBuLink = qe.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
                roleBuLink.EntityAlias = "rolebu";
                roleBuLink.Columns = new ColumnSet("name");

                var ec = RetrieveAll(svc, qe);
                foreach (var e in ec.Entities)
                {
                    roots[e.Id] = new RoleInfo
                    {
                        RoleId = e.Id,
                        Name = e.GetAttributeValue<string>("name"),
                        BusinessUnitName = GetAliasedString(e, "rolebu", "name"),
                        ParentRoleId = e.GetAttributeValue<EntityReference>("parentroleid")?.Id ?? Guid.Empty
                    };
                }
            }

            return roles
                .GroupBy(r => r.ParentRoleId != Guid.Empty ? r.ParentRoleId : r.RoleId)
                .Select(g =>
                {
                    if (roots.TryGetValue(g.Key, out var root))
                        return root;

                    var sample = g.First();
                    return new RoleInfo
                    {
                        RoleId = g.Key != Guid.Empty ? g.Key : sample.RoleId,
                        Name = sample.Name,
                        BusinessUnitName = sample.BusinessUnitName,
                        ParentRoleId = Guid.Empty
                    };
                })
                .ToList();
        }

        private static List<Guid> ResolveRootRoleIds(IOrganizationService svc, List<Guid> roleIds)
        {
            if (roleIds == null || roleIds.Count == 0)
                return new List<Guid>();

            var resolved = roleIds
                .Where(id => id != Guid.Empty)
                .Distinct()
                .ToDictionary(id => id, id => id);

            var pending = resolved.Keys.ToList();
            var safety = 0;

            while (pending.Count > 0 && safety < 20)
            {
                safety++;

                var nextPending = new List<Guid>();

                foreach (var chunk in ChunkGuids(pending, 500))
                {
                    var qe = new QueryExpression("role")
                    {
                        ColumnSet = new ColumnSet("roleid", "parentroleid"),
                        NoLock = true
                    };
                    qe.Criteria.AddCondition("roleid", ConditionOperator.In, chunk.Cast<object>().ToArray());

                    var ec = RetrieveAll(svc, qe);
                    foreach (var e in ec.Entities)
                    {
                        var roleId = e.Id;
                        var parentId = e.GetAttributeValue<EntityReference>("parentroleid")?.Id ?? Guid.Empty;

                        if (parentId != Guid.Empty && parentId != roleId)
                        {
                            resolved[roleId] = parentId;
                            nextPending.Add(parentId);
                        }
                    }
                }

                pending = nextPending
                    .Where(id => id != Guid.Empty)
                    .Distinct()
                    .Except(resolved.Where(kvp => kvp.Key == kvp.Value).Select(kvp => kvp.Key))
                    .ToList();
            }

            var finalMap = new Dictionary<Guid, Guid>();
            foreach (var originalId in roleIds.Where(id => id != Guid.Empty).Distinct())
            {
                var current = originalId;
                var seen = new HashSet<Guid>();

                while (resolved.TryGetValue(current, out var parentId)
                    && parentId != Guid.Empty
                    && parentId != current
                    && seen.Add(current))
                {
                    current = parentId;
                }

                finalMap[originalId] = current;
            }

            return finalMap.Values
                .Where(id => id != Guid.Empty)
                .Distinct()
                .ToList();
        }

        private static List<RolePrivilegeItem> RetrieveRolePrivilegesForRoles(
            IOrganizationService svc,
            List<Guid> roleIds,
            List<Guid> privilegeIds)
        {
            var results = new List<RolePrivilegeItem>();
            if (roleIds == null || roleIds.Count == 0 || privilegeIds == null || privilegeIds.Count == 0)
                return results;

            const int chunkSize = 500;

            // Use FetchExpression to query the roleprivileges intersect entity.
            // QueryExpression (direct or via link entity) may not return the
            // privilegedepthmask column from intersect entities in all Dataverse
            // environments. FetchXml reliably exposes all intersect attributes.
            results = FetchRolePrivileges(svc, roleIds, privilegeIds, chunkSize);

            // Fallback: for roles with no results, resolve to parent (root BU)
            // role IDs and retry. In Dataverse, child BU role copies share the
            // same privilege configuration as their root role.
            var rolesWithResults = new HashSet<Guid>(results.Select(r => r.RoleId));
            var rolesWithoutResults = roleIds
                .Where(id => id != Guid.Empty && !rolesWithResults.Contains(id))
                .Distinct()
                .ToList();

            if (rolesWithoutResults.Count > 0)
            {
                var childToParent = new Dictionary<Guid, Guid>();
                foreach (var chunk in ChunkGuids(rolesWithoutResults, chunkSize))
                {
                    var qeRole = new QueryExpression("role")
                    {
                        ColumnSet = new ColumnSet("roleid", "parentroleid"),
                        NoLock = true
                    };
                    qeRole.Criteria.AddCondition("roleid", ConditionOperator.In, chunk.Cast<object>().ToArray());

                    var roleEc = RetrieveAll(svc, qeRole);
                    foreach (var e in roleEc.Entities)
                    {
                        var parentRef = e.GetAttributeValue<EntityReference>("parentroleid");
                        if (parentRef != null && parentRef.Id != Guid.Empty && parentRef.Id != e.Id)
                            childToParent[e.Id] = parentRef.Id;
                    }
                }

                if (childToParent.Count > 0)
                {
                    var parentIds = childToParent.Values.Distinct().ToList();
                    var parentResults = FetchRolePrivileges(svc, parentIds, privilegeIds, chunkSize);

                    // Build reverse map: parentId -> list of childIds
                    var parentToChildren = new Dictionary<Guid, List<Guid>>();
                    foreach (var kvp in childToParent)
                    {
                        if (!parentToChildren.TryGetValue(kvp.Value, out var list))
                        {
                            list = new List<Guid>();
                            parentToChildren[kvp.Value] = list;
                        }
                        list.Add(kvp.Key);
                    }

                    foreach (var pr in parentResults)
                    {
                        if (parentToChildren.TryGetValue(pr.RoleId, out var children))
                        {
                            foreach (var childId in children)
                            {
                                results.Add(new RolePrivilegeItem
                                {
                                    RoleId = childId,
                                    PrivilegeId = pr.PrivilegeId,
                                    DepthMask = pr.DepthMask
                                });
                            }
                        }
                    }
                }
            }

            return results;
        }

        private static List<RolePrivilegeItem> FetchRolePrivileges(
            IOrganizationService svc,
            List<Guid> roleIds,
            List<Guid> privilegeIds,
            int chunkSize)
        {
            var results = new List<RolePrivilegeItem>();

            // Use the standard N:N intersect query pattern in FetchXml.
            // Query 'privilege' as main entity with 'roleprivileges' as an
            // intersect link-entity (intersect='true') and 'role' as a
            // nested filter. This reliably returns intersect attributes
            // like privilegedepthmask across all Dataverse environments,
            // including child-BU role copies.
            foreach (var roleChunk in ChunkGuids(roleIds, chunkSize))
            {
                foreach (var privChunk in ChunkGuids(privilegeIds, chunkSize))
                {
                    var roleValues = string.Join("", roleChunk.Select(id => $"<value>{id}</value>"));
                    var privValues = string.Join("", privChunk.Select(id => $"<value>{id}</value>"));

                    var fetchXml =
                        "<fetch no-lock='true'>" +
                        "  <entity name='privilege'>" +
                        "    <attribute name='privilegeid' />" +
                        "    <filter>" +
                        "      <condition attribute='privilegeid' operator='in'>" +
                                 privValues +
                        "      </condition>" +
                        "    </filter>" +
                        "    <link-entity name='roleprivileges' from='privilegeid' to='privilegeid' visible='false' intersect='true' alias='rp'>" +
                        "      <attribute name='privilegedepthmask' />" +
                        "      <attribute name='roleid' />" +
                        "      <link-entity name='role' from='roleid' to='roleid'>" +
                        "        <filter>" +
                        "          <condition attribute='roleid' operator='in'>" +
                                     roleValues +
                        "          </condition>" +
                        "        </filter>" +
                        "      </link-entity>" +
                        "    </link-entity>" +
                        "  </entity>" +
                        "</fetch>";

                    var ec = svc.RetrieveMultiple(new FetchExpression(fetchXml));

                    foreach (var e in ec.Entities)
                    {
                        var privId = e.Id;

                        var roleId = Guid.Empty;
                        var roleAv = e.GetAttributeValue<AliasedValue>("rp.roleid");
                        if (roleAv?.Value is Guid rg) roleId = rg;

                        var depth = 0;
                        var depthAv = e.GetAttributeValue<AliasedValue>("rp.privilegedepthmask");
                        if (depthAv?.Value is int di) depth = di;
                        else if (depthAv?.Value is long dl) depth = (int)dl;
                        else if (depthAv?.Value != null) int.TryParse(depthAv.Value.ToString(), out depth);

                        if (roleId == Guid.Empty || privId == Guid.Empty) continue;

                        results.Add(new RolePrivilegeItem
                        {
                            RoleId = roleId,
                            PrivilegeId = privId,
                            DepthMask = depth
                        });
                    }
                }
            }

            return results;
        }

        // -----------------------
        // Privileges Discovery: Data Model classes
        // -----------------------
        private sealed class SelectedRecordInfo
        {
            public string Mode { get; set; }
            public Guid RecordId { get; set; }
            public string RecordName { get; set; }
        }

        private sealed class PrivilegeTabInfo
        {
            public string Mode { get; set; }
            public Guid RecordId { get; set; }
            public string RecordName { get; set; }
            public ComboBox EntityComboBox { get; set; }
            public DataGridView PrivilegeGrid { get; set; }
        }

        private sealed class EntityMetadataItem
        {
            public string LogicalName { get; set; }
            public string DisplayName { get; set; }
        }

        private sealed class EntityComboItem
        {
            public string LogicalName { get; }
            public string DisplayText { get; }

            public EntityComboItem(string logicalName, string displayName)
            {
                LogicalName = logicalName;
                DisplayText = string.IsNullOrWhiteSpace(displayName)
                    ? logicalName
                    : $"{displayName} ({logicalName})";
            }

            public override string ToString() => DisplayText;
        }

        private sealed class PrivilegeMetadataItem
        {
            public Guid PrivilegeId { get; set; }
            public string Name { get; set; }
            public string PrivilegeType { get; set; }
        }

        private sealed class RolePrivilegeItem
        {
            public Guid RoleId { get; set; }
            public Guid PrivilegeId { get; set; }
            public int DepthMask { get; set; }
        }

        private sealed class PrivilegeDiscoveryResult
        {
            public string EntityLogicalName { get; set; }
            public List<RoleInfo> Roles { get; set; }
            public List<PrivilegeMetadataItem> EntityPrivileges { get; set; }
            public List<RolePrivilegeItem> RolePrivileges { get; set; }
        }

        private sealed class LoadResult
        {
            public List<MatrixRow> Rows { get; set; }
            public List<string> BusinessUnits { get; set; }
            // added to carry team names back to UI thread
            public List<string> Teams { get; set; }

            // Views for the current entity (systemuser / team)
            public List<ViewDefinition> Views { get; set; }
            public string ViewsEntityLogicalName { get; set; }
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

                var teamBUs = _table.AsEnumerable()
                    .Select(r => r.Field<string>("Team Business Unit"))
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var roleBUs = _table.AsEnumerable()
                    .Select(r => r.Field<string>("Role Business Unit"))
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

                tscTeamBU.Items.Clear();
                tscTeamBU.Items.Add("All");
                foreach (var tbu in teamBUs) tscTeamBU.Items.Add(tbu);
                tscTeamBU.SelectedIndex = 0;

                tscRoleBU.Items.Clear();
                tscRoleBU.Items.Add("All");
                foreach (var rbu in roleBUs) tscRoleBU.Items.Add(rbu);
                tscRoleBU.SelectedIndex = 0;
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

            var teamBU = (tscTeamBU.SelectedItem?.ToString() ?? "All").Trim();
            if (!string.Equals(teamBU, "All", StringComparison.OrdinalIgnoreCase))
                filters.Add($"[Team Business Unit] = '{EscapeRowFilterValue(teamBU)}'");

            var roleBU = (tscRoleBU.SelectedItem?.ToString() ?? "All").Trim();
            if (!string.Equals(roleBU, "All", StringComparison.OrdinalIgnoreCase))
                filters.Add($"[Role Business Unit] = '{EscapeRowFilterValue(roleBU)}'");

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

        
        // -----------------------
        // Views (SavedQuery/UserQuery) + execution (FetchXml -> QueryExpression)
        // -----------------------
        private static List<ViewDefinition> RetrieveViewsForEntity(IOrganizationService svc, string entityLogicalName)
        {
            var views = new List<ViewDefinition>();

            // System views (savedquery)
            var qeSys = new QueryExpression("savedquery")
            {
                ColumnSet = new ColumnSet("savedqueryid", "name", "fetchxml", "returnedtypecode", "statecode"),
                NoLock = true
            };
            qeSys.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, entityLogicalName);
            qeSys.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); // Active
            qeSys.Orders.Add(new OrderExpression("name", OrderType.Ascending));

            var sys = RetrieveAll(svc, qeSys);

            foreach (var e in sys.Entities)
            {
                var name = e.GetAttributeValue<string>("name");
                var fetch = e.GetAttributeValue<string>("fetchxml");

                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(fetch))
                    continue;

                views.Add(new ViewDefinition
                {
                    Source = ViewSource.System,
                    Id = e.Id,
                    Name = name,
                    EntityLogicalName = entityLogicalName,
                    FetchXml = fetch
                });
            }

            // Personal views (userquery)
            var qeUser = new QueryExpression("userquery")
            {
                ColumnSet = new ColumnSet("userqueryid", "name", "fetchxml", "returnedtypecode", "statecode"),
                NoLock = true
            };
            qeUser.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, entityLogicalName);
            qeUser.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0); // Active
            qeUser.Orders.Add(new OrderExpression("name", OrderType.Ascending));

            var usr = RetrieveAll(svc, qeUser);

            foreach (var e in usr.Entities)
            {
                var name = e.GetAttributeValue<string>("name");
                var fetch = e.GetAttributeValue<string>("fetchxml");

                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(fetch))
                    continue;

                views.Add(new ViewDefinition
                {
                    Source = ViewSource.Personal,
                    Id = e.Id,
                    Name = name,
                    EntityLogicalName = entityLogicalName,
                    FetchXml = fetch
                });
            }

            return views;
        }

        private static string GetPrimaryIdAttributeName(string entityLogicalName)
        {
            if (string.Equals(entityLogicalName, "systemuser", StringComparison.OrdinalIgnoreCase))
                return "systemuserid";

            if (string.Equals(entityLogicalName, "team", StringComparison.OrdinalIgnoreCase))
                return "teamid";

            // Best-effort fallback (works for most entities)
            return entityLogicalName + "id";
        }

        private static IReadOnlyCollection<Guid> RetrieveRecordIdsFromView(IOrganizationService svc, string entityLogicalName, ViewDefinition view)
        {
            if (view == null || string.IsNullOrWhiteSpace(view.FetchXml))
                return Array.Empty<Guid>();

            // Convert FetchXML stored in the View to QueryExpression (no hardcoded FetchXML in the tool)
            var req = new FetchXmlToQueryExpressionRequest
            {
                FetchXml = view.FetchXml
            };

            var resp = (FetchXmlToQueryExpressionResponse)svc.Execute(req);
            var qe = resp.Query;

            qe.NoLock = true;
            qe.ColumnSet = new ColumnSet(GetPrimaryIdAttributeName(entityLogicalName));

            var ec = RetrieveAll(svc, qe);

            return ec.Entities
                .Select(e => e.Id)
                .Where(id => id != Guid.Empty)
                .Distinct()
                .ToList();
        }

        private static IEnumerable<Guid[]> ChunkGuids(IEnumerable<Guid> ids, int chunkSize)
        {
            if (ids == null) yield break;

            var list = ids
                .Where(id => id != Guid.Empty)
                .Distinct()
                .ToList();

            for (int i = 0; i < list.Count; i += chunkSize)
                yield return list.Skip(i).Take(chunkSize).ToArray();
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

        
        private static Dictionary<Guid, UserInfo> RetrieveAllUsers(IOrganizationService svc, IReadOnlyCollection<Guid> userIdsFilter = null)
        {
            var users = new Dictionary<Guid, UserInfo>();

            // If a view filter is provided, query users in chunks
            if (userIdsFilter != null)
            {
                var ids = userIdsFilter.Where(id => id != Guid.Empty).Distinct().ToList();
                if (ids.Count == 0)
                    return users;

                const int chunkSize = 500;

                foreach (var chunk in ChunkGuids(ids, chunkSize))
                {
                    var qe = new QueryExpression("systemuser")
                    {
                        ColumnSet = new ColumnSet("systemuserid", "fullname", "internalemailaddress", "isdisabled", "businessunitid"),
                        NoLock = true
                    };

                    qe.Criteria.AddCondition("systemuserid", ConditionOperator.In, chunk.Cast<object>().ToArray());

                    var buLink = qe.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
                    buLink.EntityAlias = "bu";
                    buLink.Columns = new ColumnSet("name");

                    var ec = RetrieveAll(svc, qe);

                    foreach (var e in ec.Entities)
                    {
                        var buRef = e.GetAttributeValue<EntityReference>("businessunitid");
                        users[e.Id] = new UserInfo
                        {
                            UserId = e.Id,
                            FullName = e.GetAttributeValue<string>("fullname"),
                            Email = e.GetAttributeValue<string>("internalemailaddress"),
                            BusinessUnitId = buRef?.Id ?? Guid.Empty,
                            BusinessUnitName = GetAliasedString(e, "bu", "name"),
                            IsDisabled = e.GetAttributeValue<bool?>("isdisabled") ?? false
                        };
                    }
                }

                return users;
            }

            // No filter => load all users
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

                foreach (var e in ec.Entities)
                {
                    var buRef = e.GetAttributeValue<EntityReference>("businessunitid");
                    users[e.Id] = new UserInfo
                    {
                        UserId = e.Id,
                        FullName = e.GetAttributeValue<string>("fullname"),
                        Email = e.GetAttributeValue<string>("internalemailaddress"),
                        BusinessUnitId = buRef?.Id ?? Guid.Empty,
                        BusinessUnitName = GetAliasedString(e, "bu", "name"),
                        IsDisabled = e.GetAttributeValue<bool?>("isdisabled") ?? false
                    };
                }

                return users;
            }
        }


        private static Dictionary<Guid, List<RoleInfo>> RetrieveDirectUserRoles(IOrganizationService svc, IReadOnlyCollection<Guid> userIdsFilter = null)
        {
            // systemuserroles: intersect entity systemuser <-> role
            var dict = new Dictionary<Guid, List<RoleInfo>>();

            // If a view filter is provided, query systemuserroles in chunks
            if (userIdsFilter != null)
            {
                var ids = userIdsFilter.Where(id => id != Guid.Empty).Distinct().ToList();
                if (ids.Count == 0)
                    return dict;

                const int chunkSize = 500;

                foreach (var chunk in ChunkGuids(ids, chunkSize))
                {
                    var qe = new QueryExpression("systemuserroles")
                    {
                        ColumnSet = new ColumnSet("systemuserid", "roleid"),
                        NoLock = true
                    };
                    qe.Criteria.AddCondition("systemuserid", ConditionOperator.In, chunk.Cast<object>().ToArray());

                    var roleLink = qe.AddLink("role", "roleid", "roleid", JoinOperator.Inner);
                    roleLink.EntityAlias = "role";
                    roleLink.Columns = new ColumnSet("name", "businessunitid");

                    var roleBuLink = roleLink.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
                    roleBuLink.EntityAlias = "rolebu";
                    roleBuLink.Columns = new ColumnSet("name");

                    var ec = RetrieveAll(svc, qe);

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

                        if (!list.Any(r => r.RoleId == role.RoleId))
                            list.Add(role);
                    }
                }

                return dict;
            }

            // No filter => load all direct assignments
            {
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

                    if (!list.Any(r => r.RoleId == role.RoleId))
                        list.Add(role);
                }

                return dict;
            }
        }


        private static Dictionary<Guid, List<TeamRoleInfo>> RetrieveTeamUserRoles(IOrganizationService svc, IReadOnlyCollection<Guid> userIdsFilter = null)
        {
            // For each user: their teams -> those teams' roles
            // We retrieve team membership via teammembership and then teamroles.

            // 1) teammembership: systemuser <-> team
            var membership = new List<(Guid UserId, Guid TeamId, string TeamName, string TeamBu, bool IsDefault)>();

            Action<QueryExpression> addTeamJoins = (qeMember) =>
            {
                // join team to filter owner teams only
                var teamLink = qeMember.AddLink("team", "teamid", "teamid", JoinOperator.Inner);
                teamLink.EntityAlias = "team";
                teamLink.Columns = new ColumnSet("name", "teamtype", "businessunitid", "isdefault");
                teamLink.LinkCriteria.AddCondition("teamtype", ConditionOperator.Equal, 0); // Owner Teams

                var teamBuLink = teamLink.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
                teamBuLink.EntityAlias = "teambu";
                teamBuLink.Columns = new ColumnSet("name");
            };

            if (userIdsFilter != null)
            {
                var ids = userIdsFilter.Where(id => id != Guid.Empty).Distinct().ToList();
                if (ids.Count == 0)
                    return new Dictionary<Guid, List<TeamRoleInfo>>();

                const int chunkSize = 500;

                foreach (var chunk in ChunkGuids(ids, chunkSize))
                {
                    var qeMember = new QueryExpression("teammembership")
                    {
                        ColumnSet = new ColumnSet("systemuserid", "teamid"),
                        NoLock = true
                    };
                    qeMember.Criteria.AddCondition("systemuserid", ConditionOperator.In, chunk.Cast<object>().ToArray());

                    addTeamJoins(qeMember);

                    var memberEc = RetrieveAll(svc, qeMember);

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
                }
            }
            else
            {
                var qeMember = new QueryExpression("teammembership")
                {
                    ColumnSet = new ColumnSet("systemuserid", "teamid"),
                    NoLock = true
                };

                addTeamJoins(qeMember);

                var memberEc = RetrieveAll(svc, qeMember);

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
            }

            // If no memberships, return empty
            var teamIds = membership.Select(m => m.TeamId).Where(id => id != Guid.Empty).Distinct().ToList();
            if (teamIds.Count == 0)
                return new Dictionary<Guid, List<TeamRoleInfo>>();

            // 2) teamroles: team <-> role (only for teams found above)
            var teamRolesMap = new Dictionary<Guid, List<TeamRoleInfo>>();
            const int teamChunkSize = 500;

            foreach (var teamChunk in ChunkGuids(teamIds, teamChunkSize))
            {
                var qeTeamRoles = new QueryExpression("teamroles")
                {
                    ColumnSet = new ColumnSet("teamid", "roleid"),
                    NoLock = true
                };
                qeTeamRoles.Criteria.AddCondition("teamid", ConditionOperator.In, teamChunk.Cast<object>().ToArray());

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
                teamLink2.LinkCriteria.AddCondition("teamtype", ConditionOperator.Equal, 0); // Owner Teams

                var teamBuLink2 = teamLink2.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
                teamBuLink2.EntityAlias = "teambu";
                teamBuLink2.Columns = new ColumnSet("name");

                var teamRoleEc = RetrieveAll(svc, qeTeamRoles);

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


        private static Dictionary<Guid, OwnerTeamInfo> RetrieveOwnerTeams(IOrganizationService svc, IReadOnlyCollection<Guid> teamIdsFilter = null)
        {
            var dict = new Dictionary<Guid, OwnerTeamInfo>();

            if (teamIdsFilter != null)
            {
                var ids = teamIdsFilter.Where(id => id != Guid.Empty).Distinct().ToList();
                if (ids.Count == 0)
                    return dict;

                const int chunkSize = 500;

                foreach (var chunk in ChunkGuids(ids, chunkSize))
                {
                    var qe = new QueryExpression("team")
                    {
                        ColumnSet = new ColumnSet("teamid", "name", "teamtype", "businessunitid"),
                        NoLock = true
                    };
                    qe.Criteria.AddCondition("teamtype", ConditionOperator.Equal, 0); // Owner Teams
                    qe.Criteria.AddCondition("teamid", ConditionOperator.In, chunk.Cast<object>().ToArray());

                    var buLink = qe.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
                    buLink.EntityAlias = "teambu";
                    buLink.Columns = new ColumnSet("name");

                    var ec = RetrieveAll(svc, qe);

                    foreach (var e in ec.Entities)
                    {
                        dict[e.Id] = new OwnerTeamInfo
                        {
                            TeamId = e.Id,
                            Name = e.GetAttributeValue<string>("name"),
                            BusinessUnitName = GetAliasedString(e, "teambu", "name")
                        };
                    }
                }

                return dict;
            }

            // No filter => load all owner teams
            {
                var qe = new QueryExpression("team")
                {
                    ColumnSet = new ColumnSet("teamid", "name", "teamtype", "businessunitid"),
                    NoLock = true
                };
                qe.Criteria.AddCondition("teamtype", ConditionOperator.Equal, 0); // Owner Teams

                var buLink = qe.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
                buLink.EntityAlias = "teambu";
                buLink.Columns = new ColumnSet("name");

                var ec = RetrieveAll(svc, qe);

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
        }


        private static List<TeamRoleInfo> RetrieveOwnerTeamRoles(IOrganizationService svc, IReadOnlyCollection<Guid> teamIdsFilter = null)
        {
            var results = new List<TeamRoleInfo>();

            Action<QueryExpression> addJoins = (qe) =>
            {
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
                roleLink.Columns = new ColumnSet("name", "businessunitid");

                var roleBuLink = roleLink.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
                roleBuLink.EntityAlias = "rolebu";
                roleBuLink.Columns = new ColumnSet("name");
            };

            if (teamIdsFilter != null)
            {
                var ids = teamIdsFilter.Where(id => id != Guid.Empty).Distinct().ToList();
                if (ids.Count == 0)
                    return results;

                const int chunkSize = 500;

                foreach (var chunk in ChunkGuids(ids, chunkSize))
                {
                    var qe = new QueryExpression("teamroles")
                    {
                        ColumnSet = new ColumnSet("teamid", "roleid"),
                        NoLock = true
                    };
                    qe.Criteria.AddCondition("teamid", ConditionOperator.In, chunk.Cast<object>().ToArray());

                    addJoins(qe);

                    var ec = RetrieveAll(svc, qe);

                    foreach (var e in ec.Entities)
                    {
                        var teamId = GetIdFromAttribute(e, "teamid");
                        var roleId = GetIdFromAttribute(e, "roleid");
                        if (teamId == Guid.Empty || roleId == Guid.Empty) continue;

                        results.Add(new TeamRoleInfo
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
                }
            }
            else
            {
                var qe = new QueryExpression("teamroles")
                {
                    ColumnSet = new ColumnSet("teamid", "roleid"),
                    NoLock = true
                };

                addJoins(qe);

                var ec = RetrieveAll(svc, qe);

                foreach (var e in ec.Entities)
                {
                    var teamId = GetIdFromAttribute(e, "teamid");
                    var roleId = GetIdFromAttribute(e, "roleid");
                    if (teamId == Guid.Empty || roleId == Guid.Empty) continue;

                    results.Add(new TeamRoleInfo
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
            }

            // De-dupe TeamId + RoleId
            return results
                .GroupBy(x => $"{x.TeamId:N}|{x.RoleId:N}")
                .Select(g => g.First())
                .ToList();
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

        private static Guid GetAliasedEntityRefId(Entity e, string alias, string attributeName)
        {
            var key = $"{alias}.{attributeName}";
            if (!e.Attributes.ContainsKey(key)) return Guid.Empty;
            var val = e.Attributes[key] as AliasedValue;
            if (val?.Value is EntityReference er) return er.Id;
            if (val?.Value is Guid g) return g;
            return Guid.Empty;
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
            public Guid BusinessUnitId { get; set; }
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
            public Guid ParentRoleId { get; set; }
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
