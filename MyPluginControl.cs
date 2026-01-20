
using ClosedXML.Excel;
using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using System.Windows.Forms;
using XrmToolBox.Extensibility;

namespace GM.XrmToolBox.UserRoleMatrix
{
    public partial class MyPluginControl : PluginControlBase
    {
        private DataTable _table;
        private DataView _view;
        private readonly BindingSource _bindingSource = new BindingSource();
        private bool _updatingFilters;

        // Column list (export + filters)
        private static readonly string[] ExportColumns = new[]
        {
            "User",
            "Email",
            "User Business Unit",
            "Assignment Type",
            "Team",
            "Team Business Unit",
            "Role",
            "Role Business Unit",
            "Duplicate"
        };

        public MyPluginControl()
        {
            InitializeComponent();

            dgvResults.AutoGenerateColumns = true;
            dgvResults.ReadOnly = true;
            dgvResults.AllowUserToAddRows = false;
            dgvResults.AllowUserToDeleteRows = false;
            dgvResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvResults.MultiSelect = false;
            dgvResults.DataSource = _bindingSource;

            // Events
            tsbLoad.Click += (s, e) => ExecuteMethod(LoadUsersAndRoles);

            tsmiExportCsv.Click += (s, e) => ExportCsv();
            tsmiExportExcel.Click += (s, e) => ExportExcel();

            tstSearch.TextChanged += (s, e) => ApplyAllFilters();

            tscBusinessUnit.SelectedIndexChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };
            tscTeam.SelectedIndexChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };
            tscAssignment.SelectedIndexChanged += (s, e) => { if (!_updatingFilters) ApplyAllFilters(); };

            dgvResults.DataBindingComplete += (s, e) =>
            {
                HideTechnicalColumns();
                ApplyDuplicateRowHighlight();
            };

            InitializeStaticFilters();
        }

        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);
            ClearResults();
        }

        private void ClearResults()
        {
            _table = null;
            _view = null;
            _bindingSource.DataSource = null;

            _updatingFilters = true;
            try
            {
                tscBusinessUnit.Items.Clear();
                tscTeam.Items.Clear();
                InitializeStaticFilters();
            }
            finally
            {
                _updatingFilters = false;
            }

            tslCount.Text = "Rows: 0";
        }

        private void InitializeStaticFilters()
        {
            _updatingFilters = true;
            try
            {
                // Assignment
                tscAssignment.Items.Clear();
                tscAssignment.Items.Add("All");
                tscAssignment.Items.Add("Direct");
                tscAssignment.Items.Add("Team");
                tscAssignment.SelectedIndex = 0;

                // BU
                tscBusinessUnit.Items.Clear();
                tscBusinessUnit.Items.Add("All");
                tscBusinessUnit.SelectedIndex = 0;

                // Team
                tscTeam.Items.Clear();
                tscTeam.Items.Add("All");
                tscTeam.SelectedIndex = 0;
            }
            finally
            {
                _updatingFilters = false;
            }
        }

        private void LoadUsersAndRoles()
        {
            tsbLoad.Enabled = false;

            WorkAsync(new WorkAsyncInfo
            {
                Message = "Loading users and roles...",
                MessageWidth = 360,
                MessageHeight = 150,

                Work = (worker, args) =>
                {
                    worker.ReportProgress(0, "Retrieving users...");
                    var users = RetrieveAllUsers(Service);

                    worker.ReportProgress(0, "Retrieving direct user roles...");
                    var directRoles = RetrieveDirectUserRoles(Service);

                    worker.ReportProgress(0, "Retrieving roles via Owner Teams...");
                    var teamRoles = RetrieveTeamUserRoles(Service);

                    worker.ReportProgress(0, "Building result rows (with duplicate detection)...");
                    var rows = BuildRows(users, directRoles, teamRoles);

                    args.Result = rows;
                },

                ProgressChanged = e =>
                {
                    if (e.UserState != null)
                        SetWorkingMessage(e.UserState.ToString());
                },

                PostWorkCallBack = e =>
                {
                    tsbLoad.Enabled = true;

                    if (e.Error != null)
                    {
                        ShowErrorDialog(e.Error);
                        return;
                    }

                    var rows = (List<UserRoleRow>)e.Result;

                    BindRows(rows);
                    PopulateDropdownFiltersFromTable();

                    ApplyAllFilters();
                }
            });
        }

        private void BindRows(List<UserRoleRow> rows)
        {
            _table = CreateSchema();

            foreach (var r in rows)
            {
                var dr = _table.NewRow();

                // Technical columns
                dr["UserId"] = r.UserId;
                dr["RoleId"] = r.RoleId;
                dr["TeamId"] = r.TeamId;

                // Display columns
                dr["User"] = r.UserFullName ?? "";
                dr["Email"] = r.UserEmail ?? "";
                dr["User Business Unit"] = r.UserBusinessUnit ?? "";
                dr["Assignment Type"] = r.AssignmentType ?? "";
                dr["Team"] = r.TeamName ?? "";
                dr["Team Business Unit"] = r.TeamBusinessUnit ?? "";
                dr["Role"] = r.RoleName ?? "";
                dr["Role Business Unit"] = r.RoleBusinessUnit ?? "";
                dr["Duplicate"] = r.Duplicate;

                _table.Rows.Add(dr);
            }

            _view = new DataView(_table);
            _bindingSource.DataSource = _view;

            HideTechnicalColumns();
            ApplyDuplicateRowHighlight();
        }

        private static DataTable CreateSchema()
        {
            var dt = new DataTable("UserRoles");

            // Technical columns (hidden in grid)
            dt.Columns.Add("UserId", typeof(Guid));
            dt.Columns.Add("RoleId", typeof(Guid));
            dt.Columns.Add("TeamId", typeof(Guid));

            // Display columns (English only)
            dt.Columns.Add("User", typeof(string));
            dt.Columns.Add("Email", typeof(string));
            dt.Columns.Add("User Business Unit", typeof(string));
            dt.Columns.Add("Assignment Type", typeof(string)); // Direct / Team / None
            dt.Columns.Add("Team", typeof(string));
            dt.Columns.Add("Team Business Unit", typeof(string));
            dt.Columns.Add("Role", typeof(string));
            dt.Columns.Add("Role Business Unit", typeof(string));
            dt.Columns.Add("Duplicate", typeof(bool));

            return dt;
        }

        private void HideTechnicalColumns()
        {
            if (dgvResults.Columns["UserId"] != null) dgvResults.Columns["UserId"].Visible = false;
            if (dgvResults.Columns["RoleId"] != null) dgvResults.Columns["RoleId"].Visible = false;
            if (dgvResults.Columns["TeamId"] != null) dgvResults.Columns["TeamId"].Visible = false;
        }

        private void ApplyDuplicateRowHighlight()
        {
            if (dgvResults.Rows.Count == 0) return;
            if (dgvResults.Columns["Duplicate"] == null) return;

            // Choose a "normal" background fallback
            var normalBack = dgvResults.RowsDefaultCellStyle.BackColor;
            if (normalBack == Color.Empty)
                normalBack = SystemColors.Window;

            foreach (DataGridViewRow row in dgvResults.Rows)
            {
                var value = row.Cells["Duplicate"].Value;
                var isDup = value is bool b && b;

                row.DefaultCellStyle.BackColor = isDup ? Color.LightYellow : normalBack;
            }
        }

        private void PopulateDropdownFiltersFromTable()
        {
            if (_table == null) return;

            _updatingFilters = true;
            try
            {
                var businessUnits = _table.AsEnumerable()
                    .Select(r => r.Field<string>("User Business Unit"))
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

            var bu = (tscBusinessUnit.SelectedItem?.ToString() ?? "All").Trim();
            if (!string.Equals(bu, "All", StringComparison.OrdinalIgnoreCase))
                filters.Add($"[User Business Unit] = '{EscapeRowFilterValue(bu)}'");

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
                    $" OR [Role Business Unit] LIKE '%{s}%'" +
                    $" OR [Assignment Type] LIKE '%{s}%'" +
                    $" OR Convert([Duplicate], 'System.String') LIKE '%{s}%'";

                filters.Add("(" + searchClause + ")");
            }

            _view.RowFilter = string.Join(" AND ", filters);

            tslCount.Text = _table != null
                ? $"Rows: {_view.Count:n0} / {_table.Rows.Count:n0}"
                : "Rows: 0";

            ApplyDuplicateRowHighlight();
        }

        private static string EscapeRowFilterValue(string value)
        {
            return (value ?? "").Replace("'", "''");
        }

        // ==========================
        // EXPORT
        // ==========================

        private DataTable GetCurrentViewForExport()
        {
            if (_view == null) return null;

            // Export ONLY display columns (same order every time)
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
                FileName = $"UserRoles_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
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
            // UTF-8 with BOM for Excel-friendliness
            using (var sw = new StreamWriter(path, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true)))
            {
                // Header
                var headers = dt.Columns.Cast<DataColumn>().Select(c => CsvEscape(c.ColumnName));
                sw.WriteLine(string.Join(",", headers));

                // Rows
                foreach (DataRow row in dt.Rows)
                {
                    var fields = dt.Columns.Cast<DataColumn>()
                        .Select(c => CsvEscape(row[c]?.ToString() ?? ""));
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
                FileName = $"UserRoles_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
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
                var ws = wb.Worksheets.Add("UserRoles");
                ws.Cell(1, 1).InsertTable(dt, "UserRolesTable", true);
                ws.Columns().AdjustToContents();
                wb.SaveAs(path);
            }
        }

        // ==========================
        // QUERYEXPRESSION RETRIEVE
        // ==========================

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
            teamLink.Columns = new ColumnSet("teamid", "name", "teamtype", "businessunitid");
            teamLink.LinkCriteria.AddCondition("teamtype", ConditionOperator.Equal, 0); // 0 = Owner Team

            var teamBuLink = teamLink.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            teamBuLink.EntityAlias = "teambu";
            teamBuLink.Columns = new ColumnSet("name");

            // Team -> TeamRoles (intersect) -> Role
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
                        RoleBusinessUnitName = GetAliasedString(e, "rolebu", "name")
                    });
                }
            }

            return map;
        }

        private static List<UserRoleRow> BuildRows(
            Dictionary<Guid, UserInfo> users,
            Dictionary<Guid, List<RoleInfo>> directRoles,
            Dictionary<Guid, List<TeamRoleInfo>> teamRoles)
        {
            // Duplicate definition: (UserId, RoleId) exists both in direct and team assignments
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

            var rows = new List<UserRoleRow>(users.Count * 2);

            foreach (var u in users.Values.OrderBy(x => x.FullName, StringComparer.OrdinalIgnoreCase))
            {
                var any = false;

                if (directRoles.TryGetValue(u.UserId, out var dr))
                {
                    foreach (var r in dr.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase))
                    {
                        var isDup = duplicates.Contains((u.UserId, r.RoleId));

                        rows.Add(new UserRoleRow
                        {
                            UserId = u.UserId,
                            RoleId = r.RoleId,
                            TeamId = Guid.Empty,

                            UserFullName = u.FullName,
                            UserEmail = u.Email,
                            UserBusinessUnit = u.BusinessUnitName,

                            AssignmentType = "Direct",
                            TeamName = "",
                            TeamBusinessUnit = "",

                            RoleName = r.Name,
                            RoleBusinessUnit = r.BusinessUnitName,

                            Duplicate = isDup
                        });

                        any = true;
                    }
                }

                if (teamRoles.TryGetValue(u.UserId, out var tr))
                {
                    foreach (var t in tr.OrderBy(x => x.TeamName, StringComparer.OrdinalIgnoreCase)
                                        .ThenBy(x => x.RoleName, StringComparer.OrdinalIgnoreCase))
                    {
                        var isDup = duplicates.Contains((u.UserId, t.RoleId));

                        rows.Add(new UserRoleRow
                        {
                            UserId = u.UserId,
                            RoleId = t.RoleId,
                            TeamId = t.TeamId,

                            UserFullName = u.FullName,
                            UserEmail = u.Email,
                            UserBusinessUnit = u.BusinessUnitName,

                            AssignmentType = "Team",
                            TeamName = t.TeamName,
                            TeamBusinessUnit = t.TeamBusinessUnitName,

                            RoleName = t.RoleName,
                            RoleBusinessUnit = t.RoleBusinessUnitName,

                            Duplicate = isDup
                        });

                        any = true;
                    }
                }

                // Ensure "all users" are present, even if they have no roles
                if (!any)
                {
                    rows.Add(new UserRoleRow
                    {
                        UserId = u.UserId,
                        RoleId = Guid.Empty,
                        TeamId = Guid.Empty,

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

        // ==========================
        // PAGING HELPER (QueryExpression)
        // ==========================
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

                if (!ec.MoreRecords)
                    break;

                qe.PageInfo.PageNumber++;
                qe.PageInfo.PagingCookie = ec.PagingCookie;
            }

            return result;
        }

        // ==========================
        // HELPERS (Aliased + Id extraction)
        // ==========================
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

        // ==========================
        // MODELS
        // ==========================

        private sealed class UserInfo
        {
            public Guid UserId { get; set; }
            public string FullName { get; set; }
            public string Email { get; set; }
            public string BusinessUnitName { get; set; }
            public bool IsDisabled { get; set; }
        }

        private sealed class RoleInfo
        {
            public Guid RoleId { get; set; }
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
        }

        private sealed class UserRoleRow
        {
            public Guid UserId { get; set; }
            public Guid RoleId { get; set; }
            public Guid TeamId { get; set; }

            public string UserFullName { get; set; }
            public string UserEmail { get; set; }
            public string UserBusinessUnit { get; set; }

            public string AssignmentType { get; set; } // Direct / Team / None
            public string TeamName { get; set; }
            public string TeamBusinessUnit { get; set; }

            public string RoleName { get; set; }
            public string RoleBusinessUnit { get; set; }

            public bool Duplicate { get; set; }
        }
    }
}
