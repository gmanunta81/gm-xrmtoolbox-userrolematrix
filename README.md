# GM.XrmToolBox.UserRoleMatrix

XrmToolBox plugin for Microsoft Dataverse / Dynamics 365 that lists **users and their security roles**:
- Direct roles (System User)
- Roles inherited via **Owner Teams**
- Includes Business Unit info
- Highlights duplicates (same role assigned both Direct + Team)
- Filters (Business Unit, Team, Assignment Type)
- Export to CSV / Excel

> Repo: `gm-xrmtoolbox-userrolematrix`  
> Assembly + NuGet Id: `GM.XrmToolBox.UserRoleMatrix`  
> Tool display name: choose one:
> - `GM.XrmToolBox.UserRoleMatrix`
> - `User Role Matrix`

## CI Status

![CI](https://github.com/<YOUR_GITHUB_USERNAME>/gm-xrmtoolbox-userrolematrix/actions/workflows/ci.yml/badge.svg)

> Replace `<YOUR_GITHUB_USERNAME>` with your GitHub username.

---

## Features

- Load all users from Dataverse
- Load all roles assigned:
  - Directly to the user
  - Via Owner Teams the user is member of
- Grid view with:
  - User, Email, User BU
  - Assignment Type (Direct / Team)
  - Team, Team BU
  - Role, Role BU
  - Duplicate flag
- Search box (text filter across columns)
- Dropdown filters:
  - Business Unit
  - Team
  - Assignment Type (All / Direct / Team)
- Export:
  - CSV
  - Excel (.xlsx)

---

## Requirements

- XrmToolBox installed
- A Dataverse / Dynamics 365 environment
- Enough privileges to read:
  - Users, Teams, Roles, Business Units, Team Membership, role assignments

---

## Installation (Manual)

1. Build the project (Release recommended)
2. Copy the plugin output DLL and required dependencies into XrmToolBox Plugins folder.
   - Typical location:
     - `%APPDATA%\MscrmTools\XrmToolBox\Plugins`

> If you distribute via NuGet/Tool Library later, you won't need manual copying.

---

## Usage

1. Open XrmToolBox
2. Connect to your environment
3. Open the tool:
   - `GM.XrmToolBox.UserRoleMatrix` (or `User Role Matrix`)
4. Click **Load Users & Roles**
5. Use filters, search box and export.

---

## Debugging inside XrmToolBox (developer mode)

Recommended approach (overridepath):
- Start XrmToolBox with:
  - `/overridepath:.`
- Ensure the plugin DLL is copied into:
  - `.\Plugins\`

(Exactly like you already validated in your setup.)

---

## Versioning policy

We use **Semantic Versioning**: `MAJOR.MINOR.PATCH`

- `PATCH` = bug fix only
- `MINOR` = backward-compatible new features
- `MAJOR` = breaking changes

**Git tags**:
- `v1.0.0`, `v1.1.0`, ...

Each release should:
1. Update `AssemblyInfo.cs` version (AssemblyVersion/FileVersion/InformationalVersion)
2. Update `CHANGELOG.xml`
3. Create a tag `vX.Y.Z`
4. Create a GitHub Release and attach the compiled plugin ZIP.

---

## Project structure

- `GM.XrmToolBox.UserRoleMatrix.sln` (solution)
- `GM.XrmToolBox.UserRoleMatrix/` (project)
- `.github/workflows/ci.yml` (CI build workflow)

---

## Author

Giovanni Manunta  
Email: gmanunta81@gmail.com  
XRM Toolbox Enthusiast
