# Microsoft Places Master Script

A PowerShell script for managing Microsoft Places buildings, floors, rooms,
desks, desk pools, Wi-Fi BSSIDs, and maps. Use the interactive menu for guided
operations, or load its functions into your own automation.

**Public/portable edition:** Internal contacts, internal documentation links,
tenant-specific examples, and hard-coded corporate email domains have been
removed. Microsoft Places/Exchange product names, public documentation links,
cmdlets, API types, and device metadata are retained for compatibility.

> **Administrative tool:** Changes affect your connected tenant. Start with a
> small approved batch, inspect the results, and preserve an inventory before
> destructive operations. There is no global dry-run or automatic rollback.

## 1. Requirements and setup

- **PowerShell 7+** (`pwsh`), not Windows PowerShell 5.1.
- The **MicrosoftPlaces** module; **ExchangeOnlineManagement** for mailbox and
  calendar-processing workflows.
- An account with the appropriate Places/Exchange administrative permissions.
  Activate privileged access if your organization uses PIM. Module installation
  does not grant permissions.
- Existing resource mailboxes for room onboarding, bookable desks, and workspaces.
  This script does not create Exchange mailboxes.

Check your environment:

```powershell
$PSVersionTable.PSVersion
Get-Module -ListAvailable MicrosoftPlaces, ExchangeOnlineManagement
```

Install missing modules, then change to the downloaded script folder:

```powershell
Install-Module -Name MicrosoftPlaces -Scope CurrentUser
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
Set-Location 'C:\Tools\places-master-script-vanilla'
```

If execution policy blocks the script, follow your organization's approved
signing/execution-policy process. Do not disable protections globally.

## 2. Quick start

```powershell
.\microsoft-places-mgmt-master-script.ps1
```

1. Select **0** (or press **Enter**) to connect to Microsoft Places.
2. Select **1** to inspect existing objects.
3. Choose an operation below and provide its CSV or other inputs.
4. Review the console, reports, and actual resource state.
5. Select **X** at the main menu to stop the transcript and exit.

**Using bare aliases instead of full email addresses?**

```powershell
.\microsoft-places-mgmt-master-script.ps1 -DefaultDomain 'example.com'
```

In options **2**, **3**, and **10 > 1**, `room101` then resolves to
`room101@example.com`. Full email addresses are left unchanged apart from trimming
surrounding whitespace. Without `-DefaultDomain`, those workflows require full
email addresses. Blank addresses and aliases without a configured domain stop
the batch before resource changes; no fallback mailbox is targeted.
Other workflows pass identifiers directly to their service cmdlets.

**Exchange-dependent workflows:** Options **5** and **7** include a submenu
**0** to connect. For **8 > 1** and **8 > 3**, connect in the same session before
launching:

```powershell
Connect-ExchangeOnline
.\microsoft-places-mgmt-master-script.ps1 -DefaultDomain 'example.com'
```

Select main-menu **0** as well; Exchange and Places connections are separate.

## 3. Menu reference

`4 > 1` means main-menu **4**, then submenu **1**.

| Option | Operation | Input |
|---|---|---|
| **0** | Connect to Microsoft Places | Interactive sign-in |
| **1** | Search names and display linked hierarchy | Name or wildcard pattern |
| **2** | Onboard/update existing rooms under floors | Room CSV; `sample` generates a template |
| **3** | Update an existing desk pool's label, capacity, and section | Desk pool attributes CSV |
| **4 > 1 / 2** | Create buildings / floors | Corresponding CSV |
| **5** | Enable auto-release | Single resource or `Alias` CSV; Exchange required |
| **6** | Add/update building BSSID devices | `PlaceId,BSSID` CSV |
| **7** | Apply capacity or booking-policy changes | Resource identifier or `Alias` CSV |
| **8 > 1** | Create mailbox-linked desks and set their modes | Bookable desks CSV; Exchange required |
| **8 > 2** | Create desks without mailboxes | Non-bookable desks CSV |
| **8 > 3** | Configure desk pools (Exchange workspaces) | Workspace CSV; Exchange required |
| **9** | Upload a building map | Building PlaceId and local map file |
| **10** | Unlink rooms or remove floors/buildings | Corresponding CSV; see removal warnings |
| **11** | Remove a building hierarchy | Building PlaceId; confirmation required |
| **12** | Change existing desk modes | Desk-mode CSV |
| **Help / X** | Display help / exit | Main menu |

Submenus **4**, **5**, **7**, and **10** use **9** to return toward the main menu;
main-menu **9** uploads a map. Options **1** and **12** prompt as `SearchString:`
and `CsvPath:` respectively.

## 4. CSV examples and workflows

Save examples as **comma-delimited `.csv` files**, with exactly the shown headers.
Replace all example values with your own. At interactive path prompts, paste
paths **without surrounding quotes**, including paths containing spaces.
Quotes are appropriate inside PowerShell commands.

Create buildings before floors. Rooms link directly to floors; desks and
workspaces use sections under floors. Desk workflows can prompt to create missing
sections. Use unique, exact parent names: several functions take the first match
when names are duplicated.

### Buildings and floors -- option 4

**Buildings (4 > 1):**

```csv
BuildingName,CountryOrRegion,City,State,Street,PostalCode
Building A,US,Seattle,WA,100 Example Street,98101
```

**Floors (4 > 2):**

```csv
ParentName,FloorLabel
Building A,Level 1
```

`ParentName` is the building display name. Floor creation asks you to choose if
multiple buildings match. Building creation does not check for duplicates first.

### Existing rooms -- option 2

```csv
PrimarySmtpAddress,Building,FloorLabel,Floor,Capacity,Label,MTREnabled,AudioDeviceName,DisplayDeviceName,VideoDeviceName
room101@example.com,Building A,Level 1,1,6,Meeting Room 101,False,,,
```

The building and room resource must exist. The script creates the floor if
missing, sets its sort order, links the room, and updates room properties.
Use integer `Floor` and `Capacity` values: missing/invalid values become **1**
and **0**, respectively. `Label` is written even when blank.
Use `True` or `False` for `MTREnabled`; blank attempts to preserve its current value.
Device names update only when populated.

Type `sample` instead of a path to generate `Sample(template)_OnboardRoom.csv`.
Template fields `Alias`, address fields, `RoomList`, and `WheelChairAccessible`
are not applied by this function; `REQ_Type` is only recorded in reports.

### Existing desk pool attributes -- option 3

```csv
Alias,Building Name,Floor Name,DeskPoolName,Capacity
pool101@example.com,Building A,Level 1,North Pool,12
```

The resource must already exist and is processed only if its Places type is
**Space** by default. Building and floor must exist. The function finds or offers
to create `section <Floor Name>` (such as `section Level 1`), then updates the
resource's parent, label, and capacity. Supply a valid integer capacity.

### Mailbox-linked desks -- option 8 > 1

```csv
Building,FloorLabel,Section,SpaceID,PrimarySmtpAddress,Mode,AssignedTo
Building A,Level 1,North,101,desk101@example.com,Bookable,
Building A,Level 1,North,102,desk102@example.com,Assigned,user@example.com
```

Creates desks named `Desk <SpaceID>`. Converts existing mailboxes to **Desk**,
hides them from address lists, and links them to desk objects. Supply `Mode`:
`Bookable`, `Walkup`, or `Assigned`; assigned desks need `AssignedTo`.
Missing/unrecognized modes become **Offline**.

**Mailbox conversion runs before hierarchy validation**, so it can succeed even
when a later row fails. Existing desks skip creation but still receive mailbox
linking and mode updates. A created desk is not necessarily a successfully
linked, bookable desk; inspect the final state.

### Desks without mailboxes -- option 8 > 2

```csv
Building,FloorLabel,Section,SpaceID,Capacity
Building A,Level 1,North,103,1
```

Creates `Desk 103` without a mailbox. Optional `Capacity` defaults to **1**.
Existing desks are skipped. This workflow does not explicitly set desk mode.

### Desk pools / workspaces -- option 8 > 3

```csv
Building,FloorLabel,Section,Display Name,PrimarySmtpAddress,Capacity
Building A,Level 1,North,North Pool,pool101@example.com,12
```

Converts existing mailboxes to **Workspace**, enables capacity enforcement, and
links them to sections. Optional capacity defaults to **1**.
Mailbox changes occur before hierarchy validation. `Display Name` is required
but is used for reporting, not renaming. There is no dedicated results CSV for
this workflow; inspect the console/transcript and service state.

### Existing desk modes -- option 12

```csv
Display Name,Mode,AssignedTo
Desk 101,Bookable,
Desk 102,Assigned,user@example.com
Desk 103,Walkup,
Desk 104,Offline,
```

Modes are **Bookable**, **Assigned**, **Walkup**, and **Offline**. Always provide
an assignee email for `Assigned`. Blank/unknown modes become **Offline**.
The function passes `Display Name` directly to `Get-PlaceV3 -Identity`.
If names do not resolve, put PlaceIds in that same column.
Changing mode does not create/link mailboxes or validate licensing.

### Auto-release and booking-policy changes -- options 5 and 7

Both support single-resource operations or this bulk format:

```csv
Alias
room101@example.com
room102@example.com
```

**Option 5:** Connect using submenu **0**, then select **1** (single) or **2**
(bulk). Choose **1-20 minutes**, or Enter for **15**. Enables `EnableAutoRelease`
and sets `PostReservationMaxClaimTimeInMinutes`; there is no disable action.
The selected timeout applies to the whole batch. Service/device prerequisites
still apply.

**Option 7:** Use resource mailboxes, not user mailboxes.

| Submenu | Change |
|---|---|
| **0** | Connect to Exchange Online |
| **1 / 2** | Set `EnforceCapacity = True` for one resource / CSV |
| **3 / 4** | Set `AllBookInPolicy = True` for one resource / CSV |
| **5** | Update numeric Places capacity for one resource |
| **?** | Display the bulk CSV format |

Enforcing capacity does not set its numeric value. `AllBookInPolicy = True`
broadens in-policy booking access; it is not a universal availability fix.

### BSSID devices -- option 6

```csv
PlaceId,BSSID
<building-place-id>,AA:BB:CC:DD:EE:01
```

Supply one building PlaceId and Wi-Fi BSSID per row. Reports classify devices as
added, updated, ignored, or failed.

**Reruns:** `processed_buildings.csv` skips previously recorded building IDs,
including when input changes. A building may be marked processed despite
individual failures. Review reports; to deliberately reprocess a building, back
up the tracking file and remove only that building's row before rerunning.

### Building maps -- option 9

Enter a building PlaceId or type `find` to search by name/wildcard. Supply the
full map path without quotes. The script calls **`New-Map`**, not a separate
update/replacement command; an existing map may cause failure.
Its extension check accepts `.zip`, `.png`, `.jpg`, `.jpeg`, and `.svg`, but
actual content/format acceptance is controlled by the service. Use a map package
supported by Microsoft Places.

## 5. Removal: options 10 and 11

**Review targets first. Some actions have no confirmation, and deletions are not
automatically reversible.** Normally unlink rooms/workspaces, then remove
sections/floors, then the building.

| Action | CSV headers / input | Behavior and limitations |
|---|---|---|
| **10 > 1** | `PrimarySmtpAddress` | Clears room `ParentId`; does not delete mailboxes |
| **10 > 2** | Disabled | Displays an unavailable warning; does not prompt for input or change resources |
| **10 > 3** | `ParentName,FloorLabel` | Deletes floors after `Yes`; does not unlink children for you |
| **10 > 4** | `BuildingName,BuildingId` | Looks up by **BuildingName**, takes the first match, and deletes it; CSV `BuildingId` is only logged |
| **11** | Building PlaceId | Exports inventory, asks for `Yes`, unlinks rooms/desks, removes other returned types, then removes the building |

Use the exact CSV headers above with one target per row. Floor removal uses
wildcard matching internally and rejects ambiguous matches; use exact names.
For **11**, review `Building_Hierarchy_Export_<timestamp>.csv` before confirming.
This inventory is not a restorable backup. The function returns without deletion
when no children are found. Only `Room` and `Desk` types receive explicit unlink
handling; other returned types, including any `Space` resources, use `Remove-Place`.

**Floor-removal preview:** `Remove-Floors -WhatIf` looks up and previews matching
floors without deleting them, asking for deletion confirmation, or writing a log.
Normal execution retains the batch `Yes` confirmation; `-Confirm` additionally
requests confirmation for each floor. See the example below.

## 6. Use functions in automation

Dot-source with **`-NoMenu`** to load functions without a menu, banner, connection,
or session transcript. It does not make every function non-interactive.

```powershell
. .\microsoft-places-mgmt-master-script.ps1 -NoMenu -DefaultDomain 'example.com'
Connect-MicrosoftPlaces

# Read-only search; wildcards are supported.
FindBuildingOrFloor -SearchString '*Building*' -ExportResults

# Preview desk mode changes without applying them or exporting a result CSV.
SetDeskMode -CsvPath 'C:\PlacesData\desk-modes.csv' -WhatIf -NoExport

# Preview floor removal without deleting floors or writing a log.
Remove-Floors -CsvFilePath 'C:\PlacesData\floors-to-remove.csv' -WhatIf
```

After review, apply the same file:

```powershell
SetDeskMode -CsvPath 'C:\PlacesData\desk-modes.csv'
```

Other parameterized functions:

```powershell
OnboardRoom -FileName 'C:\PlacesData\rooms.csv'
DeskPoolAttri -InputFile 'C:\PlacesData\pool-attributes.csv'
BSSID -InputFile 'C:\PlacesData\bssids.csv'
```

Desk pool updates may prompt to create sections. Other workflows also prompt,
some error paths use `exit`, and the script does not guarantee job-friendly exit
codes or structured results. For unattended jobs, use functions whose prerequisites
are already satisfied and establish approved non-interactive service
authentication separately. Never embed credentials or tokens in this script.

`SetDeskMode -WhatIf` guards mode writes, but its summary can still report
"Success" for simulated actions. `Remove-Floors -WhatIf` reports previewed floors
separately from successful deletions. Other functions do not offer a reliable
general-purpose preview.

## 7. Reports, verification, and troubleshooting

Output paths are relative to the **current PowerShell directory**, not necessarily
the script folder. Keep the directory writable and archive reports before reruns.

| Workflow | Output |
|---|---|
| Interactive session | `Script_Transcript_<timestamp>.txt` |
| Rooms | `update_placesv3_output_<timestamp>.csv`, `Failed_Onboarding_<timestamp>.csv`, `update_places_transcript_<timestamp>.txt` |
| Pool attributes | `DeskPool_Results_<timestamp>.csv` |
| Building creation | `NewBuilding_<BuildingName>.csv` |
| Desks | `DesksCreated_<timestamp>.csv` or `NonBookableDesksCreated_<timestamp>.csv` |
| Desk modes | `DeskModeUpdate_<timestamp>.csv` |
| BSSID | `output\<PlaceId>-Added.csv`, `-Updated.csv`, `-Ignored.csv`, `-Failed.csv`; `processed_buildings.csv` |
| Room unlink / building removal | `Remove_Room_Output.csv` / `Remove_bldg_Output.csv` |
| Floor / hierarchy removal | `Floor_Removal_Log_<timestamp>.csv` / `Building_Hierarchy_Export_<timestamp>.csv` |
| Search with `-ExportResults` | `PlacesSearch_<timestamp>.csv` |

Some outputs appear only when results exist. Fixed-name outputs may be overwritten
or appended. `-NoMenu` disables only the session transcript; functions can still
create their own reports/transcripts.

**Verify actual state:** Some inherited cmdlet calls can emit non-terminating
errors without a reliable failure status. Desk exports can contain rows recorded
before mailbox linking/mode changes finish; hierarchy searches may count
overlapping descendants.

In a connected session outside the menu:

```powershell
Get-PlaceV3 -Identity '<PlaceId-or-resource-email>' | Format-List *

# Requires an Exchange Online connection.
Get-CalendarProcessing -Identity 'room101@example.com' |
    Format-List EnableAutoRelease, PostReservationMaxClaimTimeInMinutes,
                EnforceCapacity, AllBookInPolicy
```

| Problem | Check |
|---|---|
| Cmdlet not found | Correct module, PowerShell 7+, and required service connection |
| Access denied | Tenant/account, administrative role, and privileged-role activation |
| File/header error | Unquoted prompt path, exact headers including spaces, comma-delimited CSV |
| Alias requires a domain | Use a full email or launch/load with `-DefaultDomain` |
| Parent not found | Exact names, existing hierarchy, and duplicate-name ambiguity |
| Pool skipped / BSSID skipped | Places type `Space` / `processed_buildings.csv`, respectively |
| Unexpected Offline desk | Missing/invalid `Mode` in the input CSV |

### Publishing safely

Upload **this folder only**, not the original script, tenant CSVs, logs, or maps.
The included `.gitignore` excludes common input/output files and map formats;
it does not sanitize their contents or untrack files already committed.
Reports can contain resource/user addresses, assignments, locations, and BSSIDs.
Check every staged file and ensure you have permission to publish the source;
the included MIT license does not replace that permission check.

### License

Released under the [MIT License](LICENSE). See `LICENSE` for permissions,
conditions, and the warranty disclaimer.

### Public references

- [Microsoft Places overview](https://learn.microsoft.com/en-us/microsoft-365/places/places-overview)
- [Configure desk booking](https://learn.microsoft.com/en-us/microsoft-365/places/configure-desk-booking)
