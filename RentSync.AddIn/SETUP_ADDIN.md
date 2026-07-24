# RentSync.AddIn — setup in Visual Studio 2026 (step by step)

The VSTO project itself must be created from the VS template (it generates
the .csproj with VSTO build targets, manifests, and Globals — these cannot
be hand-written reliably). Then you drop in the files from this folder.

## Step 1 — Create the project

1. Open `RentSyncAnalytics.sln` in Visual Studio 2026.
2. Right-click solution → Add → New Project → search **"Excel VSTO Add-in"**
   (C#). If the template is missing: VS Installer → Modify → check
   **Office/SharePoint development** workload.
3. Name: `RentSync.AddIn`, location: the solution folder.
4. Target framework: **.NET Framework 4.8**.

## Step 2 — Reference Core and packages

1. Right-click RentSync.AddIn → Add → Project Reference → **RentSync.Core**.
   (Core multi-targets net10.0 + net48; the add-in automatically consumes
   the net48 build.)
2. NuGet for RentSync.AddIn:
   - `Microsoft.Extensions.DependencyInjection` 8.0.1
   - `Portable.System.DateTimeOnly` 8.0.2
3. In project Properties → Build → set **LangVersion** by editing the
   .csproj: add `<LangVersion>latest</LangVersion>` to the first
   PropertyGroup. (Modern syntax compiles fine on Fx 4.8; runtime-dependent
   features are covered by the polyfill packages in Core.)

## Step 3 — Drop in the files

1. Replace the generated `ThisAddIn.cs` **body** with ours (keep the
   generated `ThisAddIn.Designer.cs` untouched).
2. Add: `RibbonController.cs`, `ExcelActions.cs`, `VbaBridge.cs`,
   `SettingsForm.cs`, `WindowsCredentialTokenProvider.cs`.
3. Add `RibbonUI.xml` to the project, then:
   Project Properties → Resources → Add Existing File → `RibbonUI.xml`
   (this makes `Properties.Resources.RibbonUI` available).
   Set the xml file's Build Action to **None** (it lives in Resources).
4. Wire the ribbon: in `ThisAddIn.cs` VSTO region add the override:

```csharp
protected override Microsoft.Office.Core.IRibbonExtensibility
    CreateRibbonExtensibilityObject() => new RibbonController();
```

## Step 4 — First run

1. Set RentSync.AddIn as startup project → F5. VS registers the add-in and
   launches Excel.
2. You should see the **RentSync Data** tab. "Refresh Data" will call the
   demo endpoint; with no token stored it runs in anonymous mode.
3. Logs: `%LocalAppData%\RentSync\logs\` (CLEF JSON).
   Telemetry: `%LocalAppData%\RentSync\telemetry\*.jsonl`.
4. VBA test (Alt+F11, any module):

```vb
Sub TestRentSync()
    Dim rs As Object
    Set rs = Application.COMAddIns("RentSync.AddIn").Object
    MsgBox rs.GetVersion
    rs.RefreshData
End Sub
```

## Known first-run issues

- **"Office solution failed to load"** → clean the VSTO cache:
  close Excel, delete `%LocalAppData%\Apps\2.0`, rebuild.
- **Ribbon not visible** → the CreateRibbonExtensibilityObject override
  from Step 3.4 is missing, or RibbonUI.xml is not in Resources.
- **Core restore fails for net48 + DuckDB.NET** → send me the exact error;
  fallback plan exists (conditional SQLite cache for the net48 target).

## What is deliberately NOT here yet (next iterations)

1. DevExpress-styled forms (SettingsForm is plain WinForms on purpose).
2. AI Smart Import (Claude API column mapper).
3. MSAL / Entra ID login.
4. FastReport PDF export; WiX MSI installer; code signing pipeline.
