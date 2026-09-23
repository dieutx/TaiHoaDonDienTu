# MIGRATION RESULT

Original: `TaiHoaDonDienTu_v6.3.xlsm`

Built: `App_v6.3.0.xlsm`

Status: **PARTIAL**

| Area | Result | Evidence |
|---|---|---|
| Workbook structure | PASS | 9 worksheet names, CodeNames and visibility values match. |
| VBA Modules | PASS | 11 standard modules match by component name, type and normalized code hash. |
| Classes | PASS | No rebuildable class modules exist in the original or built workbook. |
| UserForms | PASS | 4 UserForm component names and designer metadata are present. |
| UserForm Controls | PASS | Control names and ProgIDs match the original inventory. |
| Sheet CodeNames | PASS | All 9 worksheet CodeNames match. |
| Named Ranges | PASS | All 3 workbook names match. |
| VBA References | PASS | All 9 references match; none is marked missing on this machine. |
| Macro Assignments | PASS | Shape targets match after normalizing the workbook-qualified filename. |
| Reopen test | PASS | The built workbook was closed and reopened in a fresh Excel COM instance. |
| UserForm runtime smoke | NOT VERIFIED | One earlier run instantiated 4/4 forms, but repeat runs were not stable enough in the unattended automation host to count as reproducible evidence. |
| Critical business macros | NOT VERIFIED | Download/login/write macros require credentials, network access and representative test data and can mutate workbook data. |
| VBA compile | NOT VERIFIED | Excel exposes no supported programmatic compile API. |

## Expected serialization differences

Excel rebound three shape `OnAction` prefixes on sheet `MENU` from `TaiHoaDonDienTu_v6.3.xlsm!…` to `App_v6.3.0.xlsm!…`. The macro targets remain `dangnhap`, `TaiHoaDon` and `moUFTrichXuat`.

UserForm import adds one boundary blank line to each form code module. Normalized code hashes match after trimming leading/trailing whitespace; control metadata and `.frx` files are present.

## Dependencies and risks

The VBA project references VBA, Excel, stdole, Office, MSForms, MSXML2 6.0, SHDocVw, Scripting and MSHTML. None is missing on the build machine. `SHDocVw` and `MSHTML` are legacy Windows/Internet Explorer-era dependencies and should be checked on every target workstation. ActiveX/MSForms behavior can differ between 32-bit and 64-bit Office.

The source scan found no `Workbooks("…")`, `Windows("…")`, or workbook-qualified `Application.Run` hard-codes in VBA code. Workbook-qualified `Shape.OnAction` values were detected and Excel updated them automatically during `SaveAs`.

The build version is carried by the output filename. No custom document property is added because repeated COM access to `CustomDocumentProperties` was unstable in this environment.

## Manual verification still required

Open `dist/App_v6.3.0.xlsm` interactively, enable macros, instantiate/show every form, exercise the three menu buttons, run representative purchase/sales download flows with non-production test credentials, and validate written invoice totals against known source XML. Use **Debug > Compile VBAProject** in the VBA editor and confirm no target-machine reference is missing.
