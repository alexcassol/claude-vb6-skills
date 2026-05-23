# Claude — VB6 Guidelines

Condensed conventions for Visual Basic 6 development. Drop this file at the
root of a VB6 project to give Claude Code persistent context about the
codebase style. For full guidance install the plugin (see project README).

---

## 1. Case-preservation is absolute in existing code

VB6 is case-insensitive; diff tools are not. Editing `m_StrName` to `m_strName`
in existing code pollutes the diff and hides real changes. **Never re-case
existing identifiers** — even keywords, type names, and constants. New code
uses the conventions below.

## 2. Mandatory basics

- `Option Explicit` at the top of every module
- Explicit type on every `Dim`, `Function`, `Sub`
- `Long` over `Integer` (Integer is 16-bit, overflows at 32,768)
- `ByVal`/`ByRef` explicit on every parameter
- No `Variant` except for ADO `Variant`-returning APIs
- Encoding: Windows-1252, CRLF, 4-space indent
- Never edit `.frx` files
- Never reformat `.frm` headers

## 3. Hungarian notation — scope + type prefix, lowercase

| Scope                   | Prefix |
| ----------------------- | ------ |
| Module-level (`.bas`)   | `m`    |
| Module-level (`.cls`)   | `m_`   |
| Global public (`.bas`)  | `g`    |
| Parameter / local       | (none) |

| Type     | Prefix | Example                            |
| -------- | ------ | ---------------------------------- |
| String   | `str`  | `Private m_strName As String`      |
| Integer  | `int`  | `Dim intIndex As Integer`          |
| Long     | `lng`  | `ByVal lngCount As Long`           |
| Currency | `cur`  | `Private m_curTotal As Currency`   |
| Boolean  | `bln`  | `Private m_blnActive As Boolean`   |
| Double   | `dbl`  |                                    |
| Date     | `dtm`  |                                    |
| Object   | `obj`  | `Private mobjCnn As Object`        |
| Variant  | `var`  |                                    |
| Recordset| `rs`   | `Private mrsCustomers As Object`   |
| Dictionary| `dic` | `Private mdicConfig As Dictionary` |

Acronyms preserve uppercase: `m_strNSU`, `m_strCNPJ`, `m_strXML`.

## 4. Class module structure

```vb
'--------------------------------------------------------------------------------
'    Component  : clsCustomer
'    Project    : <YourProject>
'
'    Description: ...
'    Modified   :
'--------------------------------------------------------------------------------
Option Explicit

' Variable to hold 'Name' property value
Private m_strName As String

Public Property Get Name() As String
    Name = m_strName
End Property

Public Property Let Name(ByVal strValue As String)
    m_strName = strValue
End Property
```

- 80-hyphen header delimiters
- Comment line above each private field
- `Property Get` / `Let` / `Set` — never public fields
- No validation/logic in `Let`; that goes in named methods

## 5. Error handling — CSEH pattern

Mark the procedure with `'CSEH: ErrRaise` (or `'CSEH: MsgBox` for top-level
UI handlers). Use `<EhHeader>` / `<EhFooter>` tags (do not alter their
format — automation tools depend on them):

```vb
'CSEH: ErrRaise
Public Function DoWork(...) As <Type>

        '<EhHeader>
        On Error GoTo DoWork_Err
        EnterMethod "modX", "DoWork"
        '</EhHeader>

100     ' ... numbered executable lines ...

    '<EhFooter>
        On Error GoTo 0

DoWork_Exit:
        ' cleanup
        ExitMethod "modX", "DoWork"
        Exit Function

DoWork_Err:
        Dim sDoWork_Err As String
        sDoWork_Err = "Error: " & Err.Number & " - " & Err.Description & " (" & Erl & ")"
        ' cleanup
        ExitMethod "modX", "DoWork"
        Err.Raise vbObjectError + 100, "<App>.modX.DoWork", sDoWork_Err
    '</EhFooter>
End Function
```

Rules:

- Line numbering: start at 100, increment by 5, executable lines only
- Cleanup before `Err.Raise` (everything after is unreachable)
- `ExitMethod` called in **both** `_Exit` and `_Err`
- `On Error Resume Next` is forbidden outside cleanup/dispose

## 6. ADO via late binding

- `Private mobjCnn As Object`, never `As ADODB.Connection`
- `Set objRs = CreateObject("ADODB.Recordset")`, never `New`
- No reference to the ADO type library
- Define ADO constants locally (`Private Const adInteger As Long = 3`)
- `ConnectionTimeout` set before `Open` (mandatory)
- Disconnected pattern for SELECT: `CursorLocation = adUseClient`, then
  `Set rs.ActiveConnection = Nothing` after `Open`
- **Parameterized queries** via `ADODB.Command` for any user-supplied value;
  concatenation only for structural SQL (table names, conditional JOINs)
- Multibank: `SetParameter` / `GetParameter` helpers abstract the
  `@param` (MSSQL) vs `in_param`/`out_param` (PostgreSQL) naming

## 7. Trace pattern (`EnterMethod` / `ExitMethod`)

Every non-trivial procedure pushes itself onto a manual stack on entry and
pops on exit. Resolves VB6's lack of native stack trace. See
`modLog.bas` for the implementation.

`ExitMethod` must be called on **both** the success and error paths.

## 8. Use existing utility wrappers

- `MsgBoxCritical`, `MsgBoxInfo`, `MsgBoxWarning`, `MsgBoxQuestion` —
  never raw `MsgBox`
- `fg_isNull`, `fg_Coalesce`, `fg_NullIf` — NULL handling
- `IsValidCNPJ`, `IsValidCPF` — tax ID validation
- `PadLeft`, `PadRight`, `RemoveAccents`, `TrimAll` — string ops
- `ToSQL` / `SanitizeSqlString` — legacy SQL building (prefer parameters)

When in doubt, search the project's utility module (`modUtils`, `modFuncoes`,
`modCommon`) before writing a new helper.

## 9. Pre-flight check

- [ ] `Option Explicit` present
- [ ] All Dims and Functions typed
- [ ] `ByVal`/`ByRef` explicit
- [ ] Case of existing identifiers untouched
- [ ] CSEH pattern applied to non-trivial procedures
- [ ] Lines numbered where `Erl` is used
- [ ] `EnterMethod`/`ExitMethod` paired correctly (including in `_Err`)
- [ ] ADO via late binding; parameters for user values
- [ ] `.frx` files untouched
- [ ] Encoding (Windows-1252) and CRLF preserved
- [ ] Edited only what was requested

---

## Project-specific overrides

Append project-specific rules below this line:

<!-- Example:

## Project Acme POS

- Module naming: `modX` prefix for shared modules
- Database: MSSQL only (no Postgres branch)
- Error offset range: 100-199 (validation), 200-299 (db), 300+ (IO)
- Connection string built by `BuildConnectionString` in `modDB`

-->
