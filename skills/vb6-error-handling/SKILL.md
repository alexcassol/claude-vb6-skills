---
name: vb6-error-handling
description: Applies the CSEH (Code Style for Error Handler) pattern when adding or editing error handling in Visual Basic 6 procedures. Covers the structured On Error GoTo template with EhHeader/EhFooter regions, two CSEH styles (ErrRaise for utility functions that propagate to callers, and MsgBox for top-level UI procedures), line numbering (starting at 100, increments of 5) to enable Erl in error handlers, structured re-raise via Err.Raise with vbObjectError offset and qualified source string in App.Module.Procedure format, error message format including Err.Number, Err.Description and Erl, cleanup-before-raise discipline, and the legitimate use of On Error Resume Next strictly limited to cleanup and dispose routines. Activates whenever modifying VB6 procedures that have or need On Error handlers, the CSEH marker comment, numbered code lines, Err.Raise calls, or any discussion of error propagation, error logging, or fault tolerance in VB6.
---

# VB6 Error Handling — CSEH Pattern

VB6 has no `try`/`catch`. The model is `On Error GoTo <label>` with structured
cleanup. The **CSEH** (Code Style for Error Handler) convention standardizes
this pattern across the codebase so error paths are predictable, log output is
consistent, and `Erl` (error line) is meaningful in production logs.

CSEH style is declared by a marker comment above the procedure signature:

```vb
'CSEH: ErrRaise
Public Function GetActiveCustomer(...) As Object
```

The `<EhHeader>` and `<EhFooter>` tags delimit regions that may be regenerated
by external tooling (MZ-Tools or similar). **Never alter the delimiter format
or remove the tags** — automation depends on them.

## 1. Two CSEH styles

| Style       | Marker             | Handler behavior                                                  | Use for                                  |
| ----------- | ------------------ | ----------------------------------------------------------------- | ---------------------------------------- |
| `ErrRaise`  | `'CSEH: ErrRaise`  | Re-raises with `Err.Raise vbObjectError + <offset>, ..., msg`     | Utility functions, library routines, any procedure called by other code |
| `MsgBox`    | `'CSEH: MsgBox`    | Displays critical dialog and returns to caller without re-raising | Top-level UI event handlers, command buttons, menu items |

**Default for new code: `ErrRaise`**. The `MsgBox` style is reserved for the
outermost layer (event handlers) where errors must be surfaced to the user and
the call stack ends.

## 2. ErrRaise template

Canonical structure for a function that propagates errors to its caller:

```vb
'CSEH: ErrRaise
Public Function GetActiveCustomer(ByVal lngCustomerID As Long) As Object

        '<EhHeader>
        On Error GoTo GetActiveCustomer_Err

        EnterMethod "modDB", "GetActiveCustomer"
        '</EhHeader>

        Dim objCmd As Object
        Dim objRs  As Object
        Dim strSQL As String

100     InitializeDBConnection

105     Set objCmd = CreateObject("ADODB.Command")
110     Set objCmd.ActiveConnection = mobjCnn
115     objCmd.CommandType = adCmdText

120     strSQL = "SELECT * FROM Customer WHERE idCustomer = ? AND custActive = 1"
125     objCmd.CommandText = strSQL

130     objCmd.Parameters.Append objCmd.CreateParameter("idCustomer", adInteger, adParamInput, , lngCustomerID)

135     Set objRs = objCmd.Execute

140     Set GetActiveCustomer = objRs

    '<EhFooter>

        On Error GoTo 0

GetActiveCustomer_Exit:
        Set objCmd = Nothing
        ExitMethod "modDB", "GetActiveCustomer"
        Exit Function

GetActiveCustomer_Err:
        Dim sGetActiveCustomer_Err As String

        sGetActiveCustomer_Err = "Error: " & Err.Number & " - " & Err.Description & " (" & Erl & ")"

        ' clean up resources here
        Set objCmd = Nothing
        Set objRs = Nothing

        ExitMethod "modDB", "GetActiveCustomer"

        Err.Raise vbObjectError + 100, "SampleApp.modDB.GetActiveCustomer", sGetActiveCustomer_Err

    '</EhFooter>
End Function
```

### Invariants

1. **Marker above the signature**: `'CSEH: ErrRaise`
2. **`<EhHeader>` region** contains exactly `On Error GoTo <Func>_Err` and
   `EnterMethod "<module>", "<Func>"`
3. **Two labels in the footer**:
   - `<Func>_Exit` (success path)
   - `<Func>_Err` (error path)
4. **`<Func>_Exit` falls through to `Exit Sub`/`Exit Function`** — never to the
   `_Err` label
5. **Error message format**: `"Error: " & Err.Number & " - " & Err.Description
   & " (" & Erl & ")"`
6. **Cleanup before raise**: every `Set ... = Nothing` happens before
   `Err.Raise`, not after (after is unreachable)
7. **`ExitMethod` is called on both paths** — success and error — to keep the
   trace stack consistent
8. **`Err.Raise` source**: qualified as `<Application>.<Module>.<Procedure>`
9. **`Err.Raise` number**: `vbObjectError + <offset>`, where `<offset>` is
   chosen per procedure (see section 5)

## 3. MsgBox template

For top-level procedures where errors terminate at the UI:

```vb
'CSEH: MsgBox
Public Sub cmdSave_Click()

        '<EhHeader>
        On Error GoTo cmdSave_Click_Err

        EnterMethod "frmCustomer", "cmdSave_Click"
        '</EhHeader>

        ' ... UI logic that calls other CSEH:ErrRaise functions ...

    '<EhFooter>

cmdSave_Click_Exit:
        ExitMethod "frmCustomer", "cmdSave_Click"
        Exit Sub

cmdSave_Click_Err:
        Dim scmdSave_Click_Err As String

        scmdSave_Click_Err = Err.Description & vbCrLf & _
                             "in SampleApp.frmCustomer.cmdSave_Click at line " & Erl

        MsgBoxCritical scmdSave_Click_Err

        GoTo cmdSave_Click_Exit

    '</EhFooter>
End Sub
```

### Differences from ErrRaise

- `MsgBoxCritical` (or equivalent project wrapper) displays the error to the
  user
- No `Err.Raise` at the end — the error stops here
- `GoTo <Func>_Exit` after the dialog so cleanup and `ExitMethod` still run

## 4. Line numbering and Erl

`Erl` is a built-in function returning the line number where the last error
occurred — but only if lines are numbered. Without numbering, `Erl` returns 0
and traceability is lost.

### Numbering convention

- Start at **100**, increment by **5**
- Number only **executable** lines (no `Dim`, no comments, no labels, no
  `<EhHeader>`/`<EhFooter>` tags)
- Multi-line SQL concatenation also gets numbered
- Do not number trivial getters/setters or one-liner functions

```vb
        Dim strSQL As String

100     strSQL = ""
105     strSQL = strSQL & "SELECT idCustomer, custName " & vbNewLine
110     strSQL = strSQL & "  FROM Customer " & vbNewLine
115     strSQL = strSQL & " WHERE custActive = 1 " & vbNewLine

120     objCmd.CommandText = strSQL
125     Set objRs = objCmd.Execute
```

### Which procedures get numbered

Number any procedure with:

- `On Error GoTo` (CSEH style)
- More than ~10 executable lines
- Database access
- File or COM-object interaction

Skip numbering for:

- Property `Get`/`Let` that just reads/assigns a field
- One-liner utility functions
- Computed-only functions with no side effects and minimal logic

## 5. Variable Err.Raise offset

The error number passed to `Err.Raise` is `vbObjectError + <offset>`. The
offset is chosen per procedure (or per error category, if the codebase
catalogs them).

| Approach                          | Pattern                                               |
| --------------------------------- | ----------------------------------------------------- |
| Per-procedure unique number       | Each procedure gets its own offset                    |
| Per-category (validation, db, IO) | Range allocation (e.g., 100-199 validation, 200-299 db) |
| Fixed offset                      | Single offset across the codebase (less informative)  |

In all cases, the **qualified source string** (`SampleApp.modDB.GetActiveCustomer`)
carries the procedure identity, so the offset is supplementary. Per-procedure
numbers help when an error code is logged without the message.

## 6. On Error Resume Next — only in cleanup

`On Error Resume Next` is forbidden in domain code. It silently swallows errors
and produces invisible failures. **One legitimate exception**: cleanup and
dispose routines, where the operation must complete even when individual steps
fail.

```vb
Public Sub CloseDBConnection()
    On Error Resume Next                    ' OK: cleanup
    If Not mobjCnn Is Nothing Then
        If mobjCnn.State <> 0 Then mobjCnn.Close
    End If
    Set mobjCnn = Nothing
End Sub
```

Other acceptable uses:

- Closing file handles in an error handler
- Releasing COM objects when their state may be invalid
- Best-effort log writes (failing to log should not crash the app)

**Not acceptable:**

- Around business logic
- "To suppress a warning I do not want to deal with"
- Wrapping an entire procedure body

## 7. EnterMethod / ExitMethod placement

Every CSEH procedure calls both:

```vb
EnterMethod "modDB", "GetActiveCustomer"  ' in <EhHeader>
ExitMethod  "modDB", "GetActiveCustomer"  ' in both _Exit and _Err
```

Calling `ExitMethod` only on the success path leaves the trace stack with the
function "still open" after an error, corrupting future trace output. The
error label must also call `ExitMethod` before the `Err.Raise`.

See `vb6-trace-pattern` for the trace mechanism details.

## 8. Cleanup discipline

Resource cleanup (`Set objX = Nothing`, closing recordsets, closing files)
appears in **both** the `_Exit` and `_Err` labels. Code paths must converge on
freed resources regardless of outcome.

Common pattern: extract cleanup into a helper or duplicate the lines in both
labels. The CSEH `_Exit`/`_Err` structure does not provide an automatic
`finally` — discipline is on the author.

## 9. Anti-patterns

| Anti-pattern                                            | Why it breaks                                                            |
| ------------------------------------------------------- | ------------------------------------------------------------------------ |
| `On Error Resume Next` around a `For` loop with DB calls | Silently skips failed rows, produces partial state with no log           |
| `Err.Raise` without cleanup before it                    | Recordsets and connections leak                                          |
| `ExitMethod` only in `_Exit`                             | Trace stack grows monotonically; future traces are wrong                 |
| Unnumbered procedure with `Erl` in the handler           | `Erl` returns 0; production logs lose the line                           |
| Empty `Err.Description` re-raised                        | Caller sees no message; first failure point is unidentifiable            |
| Missing `Exit Sub`/`Exit Function` before `_Err` label   | Execution falls through into the error handler on success                |
| Modifying `<EhHeader>`/`<EhFooter>` tag format           | Breaks regeneration tools (MZ-Tools or similar)                          |

## 10. Pre-flight check

- [ ] CSEH marker present above the signature (`'CSEH: ErrRaise` or
      `'CSEH: MsgBox`)
- [ ] `<EhHeader>` and `<EhFooter>` tags present and unmodified
- [ ] `On Error GoTo <Func>_Err` is the first line inside `<EhHeader>`
- [ ] `EnterMethod` called in `<EhHeader>`
- [ ] Both `_Exit` and `_Err` labels exist
- [ ] `_Exit` falls through to `Exit Sub`/`Exit Function`, not into `_Err`
- [ ] Error message includes `Err.Number`, `Err.Description`, and `Erl`
- [ ] Cleanup is performed before `Err.Raise` (or before `MsgBoxCritical` in
      MsgBox style)
- [ ] `ExitMethod` is called in both labels
- [ ] `Err.Raise` source string is `<App>.<Module>.<Procedure>`
- [ ] Executable lines are numbered (for procedures large enough to warrant it)
- [ ] No `On Error Resume Next` outside cleanup/dispose
