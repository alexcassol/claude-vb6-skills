---
name: vb6-trace-pattern
description: Applies the manual stack-trace pattern for Visual Basic 6 based on EnterMethod/ExitMethod calls that push and pop method names from a module-level array. This pattern compensates for the absence of a native call stack in VB6, enabling rich production logs with call depth, parameter serialization, and on-demand stack inspection in error handlers. Covers when to instrument procedures (every non-trivial function), placement rules (EnterMethod in EhHeader, ExitMethod in both success and error paths), optional parameter serialization, the TraceMethod function for inserting the current stack into error messages, log file management (rolling daily logs with retention), activation gating (only logs when explicitly enabled or running in the IDE), and automatic process dumping via WMI when automation errors are detected. Activates whenever working with EnterMethod, ExitMethod, LogMsg, TraceMethod, manual stack trace, or any VB6 task involving production diagnostics, error logging, debugging, or fault investigation.
---

# VB6 Trace Pattern

VB6 has no native call stack accessible to user code. When an error fires in
production, `Err.Description` tells you **what** but not **where** in the
chain of calls. This pattern solves that by maintaining a manual stack in a
module-level array, pushed by `EnterMethod` on entry and popped by
`ExitMethod` on exit.

The pattern is one of the most useful additions you can make to a long-lived
VB6 codebase. The cost is one line at the top and one line in each exit path
of every instrumented procedure.

## 1. When to instrument

**Every non-trivial new procedure** gets `EnterMethod`/`ExitMethod`. "Non-trivial"
means any of:

- More than ~5 lines of logic
- Has `On Error GoTo` (CSEH style)
- Accesses database, file, or network
- Called from more than one place

**Skip** for:

- Trivial getters/setters (Property Get/Let that just reads/assigns a field)
- One-line utility functions
- Stateless pure helpers (string formatting, math)

When in doubt, instrument. The overhead is negligible compared to the
diagnostic value when something fails in production.

## 2. Canonical placement

The pattern is paired with the CSEH error-handling pattern:

```vb
'CSEH: ErrRaise
Public Function GetActiveCustomer(ByVal lngCustomerID As Long) As Object

        '<EhHeader>
        On Error GoTo GetActiveCustomer_Err

        EnterMethod "modDB", "GetActiveCustomer"
        '</EhHeader>

        ' ... body ...

    '<EhFooter>
        On Error GoTo 0

GetActiveCustomer_Exit:
        ' cleanup
        ExitMethod "modDB", "GetActiveCustomer"
        Exit Function

GetActiveCustomer_Err:
        Dim sGetActiveCustomer_Err As String
        sGetActiveCustomer_Err = "Error: " & Err.Number & " - " & Err.Description & " (" & Erl & ")"

        ' cleanup
        ExitMethod "modDB", "GetActiveCustomer"

        Err.Raise vbObjectError + 100, "SampleApp.modDB.GetActiveCustomer", sGetActiveCustomer_Err
    '</EhFooter>
End Function
```

**Critical:** `ExitMethod` is called on **both** the success path (`_Exit`) and
the error path (`_Err`). Skipping `ExitMethod` in `_Err` leaves the function
"open" on the stack — every subsequent trace shows that function still as a
caller, which is wrong.

## 3. EnterMethod with parameters (optional)

When the parameter values would help diagnose a future failure, pass them to
`EnterMethod` after the module and procedure names:

```vb
EnterMethod "modDB", "GetActiveCustomer", lngCustomerID, blnIncludeInactive
```

`EnterMethod` accepts a `ParamArray` and serializes each parameter:

- **Primitives** — converted via `CStr`
- **Objects** — `"Object:"` followed by a property/field dump, or `"Nothing"`
- **Recordsets** — fast dump via `GetString`
- **Arrays** — element-by-element serialization
- **Empty / Null** — literal strings

Resulting log line: `Enter: modDB::GetActiveCustomer( 1234, False )`

**Performance caveat:** serializing a populated Recordset is expensive. Inside
hot loops (thousands of iterations), do not pass heavy parameters. Pass scalar
identifiers (IDs, counts) only.

## 4. TraceMethod() in error handlers

`TraceMethod()` returns the current stack as a formatted string. Embed it in
the error message when the error is going to surface to the UI or be sent
externally (email, ticket):

```vb
GetActiveCustomer_Err:
        Dim sGetActiveCustomer_Err As String
        sGetActiveCustomer_Err = "Error: " & Err.Number & " - " & Err.Description & " (" & Erl & ")"
        sGetActiveCustomer_Err = sGetActiveCustomer_Err & vbCrLf & TraceMethod()

        ExitMethod "modDB", "GetActiveCustomer"
        Err.Raise vbObjectError + 100, "SampleApp.modDB.GetActiveCustomer", sGetActiveCustomer_Err
```

**Do not** embed `TraceMethod()` in every handler. In re-raise chains
(function A calls function B which raises; A catches and re-raises), the stack
is implicit in the chain of qualified source strings. Duplicating it in every
message bloats the log. Use `TraceMethod()` at the boundary where the error
becomes user-visible.

## 5. Log activation

Logging only writes when `gblnRegisterLog = True`. The flag is set by
`RegisterLog` (called automatically by `LogMsg`) based on three signals:

1. The IDE is active (`fg_InIDE` returns `True`) — always log in dev
2. A file named `debug.txt` exists in `App.Path` — sysadmin enables logging by
   creating this file on the client machine
3. An explicit code-level toggle (e.g., a config flag)

In production, the absence of `debug.txt` means logging is off and the
overhead of `EnterMethod`/`ExitMethod` is reduced to a couple of array
operations and an early-exit check.

## 6. Force-log for critical events

`LogMsg` accepts a `Force` parameter:

```vb
LogMsg "Database connection failed", True
```

`Force = True` writes regardless of `gblnRegisterLog`. Use sparingly — reserve
for events that must always reach the log file:

- Critical error dialogs (the project's `MsgBoxCritical` wrapper does this)
- App startup/shutdown
- Failed connection attempts
- Schema migration steps

## 7. Automatic process dump on automation errors

`LogMsg` scans the message for `"failed"` and `"automation error"`. When
either is found, it triggers `DumpProcesses` (via WMI) which logs every
running process on the machine. This is invaluable when COM/automation errors
appear sporadically and you need to know what else was running — competing
processes, antivirus interference, stale instances of the app itself, etc.

The dump fires only once per `LogMsg` call (guarded by `bDumping` flag) and
is best-effort (`On Error Resume Next` throughout).

## 8. Log file layout

Default layout for a single-tier log:

```
App.Path\
  log\
    app-2026-05-23.txt       ' general log, rotated daily
    app-2026-05-22.txt
    ...
```

`LogMsg` opens the daily file in append mode, writes a single line with
`Format(Now, "hh:nn:ss") & " < " & msg & " >"`, and closes the handle. File
handles do not stay open between messages — the cost of FreeFile/Open/Close
is acceptable in exchange for not losing entries on crash.

Old logs are pruned automatically by `CleanupOldLogs`, called once per session
on first log entry. Default retention: 30 days (`DEF_LOG_DAYS` constant).

## 9. ExitMethod modes

`ExitMethod` supports two signatures:

```vb
ExitMethod "modDB", "GetActiveCustomer"     ' specific (paired with EnterMethod)
ExitMethod "modDB"                          ' bulk (pop all frames for this module)
```

The specific form pops one frame. The bulk form pops every frame at the top
of the stack whose object matches — useful in form `Unload` handlers where
multiple methods of the same form may still be on the stack due to
exceptional early returns.

**Prefer the specific form**. The bulk form exists as a safety net but
indicates that some `ExitMethod` was missed elsewhere.

## 10. Anti-patterns

| Anti-pattern                                                | Why it breaks                                                              |
| ----------------------------------------------------------- | -------------------------------------------------------------------------- |
| `EnterMethod` without paired `ExitMethod`                   | Stack grows unbounded; every trace is wrong from that point forward        |
| `ExitMethod` only in `_Exit`, not in `_Err`                 | After an error, the function stays "open" on the stack                     |
| Instrumenting trivial getters/setters                       | Pollutes the log; cost > benefit                                           |
| Passing a populated Recordset to `EnterMethod` in a loop    | Each iteration serializes the entire recordset; massive log spam           |
| Calling `ExitMethod` before fully cleaning up resources     | Trace looks clean but resources leak silently                              |
| Using `TraceMethod()` in every re-raise handler             | The qualified source string already carries the path; duplicates inflate log |
| Initializing `msMethods()` mid-session via `Stop`/restart   | Stack becomes inconsistent; trace shows phantom frames until next entry    |

## 11. Migration notes for existing code

Adding the trace pattern to an existing codebase is incremental:

1. Drop in `modLog.bas` (from `reference/`)
2. Add `Microsoft Scripting Runtime` reference (for `Dictionary` and
   `FileSystemObject`)
3. Define `fg_InIDE` in any module (the helper that detects IDE vs compiled
   execution)
4. Start instrumenting new code immediately
5. Add to existing code opportunistically — every procedure you touch for
   another reason gets `EnterMethod`/`ExitMethod`

There is no requirement to instrument everything at once. Partial coverage is
useful: the trace shows what is instrumented, and uninstrumented procedures
appear as "missing levels" in the trace, which itself is a hint about where
to add coverage.

## 12. Reference module

`reference/modLog.bas` is a complete, self-contained implementation of the
pattern. Copy it into a new project and it works after referencing
`Microsoft Scripting Runtime` and providing a `fg_InIDE` function.

The reference module includes:

- `LogMsg` — write a line to the daily log file, with optional `Force`
- `EnterMethod` — push a method onto the stack with optional parameter
  serialization
- `ExitMethod` — pop a method (specific or bulk)
- `TraceMethod` — return the current stack as a formatted string
- `DumpProcesses` — WMI-based process listing, auto-invoked on
  automation-error patterns
- `CleanupOldLogs` — daily-log retention
- `SerializeMethod` — diagnostic helper for serializing arbitrary parameter
  arrays
