---
name: vb6-utilities
description: Catalog of common utility functions that VB6 codebases typically maintain in a shared module (often named modUtils, modFuncoes, modCommon, or similar). Activates when writing or editing VB6 code that handles NULL coalescing, SQL value formatting, string sanitization for SQL literals, accent/diacritic removal for searching, padding (left/right), substring extraction between markers, CNPJ/CPF (Brazilian tax ID) validation, MsgBox dialog wrappers, IDE-vs-compiled detection, regex via VBScript.RegExp, ADO Command parameter helpers, ListView and grid serialization, and common helpers prefixed with fg_ (function global). The goal is to recognize these patterns and reuse existing project utilities instead of reinventing them. Triggers on tasks involving string manipulation, validation, NULL handling, formatting, or any common helper that likely already exists in the project's utility module.
---

# VB6 Utilities Catalog

VB6 codebases typically maintain a shared utility module (commonly named
`modUtils`, `modFuncoes`, `modCommon`, or similar) that accumulates helpers
over years of development. Before writing a utility function from scratch,
check whether one already exists.

This catalog describes the functions you are most likely to find. **Names and
exact signatures vary per project**; the names below are representative of
the patterns to look for.

## Conventions

### `fg_` prefix — "function global"

Many codebases use the `fg_` prefix for stateless utility functions exposed
project-wide. Examples: `fg_isNull`, `fg_Coalesce`, `fg_InIDE`. When you see
this prefix, treat the function as part of the utility layer — reusable,
side-effect-free, safe to call from anywhere.

### Dialog wrappers — `MsgBoxXxx`

Most projects wrap `MsgBox` to integrate with logging, theming, and consistent
button layouts. Typical names: `MsgBoxCritical`, `MsgBoxInfo`, `MsgBoxWarning`,
`MsgBoxQuestion`, `MsgBoxOption`. **Use the wrapper, not raw `MsgBox`**.

## 1. NULL handling

### `fg_isNull(value, defaultValue) As Variant`

Returns `defaultValue` when `value` is `Null`, otherwise returns `value`.
Equivalent to SQL's `ISNULL` or `COALESCE` with one fallback.

```vb
strName = fg_isNull(rs("custName").Value, "")
lngCount = fg_isNull(rs("total").Value, 0)
```

### `fg_Coalesce(value1, value2, ...) As Variant`

Returns the first non-`Null` value in the list, or `Null` if all are `Null`.

```vb
strContact = fg_Coalesce(strPhone, strMobile, strEmail, "no contact")
```

### `fg_NullIf(value, valueIfEqual) As Variant`

Returns the string `"Null"` (for SQL literal use) when `value` matches
`valueIfEqual`, otherwise returns `value`. Used when building SQL strings
manually.

### `fg_NullIfDB(value, valueIfEqual) As Variant`

Returns actual `Null` (not the string `"Null"`) when `value` matches
`valueIfEqual`, otherwise returns `value`. Used when assigning to ADO
parameter values.

## 2. SQL literal formatting (legacy)

### `ToSQL(value, [ForceString], [MaxChars]) As String`

Formats any Variant as a SQL literal: dates as `'dd/mm/yyyy'`, numbers
with `.` decimal separator, strings with apostrophes doubled, `Null` as
`Null`.

```vb
strSQL = "INSERT INTO Customer (custName, custBirth) VALUES (" & _
         ToSQL(strName) & ", " & ToSQL(dtmBirth) & ")"
```

**Modern code prefers parameters over `ToSQL`.** See `vb6-ado-late-binding`.
`ToSQL` remains useful for:

- Building diagnostic display SQL (e.g., `ShowADOCommand` output)
- Structural SQL where a single literal is concatenated for legitimate reasons
- Maintenance of existing code that uses it

### `SanitizeSqlString(s As String) As String`

Doubles apostrophes (`'` → `''`) to prevent SQL literal breakage. Called
internally by `ToSQL`. Direct use is mainly for backward compatibility with
older code paths.

**Note:** sanitization is a fallback. Parameters are the correct solution.

## 3. String manipulation

### `RemoveAccents(sText As String) As String`

Strips diacritical marks (á → a, ç → c, etc.) for accent-insensitive search
and matching.

### `TrimAll(sText As String) As String`

Strips all whitespace including non-breaking spaces, tabs, and CRLF — beyond
what built-in `Trim` does.

### `TrimLeadingZeros(value As String) As String`

Removes leading zeros from a numeric string. `"00012345"` → `"12345"`.

### `PadLeft(sText, [sChar], [iLen]) As String`
### `PadRight(sText, [sChar], [iLen]) As String`

String padding to a target length with a fill character (default space).

```vb
strInvoiceNum = PadLeft(CStr(lngID), "0", 9)    ' "000012345"
```

### `RemoveDup(sText, sChar) As String`

Collapses consecutive occurrences of `sChar` into one. `"a,,b,,,c"` → `"a,b,c"`.

### `RemoveDuplicates(pText, [Delimiter]) As String`

Removes duplicate entries from a delimited list. `"a,b,a,c,b"` → `"a,b,c"`.

### `RemoveEnd(sText, [iQtChar]) As String`

Removes the last N characters. Useful for trailing separators after a loop:
`PathBuild = PathBuild & dir & "\"` then `RemoveEnd(PathBuild, 1)`.

### `MidFirst(sText, sChar) As String`

Returns the substring after the first occurrence of `sChar`.

### `MidLast(sText, sChar) As String`

Returns the substring after the last occurrence of `sChar`.

### `SubstringBetween(sString, markStart, markEnd, [StartEnd]) As String`

Returns the substring between two markers. `SubstringBetween("Hello [World]!", "[", "]")` → `"World"`.

### `SubstringLeft(sString, markEnd) As String`

Returns the substring left of the marker.

### `CountOccurrences(sText, sStringToCount) As Integer`

Counts non-overlapping occurrences of a substring.

### `ReplaceBetween(Expression, FindIni, FindEnd) As String`

Replaces everything between two markers (inclusive of the markers).

## 4. Validation

### `IsValidCNPJ(CNPJ As String) As Boolean`

Validates a Brazilian company tax ID (CNPJ): format, length, and check
digits. Accepts formatted (`"12.345.678/0001-90"`) or unformatted
(`"12345678000190"`) input.

### `IsValidCPF(CPF As String) As Boolean`

Validates a Brazilian personal tax ID (CPF): format, length, and check
digits.

### `IsValidEmail(mailAddress As String) As Boolean`

Validates email address format via regex.

### `IsNumericPad(KeyCode As Integer) As Boolean`

Returns `True` if the key code corresponds to a numeric input (0-9, decimal,
backspace, arrows). Used in `KeyDown`/`KeyPress` handlers for numeric-only
text fields.

## 5. Format conversion

### `ToHour(sec As Single, [IncludeSeconds]) As String`

Converts seconds to `HH:MM:SS` or `HH:MM` format.

### `fg_NumberToWords(nValue As Double, CurrencySingular, CurrencyPlural) As String`

Renders a numeric value as written-out currency
("one hundred and twenty-three reais and forty-five cents").

### `Trunc(value, DecimalPlaces As Byte) As Currency`

Truncates a number to N decimal places (no rounding).

### `RoundOrTrunc(value, DecimalPlaces As Byte) As Currency`

Truncates or rounds based on a project-wide setting
(`gParameters.RoundOrTruncate`).

### `ApplyFiscalRounding(dblTotalValue, [decimalPlaces]) As Double`

Applies fiscal rounding rules (banker's rounding or specific Brazilian fiscal
conventions).

## 6. Encryption (lightweight, not for security)

### `Encrypt(Input As String) As String`
### `Decrypt(Input As String) As String`

Symmetric scrambling for storing low-stakes values (window position, recent
file list) in INI files. **Not cryptographically secure** — do not use for
passwords or sensitive data.

### `fg_EncryptKey(sKey As String) As String`

Project-specific key encoding for connection strings stored in config files.

## 7. Dialog wrappers

### `MsgBoxCritical(Text As String)`

Critical error dialog. Always logged with `Force = True`.

### `MsgBoxInfo(Text As String)`

Informational dialog. Logged at normal level.

### `MsgBoxWarning(Text As String, [returnValue], [YesNoButtons]) As VbMsgBoxResult`

Warning dialog with optional OK-only or Yes/No buttons.

### `MsgBoxQuestion(Text, [DefaultButton], [YesNoButtons], ...) As VbMsgBoxResult`

Yes/No confirmation dialog.

### `MsgBoxOption(Text, Option1, Option2, ...) As Byte`

Two-button choice dialog with custom button captions.

**Rule:** new code never calls `MsgBox` directly. Always use the wrapper.

## 8. Environment detection

### `fg_InIDE() As Boolean`

Returns `True` when running inside the VB6 IDE, `False` when running as a
compiled EXE. Classic implementation triggers `Debug.Print 1/0` which raises
in IDE but is removed by the compiler.

```vb
Public Function fg_InIDE() As Boolean
    On Error GoTo erro
    Static blnCached As Boolean
    Static blnComputed As Boolean
    
    If blnComputed Then
        fg_InIDE = blnCached
        Exit Function
    End If
    
    Debug.Print 1 / 0           ' raises in IDE, removed by compiler
    blnCached = False
    blnComputed = True
    fg_InIDE = False
    Exit Function
erro:
    blnCached = True
    blnComputed = True
    fg_InIDE = True
End Function
```

## 9. UI helpers

### `CenterObject(o As Object)`

Centers a form or control within its container.

### `FocusOn(ctl As Object, [SelectText As Boolean])`

Sets focus to a control with optional auto-selection of contents.

### `PopulateCombo(cbo As ComboBox, rs As Object, [defaultID])`

Populates a ComboBox from a Recordset (column 0 as value, column 1 as
display).

### `LV_ColumnSort(ListView, ...)`

Generic ListView column sort handler.

### `ListViewLineColor(lvw, ItemIndex, ToColor)`

Sets a row's background color in a ListView.

### `ListViewToJson(lvw, [useItemKey]) As String`

Exports a ListView's contents as JSON for persistence or transport.

### `LoadListViewFromJson(lvw, sJSON As String)`

Imports a ListView from a JSON string previously produced by `ListViewToJson`.

## 10. Serialization (for logging/debugging)

### `Serialize(Fields As Variant) As String`

Generic serializer that handles arrays, Dictionaries, Recordsets, and Field
collections. Returns a delimited string for logging.

### `SerializeForm(frm As Object, [HumanReadable]) As String`

Iterates a form's controls and serializes their current values. Useful in
error handlers to capture the UI state at the moment of failure.

### `SerializeListItem(lstView, lstItem) As String`
### `SerializeListGrid(lstView, [lRow]) As String`

ListView and grid row serialization.

## 11. Date helpers

### `LastDayOfMonth(iMonth, iYear) As Integer`

Returns the last day of a given month.

## 12. File system

### `CreateFolder(DestDir As String) As Boolean`

Creates a directory tree (recursive), returning `True` on success.

## 13. Regex

### `RegExp(sPattern As String, sText As String, ...) As ...`

Wraps `VBScript.RegExp` (late-bound) for regex matching. Useful when the
project does not want a reference to the Regex type library.

## 14. Other

### `EvaluateVBCode(Scode As String)`

Late-bound `MSScriptControl.ScriptControl` evaluation. Use cautiously — this
is `eval` for VB6 and carries the same risks.

### `GetSeqNumber() As String`

Returns a project-wide unique sequence number (timestamp-based or
counter-based depending on implementation).

### `Soundex(sText As String) As String`

Phonetic encoding for fuzzy name matching.

---

## Pre-flight check before writing a new utility

Before adding a new helper function, look for an existing one:

- [ ] Searched for `fg_<thing>` in the utility module
- [ ] Searched for `<Thing>` (PascalCase) in the utility module
- [ ] Searched for the operation by description (e.g., "remove accent",
      "validate cnpj", "format hour")
- [ ] Asked whether the codebase has a consistent place for this kind of
      helper (`modUtils`, `modFuncoes`, `modString`, etc.)

If the helper does not exist, follow the project's naming conventions
(`fg_` prefix or not, English or Portuguese, the file the helper belongs in)
when adding it.
