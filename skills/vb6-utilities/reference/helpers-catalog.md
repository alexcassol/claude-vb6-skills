# Quick Reference — Common VB6 Utilities

A one-page lookup. See `../SKILL.md` for full descriptions.

## NULL handling

| Function          | Returns                                                            |
| ----------------- | ------------------------------------------------------------------ |
| `fg_isNull(v, d)` | `d` if `v` is Null, else `v`                                       |
| `fg_Coalesce(...)`| First non-Null in the list                                         |
| `fg_NullIf(v, eq)`| String `"Null"` if `v = eq`, else `v` (for SQL string building)    |
| `fg_NullIfDB(v, eq)` | Actual `Null` if `v = eq`, else `v` (for ADO parameter values)  |

## SQL literal formatting (legacy — prefer parameters)

| Function                    | Returns                                       |
| --------------------------- | --------------------------------------------- |
| `ToSQL(v, [Force], [Max])`  | Variant formatted as SQL literal              |
| `SanitizeSqlString(s)`      | String with apostrophes doubled               |

## String

| Function                                | Returns                                  |
| --------------------------------------- | ---------------------------------------- |
| `RemoveAccents(s)`                      | String without diacritics                |
| `TrimAll(s)`                            | All whitespace stripped                  |
| `PadLeft(s, [c], [len])`                | Left-padded                              |
| `PadRight(s, [c], [len])`               | Right-padded                             |
| `RemoveDup(s, c)`                       | Consecutive `c` collapsed                |
| `RemoveDuplicates(s, [delim])`          | Unique entries in delimited list         |
| `RemoveEnd(s, [n])`                     | Last `n` chars removed                   |
| `MidFirst(s, c)`                        | Substring after first `c`                |
| `MidLast(s, c)`                         | Substring after last `c`                 |
| `SubstringBetween(s, start, end)`       | Between two markers                      |
| `SubstringLeft(s, end)`                 | Left of marker                           |
| `CountOccurrences(s, sub)`              | Count of `sub` in `s`                    |
| `TrimLeadingZeros(s)`                   | Leading zeros stripped                   |

## Validation

| Function                  | Returns                                  |
| ------------------------- | ---------------------------------------- |
| `IsValidCNPJ(s)`          | True if valid CNPJ (BR company tax ID)   |
| `IsValidCPF(s)`           | True if valid CPF (BR personal tax ID)   |
| `IsValidEmail(s)`         | True if valid email format               |
| `IsNumericPad(keyCode)`   | True if key is numeric-pad-compatible    |

## Numeric

| Function                                          | Returns                          |
| ------------------------------------------------- | -------------------------------- |
| `Trunc(v, places)`                                | Truncated to N decimals          |
| `RoundOrTrunc(v, places)`                         | Round or truncate per config     |
| `ApplyFiscalRounding(v, [places])`                | Fiscal rounding                  |
| `ToHour(sec, [IncludeSeconds])`                   | Seconds → HH:MM[:SS]             |

## Dialogs (use these — never raw MsgBox)

| Function                                          | Behavior                                   |
| ------------------------------------------------- | ------------------------------------------ |
| `MsgBoxCritical(text)`                            | Critical error, always logged              |
| `MsgBoxInfo(text)`                                | Info dialog                                |
| `MsgBoxWarning(text, [ret], [yesNo])`             | Warning, OK or Yes/No                      |
| `MsgBoxQuestion(text, [default], [yesNo], ...)`   | Yes/No confirmation                        |
| `MsgBoxOption(text, opt1, opt2, ...)`             | Two-button custom choice                   |

## Environment

| Function       | Returns                              |
| -------------- | ------------------------------------ |
| `fg_InIDE()`   | True if running in IDE, else False   |

## UI

| Function                                          | Behavior                          |
| ------------------------------------------------- | --------------------------------- |
| `CenterObject(o)`                                 | Center form/control               |
| `FocusOn(ctl, [select])`                          | Set focus, optionally select text |
| `PopulateCombo(cbo, rs, [defaultID])`             | Fill ComboBox from Recordset      |
| `LV_ColumnSort(lv, ...)`                          | ListView column sort              |
| `ListViewLineColor(lv, idx, color)`               | Set row color                     |
| `ListViewToJson(lv, [useKey])`                    | Export to JSON                    |
| `LoadListViewFromJson(lv, json)`                  | Import from JSON                  |

## Serialization (for logging)

| Function                              | Returns                              |
| ------------------------------------- | ------------------------------------ |
| `Serialize(fields)`                   | Generic serializer                   |
| `SerializeForm(frm, [readable])`      | Form state as string                 |
| `SerializeListItem(lv, item)`         | Single ListView row                  |
| `SerializeListGrid(grid, [row])`      | Grid row or all rows                 |

## File system

| Function              | Behavior                            |
| --------------------- | ----------------------------------- |
| `CreateFolder(dir)`   | Recursive mkdir                     |

## Date

| Function                       | Returns                          |
| ------------------------------ | -------------------------------- |
| `LastDayOfMonth(iMonth, iYear)` | Last day of month               |
