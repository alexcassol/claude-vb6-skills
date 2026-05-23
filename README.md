# claude-vb6-skills

Claude Code skills for maintaining and developing Visual Basic 6 codebases.
Covers conventions, error handling (CSEH), ADO via late binding (MSSQL and
PostgreSQL), manual stack tracing, and common utility patterns — calibrated
against real production VB6 code from long-lived projects.

## What's inside

Five skills, each loaded on demand when the conversation matches its scope:

| Skill                    | Activates on                                                       |
| ------------------------ | ------------------------------------------------------------------ |
| `vb6-guidelines`         | Any VB6 file or task (general conventions)                         |
| `vb6-error-handling`     | `On Error`, `Err.Raise`, `Erl`, CSEH markers, numbered lines       |
| `vb6-ado-late-binding`   | ADODB, Connection, Recordset, Command, Parameters, SQL queries     |
| `vb6-trace-pattern`      | `EnterMethod`, `ExitMethod`, `LogMsg`, `TraceMethod`, diagnostics  |
| `vb6-utilities`          | NULL handling, string ops, validation, dialog wrappers             |

Each skill ships with reference files (`reference/`) containing reusable
`.bas` modules and catalog documents that the skill points to when relevant.

## Six principles in one paragraph

1. **Preserve case** in existing code — VB6 is case-insensitive, diff tools
   are not.
2. **Hungarian notation** with scope prefix (`m_`/`m`/`g`) plus type prefix
   (`str`/`int`/`lng`/`bln`/...).
3. **CSEH** for error handling, with `EnterMethod`/`ExitMethod` on both
   success and error paths and numbered lines for `Erl`.
4. **Late-binding ADO** (`As Object` + `CreateObject`), with `ConnectionTimeout`
   mandatory, the disconnected-recordset pattern for reads, and parameters
   for every user-supplied value.
5. **Manual stack tracing** through a module-level array, pushed/popped by
   `EnterMethod`/`ExitMethod`, with optional parameter serialization and
   automatic WMI process dump on automation errors.
6. **Reuse utility wrappers** (`MsgBoxCritical`, `fg_isNull`, `IsValidCNPJ`,
   `PadLeft`, ...) — search before writing a new one.

## Installation

### Option A — install as a Claude Code plugin (recommended)

Add this repo as a marketplace, then install the plugin:

```bash
/plugin marketplace add <your-username>/claude-vb6-skills
/plugin install claude-vb6-skills@vb6-skills
```

The skills auto-activate when their triggers appear in the conversation
(opening a `.bas` or `.cls` file, writing `On Error GoTo`, building an ADO
command, etc.).

### Option B — drop CLAUDE.md into the project

For projects that don't install plugins, copy `CLAUDE.md` to the root of
the VB6 project. It is a condensed version of the same conventions and
loads automatically in every Claude Code conversation in that project.

The plugin version is richer (each skill has its own deep documentation
plus reference files); `CLAUDE.md` is a single-file summary suitable for
adding to a project repo as a permanent context file.

## Customization

These skills capture **patterns** the author has used across long-lived
codebases. Some details vary per project:

- **Multibank support**: examples show both MSSQL (`@param`) and PostgreSQL
  (`in_param`/`out_param`). If your project is single-database, ignore the
  alternative branch.
- **Module names**: examples use generic English names (`modDB`, `modUtils`).
  Your project may use Portuguese or domain-specific names — Claude will
  match the existing convention when editing.
- **Function library**: the `vb6-utilities` catalog lists common helpers
  (`fg_isNull`, `MsgBoxCritical`, `IsValidCNPJ`, etc.). Your project's
  helpers may have different names — point Claude at your actual utility
  module when relevant.
- **CSEH offset numbering**: examples use `vbObjectError + 100`. Allocate
  ranges however your team prefers (per-procedure, per-category, etc.).
- **Trace pattern**: the `modLog.bas` reference is self-contained and
  copy-paste ready, but assumes `Microsoft Scripting Runtime` is referenced
  and an `fg_InIDE()` function exists. Both are described in the skill.

Project-specific overrides go in `CLAUDE.md` (the section marked
"Project-specific overrides" at the bottom).

## Repository layout

```
claude-vb6-skills/
├── .claude-plugin/
│   ├── marketplace.json
│   └── plugin.json
├── skills/
│   ├── vb6-guidelines/
│   │   └── SKILL.md
│   ├── vb6-error-handling/
│   │   ├── SKILL.md
│   │   └── reference/cseh-styles.md
│   ├── vb6-ado-late-binding/
│   │   ├── SKILL.md
│   │   └── reference/
│   │       ├── ado-constants.bas
│   │       └── command-pattern.bas
│   ├── vb6-trace-pattern/
│   │   ├── SKILL.md
│   │   └── reference/modLog.bas
│   └── vb6-utilities/
│       ├── SKILL.md
│       └── reference/helpers-catalog.md
├── CLAUDE.md
├── EXAMPLES.md
├── README.md
└── LICENSE
```

## License

MIT — see [LICENSE](LICENSE).
