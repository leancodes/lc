# VBA Style & Architecture Standard (tightened) for Office 2010

Self-contained, portable, no external dependencies. Prefer WinAPI pointers where applicable. Reliability first.

## Style Standard (tightened)

* Module prefix: `CN_`
* Procedures: PascalCase
* Variables: lowerCamel
* Constants: UPPER_SNAKE; Enums: PascalCase
* Indentation: 4 spaces; max line length 120
* Strings: explicit `CStr` on conversion
* Collections: `Scripting.Dictionary` via `CreateObject("Scripting.Dictionary")`
* Verbs: must/will; no hedging language
* Code blocks: Requirements → Code → Validation; zero inline comments
* Options: `Option Explicit`; one `.bas` only
* Compile flags: `#Const DEBUG`
* No implicit default properties
* No `Select`/`Activate`
* Public API surface: `CN_Run(taskName As String, target As Range)`

## Architecture (tightened)

* Entrypoint: `CN_Run`
* Guard: `CN_GuardVersion` `Val(Application.Version) >= 14`
* Versioning: `Const CN_VERSION As String = "1.1.0"`
* Feature gate: `Const CN_SPEC_REV As Long = 110`
* State: `CN_State` captures `Calculation`, `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, `StatusBar`
* Setup/Teardown: `CN_Begin`/`CN_End`
* Error path: `CN_Fail` (logs, restores, rethrows)
* Assertions: `CN_Require` (pre), `CN_Assert` (post; debug-gated)
* Timing: `CN_TickStart`/`CN_TickEnd` with `Timer`
* Logging: `CN_Log(level, msg)` → Immediate window; levels `INFO`, `ERROR`
* Determinism: no `ActiveWorkbook`/`ActiveSheet` dependence; explicit args only
* Tasks: `CN_TaskTrim` (Unicode whitespace normalization)

## Performance Rules (tightened)

* Toggle calc/events/screen during work
* Use arrays with `.Value2` for bulk edits
* Chunking: hard cap 100,000 cells per batch
* Use `Long` not `Integer`
* Pre-size arrays; avoid `ReDim Preserve` in loops
* Minimize `WorksheetFunction`
* Avoid `DoEvents` unless explicitly enabled (default off)

## Safety Rules (tightened)

* Validate Excel version before work
* Validate target non-`Nothing`; reject merged cells unless task supports
* Single error handler per public proc
* Always restore UI flags + `StatusBar`
* No writes outside provided range
* No volatile functions inserted
* Deterministic Unicode Trim Spec (`CN_TaskTrim`)

  * Replace NBSP `&H00A0`, Narrow NBSP `&H202F`, Figure Space `&H2007` with normal space
  * Normalize CR/LF to `vbLf`
  * Collapse consecutive whitespace to single space
  * Trim leading/trailing whitespace
  * Preserve all other Unicode characters
  * Idempotent

## Constraints & Kill-criteria (tightened)

* Abort if `Excel.Version < 14.0`
* Abort if `target.Cells.CountLarge > 5000000`
* Abort if elapsed time > 120 seconds
* Constraint: single-threaded; `DoEvents` forbidden unless explicitly enabled via `#Const ALLOW_DOEVENTS`
* Memory opaque; chunk edits ≤ 100k cells
