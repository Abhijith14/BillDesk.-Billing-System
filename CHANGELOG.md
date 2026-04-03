# Changelog

## v1.2.0 — Stability Overhaul, Bug Fixes & UI Improvements (2026-04-03)

### Critical Bug Fixes

#### White Screen / Forms Disappearing — FIXED
- **Root cause:** `Form1` was set as `IsMdiContainer = True` but child forms were standalone windows, not MDI children. A white `GroupBox2` (2158x1447px) sat on top in z-order and covered child forms when they lost focus.
- **Fix:** Removed `IsMdiContainer`. Replaced `GroupBox2` with a `Panel` (`contentPanel`). Child forms are now embedded inside the panel using `TopLevel = False` and `Dock = DockStyle.Fill`. Forms can never be covered or lost.

#### App Exit Flickering / Blinking — FIXED
- **Root cause:** `Form6` (splash screen) was the `MainForm` but only hid itself, never closed. When `Form1` closed, the framework's shutdown sequence conflicted with the still-alive hidden main form, causing multi-stage flickering.
- **Fix:** Changed `MainForm` from `Form6` to `Form1`. Moved all initialization logic into `Form1_Load`. `Form6` is now a stub.

### Security Fixes

#### SQL Injection Vulnerabilities — FIXED (3 locations)
- `Form2.vb` search by enrollment number — was using string concatenation
- `Form2.vb` search by student name — was using string concatenation
- `Form5.vb` invoice lookup — was using string concatenation
- **Fix:** All three now use `OleDbParameter` with `?` placeholders.

### Resource Leak Fixes

#### OleDB Connection Leaks — FIXED (9 locations)
All database connections across `Form2`, `Form3`, `Form4`, and `Form5` now use `Using` blocks, ensuring connections are properly closed and disposed even when exceptions occur.

| File | Functions Fixed |
|------|----------------|
| Form2.vb | `load_data()`, `TextBox1_TextChanged()`, `TextBox2_TextChanged()` |
| Form3.vb | `CalculateDue()` |
| Form5.vb | `UpdateData()`, `return_length()`, `FillDetails()`, `FillSearchDetails()`, `SaveExcel()` |

#### Excel COM Object Leaks — FIXED
`Form2.SaveData()` now uses `Try/Finally` with proper null checks and `Marshal.ReleaseComObject` in reverse order, followed by a single `GC.Collect()` + `GC.WaitForPendingFinalizers()`.

### Code Quality Fixes

- **Empty catch block** in `Form4.Button1_Click` — now shows error message instead of silently swallowing
- **numtoword.vb** — Added null safety check in `GetRootNumberWord()` to prevent crash on unmapped enum values
- **Debug/Release platform mismatch** — Debug config changed from `AnyCPU` to `x64` to match Release and the 64-bit ACE.OLEDB provider
- **Dead code removed** — `UserControl1` (unused), empty event handlers, commented-out code blocks

### Improvements

- **DbHelper connection string caching** — Registry is checked once; subsequent calls reuse the cached provider string
- **Double buffering** — Enabled on Form1-5 to reduce visual flicker during transitions
- **Form lifecycle** — All child forms are created once and reused, managed by `ShowChildForm()` in Form1

### UI/UX Improvements

- **Active navigation indicator** — The currently selected sidebar button is highlighted in blue with white text
- **Button hover effects** — Navigation buttons show a light blue hover state for visual feedback
- **Flat modern button style** — Sidebar buttons use `FlatStyle.Flat` with subtle borders
- **Exit confirmation** — Logout now asks "Are you sure you want to exit?" before closing
- **Dashboard date context** — Shows "As of [today's date]" on the dashboard
- **Loading cursor** — Student data reload shows a wait cursor during load
- **Improved message boxes** — Dialogs now include proper titles ("BillDesk") and icons

### Files Modified
1. `Form1.vb` — Complete rewrite (form hosting, UI, nav highlighting)
2. `Form1.Designer.vb` — Panel replaces GroupBox2, removed MDI
3. `Form2.vb` — Connection leaks, SQL injection, COM cleanup, UI
4. `Form3.vb` — Connection leak, dashboard date display
5. `Form4.vb` — Error handling, cleanup
6. `Form5.vb` — Connection leaks, SQL injection, double buffering
7. `Form6.vb` — Stripped to stub (no longer main form)
8. `DbHelper.vb` — Connection string caching
9. `numtoword.vb` — Null safety check
10. `My Project/Application.myapp` — MainForm changed to Form1
11. `Fees_Management.vbproj` — Debug platform to x64, removed UserControl1

### Files Removed
- `UserControl1.vb` — Unused empty user control
- `UserControl1.Designer.vb`
- `UserControl1.resx`
