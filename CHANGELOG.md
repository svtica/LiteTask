# Changelog

All notable changes to this project are documented in this file.

## [Unreleased]

### Changed
- `SettingsValidator`: removed stale `TODO Reimplement in Options` comment; the validator is already wired into `OptionsForm.ValidateSettings`.
- `TaskRunner.ParseParameters`: now delegates to a new testable `LiteTask.ParameterParser` module that supports cmd-style quoting (`key="value with spaces"`) while remaining backward-compatible with `key=value`.
- `CustomScheduler` / `LiteTaskService`: scheduler tick rate reduced from 60 s to 30 s to support minute-level interval recurrences with acceptable jitter.
- `XMLManager.ParseInterval`: tolerates both the current "minutes as double" `<Interval>` format and legacy TimeSpan strings (e.g. `00:15:00`); writes use `InvariantCulture` to avoid locale-dependent decimals.
- `TaskForm`: `Interval` recurrence input is now bounded to 1-1440 whole minutes with friendly validation; existing tasks load and round to the nearest minute.

### Added
- `SettingsValidator.ValidateTaskAction(action[, siblings])`: validates name/target presence, target file existence for PowerShell/Batch/Executable types, non-negative retry parameters, timeout >= 1, and `DependsOn` referencing a sibling action's `Name`.
- `ActionDialog` and `TaskForm` now run candidate actions through `ValidateTaskAction` on save; cross-action `DependsOn` consistency is checked at the task level.
- `Tests/LiteTask.Tests.vbproj`: MSTest project covering `ParameterParser` (11 cases: basic form, quoted spaces, mixed quoting, empty values, dropped tokens, duplicate keys, whitespace tolerance).
- French translations for the new `Validation.*` keys in `LiteTaskData/lang/fr.xml`.
- `MainForm` View menu: new `Open Log File` item that opens the current log (resolved via `Logger._logFile`) in the default associated application; shows a friendly `NoLogAvailable` message when no log has been written yet.
- Sub-hour interval recurrence: scheduling a task to run every N minutes (1-1440) is now first-class - parity with the Windows Task Scheduler "Repeat task every" setting.
- `ScheduledTask.CalculateNextRunTime` / `UpdateNextRunTime`: defensive guards when `Interval` is zero (returns `DateTime.MaxValue` instead of dividing by zero or looping forever).
