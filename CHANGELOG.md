# Changelog

All notable changes to this project are documented in this file.

## [Unreleased]

### Changed
- `SettingsValidator`: removed stale `TODO Reimplement in Options` comment; the validator is already wired into `OptionsForm.ValidateSettings`.
- `TaskRunner.ParseParameters`: now delegates to a new testable `LiteTask.ParameterParser` module that supports cmd-style quoting (`key="value with spaces"`) while remaining backward-compatible with `key=value`.

### Added
- `SettingsValidator.ValidateTaskAction(action[, siblings])`: validates name/target presence, target file existence for PowerShell/Batch/Executable types, non-negative retry parameters, timeout >= 1, and `DependsOn` referencing a sibling action's `Name`.
- `ActionDialog` and `TaskForm` now run candidate actions through `ValidateTaskAction` on save; cross-action `DependsOn` consistency is checked at the task level.
- `Tests/LiteTask.Tests.vbproj`: MSTest project covering `ParameterParser` (11 cases: basic form, quoted spaces, mixed quoting, empty values, dropped tokens, duplicate keys, whitespace tolerance).
- French translations for the new `Validation.*` keys in `LiteTaskData/lang/fr.xml`.
