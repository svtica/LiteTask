# Changelog

All notable changes to this project are documented in this file.

## [Unreleased]

### Security
- **Self-update now verifies the Authenticode signature of the downloaded package before anything from it runs** (`UpdateManager.DownloadAndApplyUpdateAsync`). After extraction, `LiteTask.exe` is located with the same rules as the file-replacement script (archive root, single sub-folder, first match) and checked with `WinVerifyTrust` (new `AuthenticodeVerifier`): the signature must match the file content and the signer thumbprint must equal the pinned publisher certificate `res/certificate.cer` (CN=SVTI.ca, SHA-1 `EC2B412C4025C2F1C8DB512CEC66632E65C017F0`, the certificate `Release.yml` signs with) or the certificate that signed the running `LiteTask.exe`. Because that certificate is self-signed, `CERT_E_UNTRUSTEDROOT` counts as a verified digest and identity is enforced by the thumbprint pin instead of the machine's root store. Policy (`UpdateSignaturePolicy`):
  - signed by the expected publisher: applied. The `LiteTask-Post-Build.ps1` shipped in the archive (executed with `-ExecutionPolicy Bypass` before files are replaced) must then also be validly signed by the same publisher, otherwise the update is rejected;
  - unsigned package: rejected, **unless the running `LiteTask.exe` is itself unsigned** (development build), in which case it is applied with a visible warning so a dev build can still update when the CI signing secret is absent;
  - tampered (digest mismatch), invalid, or signed by another certificate: rejected.
  A rejected package stops before `RunPostBuildScriptAsync`, logs the reason, shows it in `UpdateProgressForm` and deletes the temporary archive and extracted folder. New `Update.Progress.*` strings with French translations in `LiteTaskData/lang/fr.xml`.
- `Release.yml` now also signs the `*.ps1` files at the package root (`LiteTask-Post-Build.ps1`, `Unblock-Files.ps1`) when the signing certificate is present, so signed releases satisfy the script check above. Rotating the signing certificate requires updating `UpdateManager.ExpectedPublisherThumbprint` in a release still signed with the previous certificate.
- Update download URL must use `https://`; the release asset name is reduced to its file name before being used as the local archive name.

### Fixed
- **Startup race after a full-server reboot silently disabling a task for days** (GenRTP incident, 2026-07-10 → 2026-07-15). Root causes addressed:
  - Per-action `TimeoutMinutes` (persisted since day one but never enforced) is now applied to every execution path: `Executable`/`Batch` child processes and `sqlcmd` run through `RunProcessWithTimeout`, which kills the entire process tree on expiry; the PowerShell in-process timeout now uses the action's configured value instead of a fixed 4 h. A scheduler-level backstop (`action timeout + 2 min`) additionally kills the tracked child processes and releases the task lock if the execution path itself hangs (e.g. output streams never closing).
  - Stale-task recovery now actually recovers: `TaskRunner` tracks child processes by PID per task (`KillProcessesForTask`), and `CleanupStaleTasks` no longer resets the running state at 15 minutes (which caused endless re-dispatch into a held lock); instead it alerts, then at `sum of action timeouts + 10 min` it kills the owning process(es), force-releases the lock, and resets the state so the next schedule fires.
  - Stale/blocked alerts now repeat every `Reliability/StaleAlertRepeatMinutes` (default 60) with an escalating `ALERT #N` counter while the condition persists; a task that keeps being skipped (lock unavailable) also raises repeating alerts instead of logging ~6,700 silent warnings.
  - `ExecuteSqlTask` no longer reports success when `sqlcmd` fails or times out.
- Scheduler tick no longer blocks for up to 30 s per task on lock acquisition: dispatch is offloaded to the thread pool with a pending-dispatch guard against double launch.

### Added
- SQL Server readiness gate (`SqlReadinessChecker`): before dispatching SQL-dependent tasks (any task with a SQL action, plus names matching `Reliability/SqlGatedTasks`, default `GenRTP,Extracteur`), the scheduler probes SQL Server with a 5 s connection test. While unavailable, the task is deferred to the next tick (schedule untouched) and probes retry with exponential backoff (15 s → 10 min cap), each attempt logged. Login/database errors count as "engine up" so a permissions issue cannot block tasks forever. Probe server: `Reliability/SqlReadinessServer`, falling back to `SqlConfiguration/DefaultServer`.
- Configurable post-reboot startup grace: the service waits `Reliability/StartupDelaySeconds` (default 90) before the first dispatch, and simultaneous due tasks are staggered by `Reliability/TaskLaunchStaggerSeconds` (default 5) instead of all launching in the same second.
- Non-interactive child execution: the service calls `SetErrorMode(SEM_FAILCRITICALERRORS Or SEM_NOGPFAULTERRORBOX Or SEM_NOOPENFILEERRORBOX)` at startup (inherited by children) so Windows hard-error dialogs cannot block a child invisibly, and the PsExec `-i` flag is only used in interactive sessions.
- New `Reliability` settings section (all keys optional, defaults apply): `StartupDelaySeconds`, `TaskLaunchStaggerSeconds`, `SqlReadinessEnabled`, `SqlReadinessServer`, `SqlGatedTasks`, `StaleAlertRepeatMinutes`.
- Task deletion not persisting when `LiteTaskService` is running: the service process keeps its own in-memory `_tasks` dictionary and re-writes the full set to `settings.xml` at the end of every execution (via `SaveTasks`). Because the GUI mutated XML directly without telling the service, the service's stale copy resurrected just-deleted tasks (and missed adds/edits) until the next service restart. The GUI now sends `LiteTaskService.CMD_RELOAD_TASKS` (custom service control 128) via `ServiceController.ExecuteCommand` after every task mutation (delete, create, edit, toggle-enable, import), which routes through `LiteTaskService.OnCustomCommand` → `CustomScheduler.ReloadTasks`. `ReloadTasks` evicts entries no longer present in XML (skipping anything currently executing) and refreshes the rest. The notification is best-effort and logs a warning when the GUI user lacks `SERVICE_USER_DEFINED_CONTROL` on the service.
- Task deletion not persisting: `XMLManager.DeleteTask` only searched the canonical `LiteTaskSettings/Tasks/Task` path while `GetAllTaskNames`/`LoadTask` accept three layouts, so non-canonical entries survived deletion. `CustomScheduler.SaveTasks` then re-merged the surviving XML entry back into `_tasks` via `GetAllTasks`, resurrecting the task. `DeleteTask` now removes every matching node across supported layouts (under the existing write lock, with atomic save) and `SaveTasks` treats the in-memory `_tasks` dictionary as the source of truth.

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
