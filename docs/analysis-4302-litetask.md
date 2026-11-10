# LiteTask regression check — GLPI #4302 (Extracteur SQL failure, 2026-10-27)

Analysis date: 2026-11-09. Repository: `svtica/LiteTask`, branch `svtica-bot/magical-euler-54oes8`.
Line references are to HEAD `540fd27` (v1.2.1.12) **before** this branch's changes unless another commit is named.
Code-only analysis: no production host, `settings.xml` or full log file was accessed. Where the log excerpt is
not sufficient to decide, the section says what to look for in the full log.

> **Scope note (2026-11-10, v1.2.1.13).** The release built from this branch is *diagnostics and alerting only*: SQL failure detail (exit code, duration, stderr, stdout tail), readiness probe outcome logging, gate-deferral alerts, e-mail retry/dead-letter robustness and the hot-reload schedule fix. The readiness allow-list, cache invalidation after a failed action, the sqlcmd `-b`/`-r0` flags, `RetryCount` support and dead-letter replay described below as "fix on this branch" or "proposed" are **deferred** to a separate decision; sqlcmd arguments and the readiness classification are byte-for-byte those of 1.2.1.12 and pinned by unit tests.

## 1. Version timeline

There are no git tags; versions come from the `Bump version to x.y.z.w` commits and `LiteTask.vbproj:10-12`.

| Date (local) | Commit | Version | Relevant content |
|---|---|---|---|
| 2026-05-15 | `ff1e9b0` → `cb694e1`… | 1.2.1.8 | GUI sends `CMD_RELOAD_TASKS`; `CustomScheduler.ReloadTasks` hot reload (`CustomScheduler.vb:1363-1413`) |
| 2026-07-15 | `3b6b418` | — | **"survive SQL-unavailable startup and recover hung tasks"**: new `SqlReadinessChecker` (gate), stale-task kill/recovery, per-action `TimeoutMinutes` enforced, `ExecuteSqlTask` now fails on non-zero sqlcmd exit, scheduler tick offloaded to thread pool, stagger, startup delay |
| 2026-07-15 | `cb694e1` | **1.2.1.9** | last version built before 2026-10-28 |
| 2026-10-28 15:15 | `8eef46d` | — | credentialed Batch/Executable tasks really run through PsExec (`PsExecCommandBuilder`) |
| 2026-10-28 15:18 | `3ccab40` | — | Authenticode check on self-update |
| 2026-10-28 16:03 | `dcaa4c3` | 1.2.1.10 | |
| 2026-10-28 14:58 | `547937f` | — | SQL password via `SQLCMDPASSWORD`; **Logger rotation fixes (service `_logFolder`, write-lock deadlock)**; `SaveToolSettings` no longer erases config; temp `.sql` GUID suffix; `ToolManager` sqlcmd corruption |
| 2026-10-28 15:02 | `44791a9` | — | **e-mail retry backoff 10/20/30 s** (was immediate) |
| 2026-10-28 16:18 | `e324446` | **1.2.1.11** | deployed to production 2026-10-28 (GLPI #4265) |
| 2026-10-29 10:32 | `6bdc080` | — | PsExec progress lines classified by regex (`PsExecCommandBuilder.IsStatusLine`), also used by `ExecuteProcessWithOutput` |
| 2026-10-29 10:43 | `540fd27` | **1.2.1.12** | current HEAD |

**Version running on 2026-10-27: v1.2.1.9.** Evidence from the log excerpt alone:

- `[SqlReadiness]` probes and `Reliability settings loaded` exist only since `3b6b418` → build ≥ 1.2.1.9.
- The 8 e-mail failures are two groups of 4 in the *same millisecond* (12:01:34.046 ×4, 12:01:35.056 ×4). That is the immediate-retry loop of `NotificationManager.vb` before `44791a9` (`Continue While` right after re-enqueue). v1.2.1.11 waits 10/20/30 s between retries → build < 1.2.1.11, and 1.2.1.10 was only built on 10-28 at 16:03.
- Note: the first 10-28 commit (`8eef46d`) was authored at 15:15 local. The 10:49 catch-up run on 10-28 therefore **predates** the v1.2.1.11 deployment and must come from another service restart (see §5).

Commits touching each area (`git log --since=2026-10-10 --until=2026-11-09` plus the July base):

| Area | Commits |
|---|---|
| SqlReadiness | `3b6b418` only (file never touched since) |
| E-mail notifications | `3b6b418` (gate/stale alerts), `547937f` (`NotificationManager.vb` 4 lines), `44791a9` (backoff) |
| SQL action / sqlcmd | `3b6b418` (timeout + failure detection), `547937f` (`SQLCMDPASSWORD`, GUID temp file), `6bdc080` (stderr status filter) |
| Scheduler | `ff1e9b0`, `3b6b418` (no change since July) |
| Retries | `44791a9` (e-mail only). Action `RetryCount` never implemented (§3) |
| Locks | `61ea96c` (May, Local mutex fallback), `3b6b418` (stale recovery, backstop) |

## 2. SqlReadiness misclassification — confirmed defect

**Classification logic.** `SqlReadinessChecker.vb:79` keeps a deny-list `NetworkErrorNumbers = {-1, 2, 40, 53, 258, 1231, 10060, 10061, 11001}`. `Probe` (`:81-103`) catches `SqlException When Not NetworkErrorNumbers.Contains(ex.Number)` at `:89` and calls `MarkReady` (`:93`): **every `SqlException` whose number is not in those 9 values counts as "engine available"**. Only listed numbers and non-`SqlException` exceptions reach the "not ready" branch (`:95-101`).

Win32/SNI error numbers that the gate therefore accepts as "available" include 64 (`ERROR_NETNAME_DELETED`, the one logged at 12:30:42), 121 (semaphore timeout, the usual TCP timeout), 233, 1225 (connection refused), 10053/10054 (reset), 5, 87, and the SqlClient connect timeout `-2`. The 12:30:42 probe is exactly this path: 42 s of named-pipes fallback, error 64, `MarkReady`, then `GenRTP` was dispatched at 12:30:42.278 and killed 5 minutes later.

**Intended?** The CHANGELOG (`3b6b418`) states the intent: "Login/database errors count as 'engine up' so a permissions issue cannot block tasks forever". The intent is right; the implementation inverted it (deny-list instead of allow-list). Introduced with the feature in `3b6b418` (2026-07-15, v1.2.1.9); unchanged on v1.2.1.11 and HEAD.

**The 12:00:29 probe passed, it was not skipped.** `Probing SQL Server ...` is written only inside `Probe` (`:83`); a cache hit or a backoff skip writes nothing (`:64-70`). On success, `MarkReady` logs only when `_consecutiveFailures > 0` (`:105-108`), so a healthy probe is silent. `Task due for execution: Extracteur` follows 46 ms later, which is the normal `CheckAndExecuteTasks` flow (`CustomScheduler.vb:426-438`). The success was then cached for 300 s (`:23`, `:64-66`), so the SQL action failure 60 s later did not and could not re-probe.

Other observations on the gate:

- `IsReady()` runs the probe synchronously inside the scheduler tick under `_syncRoot` (`:54-73`); `Connect Timeout=5` (`:22`) is not honoured by the named-pipes fallback (42 s observed), so a flapping server stalls every tick.
- A task deferred by the gate only produces a per-tick log warning (`CustomScheduler.vb:427`); nothing escalates to e-mail, unlike blocked/stale tasks (`RecordSkippedRun :852-898`, `SendStaleTaskAlert :781-843`).

**Fix on this branch**: allow-list in `Class/SqlReadinessPolicy.vb` (tests `Tests/SqlReadinessPolicyTests.vb`), success/failure outcome always logged with duration and classification, `Invalidate()` after a failed SQL action, and `RecordGatedRun` alerts after 15 minutes of deferral.

## 3. SQL action failure handling

**Is sqlcmd output captured?** `ExecuteSqlCommandWithSqlCmd` (`TaskRunner.vb:482-545`) runs sqlcmd through `ExecuteProcessWithOutput` (`:400-480`). stdout is collected (`:410-417`) but logged only in `verboseMode`, which the SQL path never sets (`:527`). On a non-zero exit, `:457-462` returns **stderr only**, or the bare `Error: Process exited with code N` when stderr is empty. sqlcmd writes its error messages (including `Sqlcmd: Error: ... Login timeout expired` and `Msg nnnn, Level 16`) to **stdout** unless `-r` is passed, so the SQL error of 12:01:29 was captured, then discarded. `ExecuteSqlTask` only logs the returned string (`:577-579`).

Aggravating factor in v1.2.1.9–v1.2.1.11: the stderr handler dropped any line containing `Connecting`, `Starting`, `Copying` or `exited` as "PsExec status" (`e324446` `TaskRunner.vb:420-422`). HEAD (`6bdc080`) replaced this with an anchored regex (`PsExecCommandBuilder.vb:136-138`, used at `TaskRunner.vb:421`).

**Flags.** Arguments at `:504` / `:507`: `-S -d -I -h-1 -i` + `-U`/`-E`. **No `-b`, no `-V`, no `-r`, no `-l` (login timeout), no `-t` (query timeout).** Consequences: without `-b`, a T-SQL error inside the script (failed INSERT, constraint, deadlock victim) exits 0 and the action is reported as a success — a silent-loss case worse than the one observed. Exit code 1 without `-b` means a sqlcmd-level failure (cannot connect/login, script not found, syntax error at parse time). The 30 s between process start (12:00:58.790) and exit (12:01:29.016) is consistent with a login failure after TCP → named-pipes fallback, but this cannot be proven because the message was discarded.

**Timeouts.** Since `3b6b418`, `GetActionTimeout` (`:56-61`) applies the action's `TimeoutMinutes` (XML default 60, `XMLManager.vb:583`) and `ExecuteProcessWithOutput` kills the process tree on expiry (`:443-455`); the scheduler adds a 2-minute backstop (`CustomScheduler.vb:1089-1099`). `GenRTP` was killed after 5 minutes, so its action has `TimeoutMinutes = 5`. **Check the `TimeoutMinutes` of "Extracteur SQL" in `settings.xml`**: normal runs take up to 6 minutes; a value of 5 (ignored before July, enforced since v1.2.1.9) would kill legitimate runs. Not a code regression, but a behaviour change of `3b6b418` that production config may not reflect.

**`RetryCount`.** Persisted (`XMLManager.vb:170`, `:584`), validated (`SettingsValidator.vb:143-147`), defined (`TaskDependencyManager.vb:84-85`), **never read by `CustomScheduler` or `TaskRunner`**. Not honoured for any action type, ever.

**Retry / gate re-check after failure.** None. The action loop (`CustomScheduler.vb:1101-1125`) throws on failure (`ContinueOnError` false), the fatal path (`:1182-1218`) advances the schedule to the next occurrence (`:1189`) and the gate is only consulted before dispatch (`:426`) with a 300 s success cache. Before `3b6b418`, `ExecuteSqlTask` ignored the result entirely and returned `True` (`3b6b418^` `TaskRunner.vb:575-576`), so v1.2.1.9 is an improvement that stopped half-way.

**Other timing fact.** 9.5 s elapsed between "Created temporary SQL file" (12:00:49.260) and "Tracking child process" (12:00:58.790): `Process.Start` + `BeginOutputReadLine` (`:436-439`). Starting `sqlcmd.exe` should take milliseconds; see §5 (Logger spin in v1.2.1.9) and the environment hypothesis in §7.

**Fix on this branch**: `SqlCmdCommandBuilder` (`-b`, `-r0`, tests), `ProcessFailureFormatter` (stderr + stdout tail in the failure text, tests), sqlcmd command line logged, `SqlReadinessChecker.Invalidate` on SQL action failure. `RetryCount` and `-l/-t` are proposed, not implemented (§8).

## 4. E-mail notification failures

**Lazy init, no race.** `NotificationManager` is a DI singleton created on first `GetService` (`ApplicationContainer.vb:158-163`) under the container lock (`:141`). "Email notification system initialized successfully" (`NotificationManager.vb:97`) only means a `SmtpClient` object was constructed (`:91-95`): no connection test, and that instance is never used for sending (each attempt builds its own at `:199-206`, with `UseSSL` from settings whereas the dead instance hard-codes `EnableSsl = True`). No reuse or disposal issue.

**2 queued, 8 failures — no duplicates.** Subjects `Task Action Failed: Extracteur - ...` and `FATAL ERROR: ... - Extracteur` differ in their ` - ` prefix (`ShouldAddToBatch :274`), so two batches were flushed 5 s after queueing (`:146-162`) → two e-mails. Each is attempted once plus 3 retries (`:210-216`) = 8 failures. The timeline fits v1.2.1.9 exactly: batch A flushed at 12:01:34.03 → 4 immediate attempts at 12:01:34.046; the queue is then empty, the loop sleeps 1 s (`:226`), batch B → 4 attempts at 12:01:35.056.

**What was lost.** `Failure sending mail.` is the generic `SmtpException` text; the inner exception (DNS, socket, TLS) is never logged (`:209`). Four attempts inside one millisecond mean no network round-trip happened (name resolution negative-cached or immediate refusal). After the third retry the message is **dropped with no log line and no file** (`:210-216`, no `Else`). `Dispose` clears the queue (`:63-66`), so messages pending at a service stop are lost too. There is no persistent queue and no later retry. At 12:35:48 sends succeeded: the SMTP path was back within 34 minutes, but v1.2.1.9 gave up in 1 s and v1.2.1.11 (`44791a9`, delays at `:215`, `:220-222`) gives up after 60 s.

**Fix on this branch**: `NotificationRetryPolicy` (1/5/15 min, tests), per-message `NotBefore` so a waiting message does not block the queue, innermost exception logged, "giving up" logged, dead-letter files in `LiteTaskData\notifications\failed\` (also at dispose). Replay of dead-letter files at startup is proposed (§8).

## 5. Missed runs, 2026-10-27 18:30 → 2026-10-28 01:45

**What the 12:01:29 failure path does (and does not) leave behind.** `ExecuteTaskAsync` (`CustomScheduler.vb:1004-1256`):

- inner `Finally :1164-1172`: `_taskRunning` removed, `taskState.IsRunning = False`, stale/blocked tracking cleared;
- fatal `Catch :1182-1218`: `IsRunning = False`, `UpdateNextRunTime` + `SaveTasks` (`:1189-1190`), FATAL e-mail, rethrow;
- outer `Finally :1219-1254`: semaphore released (`:1224`) and disposed (`:1239`), `_activeMutexes` removed (`:1244`);
- `DispatchTaskAsync :459-470` logs `Error executing task ...` (`:466`, present at 12:01:29.173, which proves the whole chain ran) and clears `_pendingDispatch` (`:468`).

`UpdateNextRunTime` for a Daily task (`ScheduledTask.vb:184-185` → `:120-131`) yields the next `DailyTimes` entry after 12:01, i.e. 18:30 on 10-27, persisted by `XMLManager.SaveTask` (`:143`) and read back by `LoadTask` (`:598`). **No lock, mutex, state flag or schedule is left in a blocking state by that path.**

**Hot reload at 13:37.** `ReloadTasks` (`:1363-1413`) evicts only names missing from XML, and for retained tasks touches none of `_taskStates`, `_taskRunning`, `_pendingDispatch`, `_activeMutexes` or the OS semaphores → it **cannot corrupt the lock state of other tasks**. The GUI toggle writes only the toggled task (`MainForm.vb:633-657`, `XMLManager.SaveTask :719-765` reloads the file from disk at `:735` and replaces one node at `:747-753`), so Extracteur's XML entry (NextRunTime 18:30, Enabled) was untouched. One real defect exists though: `:1403` replaces **every** in-memory object, including a task currently executing. That run then advances the orphaned copy (`:1128` / `:1189`) while `SaveTasks` (`:1533-1564`) persists the reloaded copy with the *pre-run* `NextRunTime` → the task is **re-dispatched on the next tick** (a duplicate run, the opposite of a miss). Present since `ff1e9b0` (v1.2.1.8), still on HEAD, fixed on this branch (`AdvanceSchedule`). Not the cause here: Extracteur was idle at 13:37.

**Silent skip paths in the tick** (`CheckAndExecuteTasks :396-452`): `Not task.Enabled` (`:416`), `NextRunTime > now` (`:417`), `taskState.IsRunning` (`:420`) and `_pendingDispatch.TryAdd` failing (`:433`) all `Continue For` **without any log line**. Only the gate (`:427`) logs. So "no Extracteur line at 18:30" is compatible with several states, and the full log must decide. Signatures to grep for in the 10-27 13:37 → 10-28 10:49 window:

| Hypothesis | Code path | Log signature |
|---|---|---|
| A. Gate deferrals (SQL still flapping; a listed number such as 53/1231/2 blocks, an unlisted one such as 64 opens) | `:426-429`, `SqlReadinessChecker.vb:95-101` | `[SqlReadiness] ... not ready (probe N failed` and `Task Extracteur deferred: SQL Server is not ready yet` every 30 s; **no e-mail is ever sent for this** |
| B. Scheduler loop stalled (probe hanging inside `connection.Open()` under `_syncRoot`, thread-pool starvation) | `LiteTaskService.vb:254-275`, `SqlReadinessChecker.vb:55-73` | the `[MemoryMonitor] Working Set:` heartbeat (`CustomScheduler.vb:479-490`, every 300 s) stops; no task of any kind dispatched |
| C. Lock held by another process (a manual "Run" from the GUI uses the same `Global\LiteTask_Task_Extracteur` semaphore, `:22`, `RunTaskAsync :1479`) | `:1026-1057` | `Task due for execution: Extracteur` then `lock timeout (attempt 1)` / `Could not acquire lock`, and after 15 min `ALERT #N: Task Blocked - Extracteur` e-mails |
| D. Silent skip (`Enabled` false, bad `NextRunTime`, `IsRunning`/`_pendingDispatch` stuck) | `:416-433` | other tasks run, heartbeat present, **nothing** for Extracteur; compare `NextRunTime`/`Enabled` in the `settings.xml` backups written by `CreateBackup` (`XMLManager.vb:727`) during the evening, and look for `Task 'Extracteur' disabled` in the GUI log |
| E. Service restart/crash during the night | `InitiateServiceRestart :592-629`, `OnStart :223-295` | `[ServiceRestart]`, `LiteTaskService is starting.`, Windows System log 7031/7034; at start `Startup delay: waiting 90s` then an immediate catch-up because `LoadTask` keeps the persisted past `NextRunTime` (`XMLManager.vb:598`) |

The 10:49 catch-up on 10-28 has signature E: a restart makes the persisted 18:30 `NextRunTime` due immediately. Since no 10-28 build existed before 15:15, that restart was manual or automatic (`AutoRestartOnCriticalMemory` is on by default, `CustomScheduler.vb:60`, `:504-515`). Whatever held the task between 18:30 and 10:49 was in-memory state or a stalled loop, not the XML. A and B are the most likely given the SQL/SMTP symptoms of the afternoon; C would have produced e-mails (they worked at 12:35:48).

**Related v1.2.1.9 defect that fits the slowness.** In service mode `_logFolder` is only set by GUI paths (`ApplicationContext.vb:94`, `MainForm.vb:2167`; `Program.vb:191` runs the service without them). In `cb694e1` `Logger.vb:176-189` the rotation loop awaits `RotateLogsAsync` and `CleanupOldLogsAsync` inside the `Try` with the 1-hour delay after them; `CleanupOldLogsAsync :49-66` does `New DirectoryInfo(_logFolder)` with `Nothing` → `ArgumentNullException` → the loop retries **without any delay**, pinning one thread-pool thread at 100 % with a continuous stream of exceptions and write-lock acquisitions for as long as the service runs. Logging itself is not blocked (writes use `SyncLock _lock`, `:363-369`). Fixed in `547937f` (`Logger.vb:42-50` on HEAD). This alone does not stop the scheduler, but it is consistent with the 9.5 s `Process.Start` and the 42 s probe, and it is the strongest "fixed in v1.2.1.11" item of this review. Check `LiteTask.exe` CPU usage on the LiteTask host before/after 10-28 if the counters are still available.

## 6. Regression assessment

Status legend: **present** = exists on v1.2.1.9, v1.2.1.11 and HEAD; **fixed 1.2.1.11** = fixed by a 10-28 commit; **fixed HEAD** = fixed on 10-29; **branch** = fixed on this branch (not merged). "Introduced" is the commit that introduced the behaviour; *regression* means a previously working behaviour was broken by a recent commit.

| # | Issue | Introduced | v1.2.1.9 (10-27) | v1.2.1.11 | HEAD 540fd27 | This branch | Regression? |
|---|---|---|---|---|---|---|---|
| 1 | Readiness gate opens on unlisted network errors (64, 121, -2…) | `3b6b418` (new feature, 07-15) | present | present | present | fixed (`SqlReadinessPolicy`) | no — defect in a new feature |
| 2 | Successful probe not logged | `3b6b418` | present | present | present | fixed | no |
| 3 | Gate deferral never alerts by e-mail | `3b6b418` | present | present | present | fixed (`RecordGatedRun`) | no |
| 4 | Failed SQL action does not invalidate the 300 s cache | `3b6b418` | present | present | present | fixed (`Invalidate`) | no |
| 5 | sqlcmd error text discarded (stdout dropped on exit ≠ 0) | pre-July (`3b6b418^` `:356-359`) | present | present | present | fixed (`ProcessFailureFormatter`) | no |
| 6 | stderr lines with "Starting/Connecting/Copying/exited" swallowed | pre-July | present | present | **fixed HEAD** (`6bdc080`) | — | no |
| 7 | No `-b`: T-SQL errors exit 0 → reported success | original code | present | present | present | fixed (`-b -r0`), **behaviour change** | no |
| 8 | `RetryCount`/`RetryDelayMinutes` never honoured | original code | present | present | present | proposed | no |
| 9 | `TimeoutMinutes` now enforced (kill at expiry) | `3b6b418` | active | active | active | — | behaviour change; verify Extracteur SQL ≥ 10 min |
| 10 | E-mail retries immediate (4 in 1 ms) | original code | present | **fixed 1.2.1.11** (`44791a9`, 10/20/30 s) | fixed | longer (1/5/15 min) | no |
| 11 | Inner SMTP exception not logged | original code | present | present | present | fixed | no |
| 12 | Message dropped silently after retries; queue cleared at dispose | original code | present | present | present | fixed (dead-letter) | no |
| 13 | Hot reload replaces an executing task → duplicate run | `ff1e9b0` (05-15, v1.2.1.8) | present | present | present | fixed (`AdvanceSchedule`) | no |
| 14 | Logger rotation loop spins at 100 % CPU under the service | ≤ v1.2.1.9 | present | **fixed 1.2.1.11** (`547937f`) | fixed | — | no |
| 15 | Logger write-lock deadlock after first archive (GUI path) | ≤ v1.2.1.9 | present | **fixed 1.2.1.11** | fixed | — | no |
| 16 | `SaveToolSettings` erased config sections | ≤ v1.2.1.9 | present | **fixed 1.2.1.11** | fixed | — | no |
| 17 | Temp `.sql` name collision within the same second | ≤ v1.2.1.9 | present | **fixed 1.2.1.11** | fixed | — | no |

**No regression introduced by the 10-28/10-29 commits was found for the five problems.** None of them touches the scheduler, the gate or the e-mail queue beyond the backoff fix. One deployment risk unrelated to 10-27: `8eef46d` makes credentialed Batch/Executable actions really run under the stored account; if "Extracteur EXE" carries a credential it now runs as that account (CHANGELOG warning in `8eef46d`).

## 7. Environment hypothesis (outside LiteTask)

Four independent symptoms point to a connectivity problem on the LiteTask host between ~12:00:30 and ~12:35: sqlcmd failing after 30 s (12:00:58 → 12:01:29), SMTP failing instantly at 12:01:34-35 then working at 12:35:48, the probe answering error 64 after 42 s of named-pipes fallback at 12:30:42, and GenRTP hanging 5 minutes from 12:30:42. LiteTask did not cause it, but its failure handling (items 1, 5, 11, 12) hid the cause and the alert. Correlate with the Windows System log on the LiteTask host (NIC, DNS Client 1014, TCP/IP), FortiGate/switch logs, and the SQL Server errorlog on the SQL Server host around 12:00–12:35.

## 8. Summary table and next steps

| Issue | Status | Commit | Proposed fix |
|---|---|---|---|
| Gate misclassifies network errors (64, 121, -2) | still present on HEAD | `3b6b418` | allow-list `SqlReadinessPolicy` (branch) |
| Gate silent (no e-mail while deferring, no success log) | still present on HEAD | `3b6b418` | `RecordGatedRun` alerts + outcome logging (branch) |
| SQL error text lost; no `-b` | still present on HEAD | pre-July | `ProcessFailureFormatter`, `SqlCmdCommandBuilder -b -r0` (branch); add `-l 30` and `SqlConfiguration/LoginTimeoutSeconds` (proposed) |
| `RetryCount` ignored | still present on HEAD | original | implement in `ExecuteTaskAsync` action loop: on failure, re-run the action after `RetryDelayMinutes` up to `RetryCount` times, re-checking `IsReady()` for SQL actions, without advancing `NextRunTime` (proposed) |
| E-mail retries too short / silent drop / lost at shutdown | retry spacing fixed 1.2.1.11; rest present | `44791a9` | 1/5/15 min + dead-letter (branch); replay `notifications\failed` at startup and on a 15-min timer (proposed) |
| Hot reload during a run → duplicate run | still present on HEAD | `ff1e9b0` | `AdvanceSchedule` (branch) |
| Logger spin / deadlock, config erase, temp file collision | fixed 1.2.1.11 | `547937f` | deploy ≥ 1.2.1.11 (done 10-28) |
| Silent skip paths in the tick (`:416-433`) | present | original | add a per-task "skipped because …" Debug log and expose the reason in the GUI status (proposed) |
| `TimeoutMinutes` vs 6-min Extracteur SQL runs | config | `3b6b418` | verify `settings.xml` value ≥ 10 for "Extracteur SQL" |

Recommended next steps, in order:

1. **Pull the full LiteTask log for 2026-10-27 13:30 → 2026-10-28 11:00** and classify the night with the table in §5 (`[SqlReadiness]`, `deferred`, `[MemoryMonitor]`, `lock timeout`, `[ServiceRestart]`, `is starting`). Also check the GUI log for `Task 'Extracteur' disabled` and the `settings.xml` backups of that evening for Extracteur's `NextRunTime`/`Enabled`. Attach the findings to GLPI #4302.
2. **Verify `TimeoutMinutes` of "Extracteur SQL"** and the credential on "Extracteur EXE" (PsExec change in v1.2.1.11) in `settings.xml` — read-only check.
3. **Correlate the 12:00–12:35 window** with the LiteTask host System log, DNS, FortiGate and the SQL Server host errorlog (§7).
4. **Build and run the unit tests of this branch on a Windows machine** (`dotnet test Tests/LiteTask.Tests.vbproj`): no .NET SDK is available in this analysis environment, so the VB changes were reviewed but not compiled. Then run "Extracteur SQL" in test with the new `-b -r0` flags before any production deployment, since scripts that raise T-SQL errors will now fail the action.
5. **Deploy via the normal GLPI change** (new ticket referencing #4302 and #4265, target version 1.2.1.13), after the branch is reviewed. Do not merge from this analysis.
6. Follow-ups to schedule as separate changes: `RetryCount` support, dead-letter replay, configurable sqlcmd login timeout, and a readiness probe forced to `tcp:` when the production tasks use TCP (avoids 42 s named-pipes fallbacks inside the tick).

## 9. Files changed on this branch (superseded by the 1.2.1.13 scope note above; see the pull request for the final list)

- `Class/SqlReadinessPolicy.vb` (new), `Class/SqlReadinessChecker.vb`
- `Class/SqlCmdCommandBuilder.vb` (new), `Class/ProcessFailureFormatter.vb` (new), `Class/TaskRunner.vb`
- `Class/NotificationRetryPolicy.vb` (new), `Class/NotificationManager.vb`
- `Class/CustomScheduler.vb` (`RecordGatedRun`, `ClearGatedTracking`, `AdvanceSchedule`, cache invalidation, reload warning)
- `Tests/SqlReadinessPolicyTests.vb`, `Tests/SqlCmdCommandBuilderTests.vb`, `Tests/ProcessFailureFormatterTests.vb`, `Tests/NotificationRetryPolicyTests.vb` (new, MSTest)
- `CHANGELOG.md`, `docs/analysis-4302-litetask.md`
