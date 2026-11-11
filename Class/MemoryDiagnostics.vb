' -----------------------------------------------------------------------------
' Copyright (c) svtica. All rights reserved.
' File:    MemoryDiagnostics.vb
' Author:  LiteTask contributors
' Date:    2026-11-11
' Purpose: Measure before removing. LiteTask forces full garbage collections after
'          every task (one of them blocking, with WaitForPendingFinalizers, after
'          every PowerShell task) and restarts itself on high memory. These helpers
'          log the memory each run leaves behind (private bytes, GC heap, handles),
'          so the task that grows memory can be identified, and make the forced
'          collections a setting (Reliability/ForceGcAfterTask) so the service can
'          be compared with and without them. Formatting is pure for unit tests.
' -----------------------------------------------------------------------------
Imports System.Globalization

Namespace LiteTask

    ''' <summary>Process memory at one point in time.</summary>
    Public Structure MemorySnapshot
        Public PrivateBytes As Long
        Public WorkingSet As Long
        Public GcHeapBytes As Long
        Public HandleCount As Integer
        Public Gen2Collections As Integer

        Public Shared Function Capture() As MemorySnapshot
            Dim snapshot As New MemorySnapshot With {
                .GcHeapBytes = GC.GetTotalMemory(False),
                .Gen2Collections = GC.CollectionCount(2)
            }
            Try
                Using current = Process.GetCurrentProcess()
                    snapshot.PrivateBytes = current.PrivateMemorySize64
                    snapshot.WorkingSet = current.WorkingSet64
                    snapshot.HandleCount = current.HandleCount
                End Using
            Catch
                ' Counters unavailable: keep the GC values only
            End Try
            Return snapshot
        End Function
    End Structure

    Public NotInheritable Class MemoryDiagnostics
        Private Sub New()
        End Sub

        ''' <summary>
        ''' Forced full garbage collections after each task (Reliability/ForceGcAfterTask,
        ''' True by default = previous behaviour). Set by CustomScheduler.LoadReliabilitySettings.
        ''' </summary>
        Public Shared Property ForceGcAfterTask As Boolean = True

        Private Shared Function Mb(bytes As Long) As String
            Return (bytes / 1048576.0R).ToString("0.0", CultureInfo.InvariantCulture)
        End Function

        Private Shared Function SignedMb(bytes As Long) As String
            Return If(bytes >= 0, "+", "") & Mb(bytes)
        End Function

        ''' <summary>
        ''' "private 412.3 MB (+12.5), GC heap 85.1 MB (+3.0), handles 1204 (+6), gen2 GCs +1".
        ''' </summary>
        Public Shared Function FormatDelta(before As MemorySnapshot, after As MemorySnapshot) As String
            Dim handleDelta = after.HandleCount - before.HandleCount
            Return $"private {Mb(after.PrivateBytes)} MB ({SignedMb(after.PrivateBytes - before.PrivateBytes)}), " &
                   $"GC heap {Mb(after.GcHeapBytes)} MB ({SignedMb(after.GcHeapBytes - before.GcHeapBytes)}), " &
                   $"handles {after.HandleCount} ({If(handleDelta >= 0, "+", "")}{handleDelta}), " &
                   $"gen2 GCs +{after.Gen2Collections - before.Gen2Collections}"
        End Function

        ''' <summary>
        ''' Full blocking collection used after a PowerShell task (COM objects created by
        ''' scripts, e.g. Excel automation, are only released once their wrappers are
        ''' finalized). No-op when ForceGcAfterTask is off.
        ''' </summary>
        Public Shared Sub CollectAfterPowerShellTask()
            If Not ForceGcAfterTask Then Return
            Try
                GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking:=True)
                GC.WaitForPendingFinalizers()
                GC.Collect(2, GCCollectionMode.Forced, blocking:=False)
            Catch
            End Try
        End Sub

        ''' <summary>Background full collection after any task. No-op when ForceGcAfterTask is off.</summary>
        Public Shared Sub CollectAfterTask()
            If Not ForceGcAfterTask Then Return
            Try
                GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking:=False)
            Catch
            End Try
        End Sub
    End Class

End Namespace
