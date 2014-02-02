Imports System.Runtime.InteropServices

Public Class Unpacker
    Private Shared debugger As New NonIntrusive.NIDebugger()

    Public Shared Sub UnpackePetite(path As String)
        Dim SO As New NonIntrusive.NIStartupOptions
        With SO
            .commandLine = ""
            .executable = path
            .resumeOnCreate = False
        End With

        Dim DumpOpts As New NonIntrusive.NIDumpOptions

        With DumpOpts
            .ChangeEP = True
            .PerformDumpFix = True
            .OutputPath = Strings.Left(path, path.Length - 4) & "_dump.exe"
        End With

        With debugger
            .AutoClearBP = True
            .Execute(SO)
            .InstallHardVEH()
            .SingleStep(5)
            .SetHardBreakPoint(.Context.Eip, NonIntrusive.HWBP_MODE.MODE_LOCAL, NonIntrusive.HWBP_TYPE.TYPE_READWRITE, NonIntrusive.HWBP_SIZE.SIZE_4)
            .Continue()
            Dim NewEP As UInteger
            NewEP = debugger.LastBreak.BreakAddress - debugger.Process.MainModule.BaseAddress
            .SetBreakpoint(debugger.LastBreak.BreakAddress + debugger.GetInstrLength())
            .Continue()
            .SetHardBreakPoint(.Context.Eip + .GetInstrLength, NonIntrusive.HWBP_MODE.MODE_LOCAL, NonIntrusive.HWBP_TYPE.TYPE_EXECUTE, NonIntrusive.HWBP_SIZE.SIZE_1)
Again:
            .Continue()
            Dim JmpAddr As UInteger = debugger.LastBreak.BreakAddress
            .SetBreakpoint(debugger.LastBreak.BreakAddress)
            .Continue()
            If .Context.Eax = &H0 Then
                .SingleStep()
                GoTo Again
            Else
                .SingleStep()
                Dim SEArchOpts As New NonIntrusive.NISearchOptions
                With SEArchOpts
                    .SearchString = "9D5FF3AA61669D83C408"
                    .SearchImage = True
                    .MaxOccurs = 1
                End With
                Dim Lst() As UInteger = {}
                .SearchMemory(SEArchOpts, Lst)
                If Lst.Length > 0 Then
                    .SetHardBreakPoint(Lst(0) + &HA, NonIntrusive.HWBP_MODE.MODE_LOCAL, NonIntrusive.HWBP_TYPE.TYPE_EXECUTE, NonIntrusive.HWBP_SIZE.SIZE_1)
                    .Continue()
                    .SetBreakpoint(.LastBreak.BreakAddress)
                    .Continue()
                    .SingleStep()
                Else
                    Exit Sub
                End If
            End If
            NewEP = .Context.Eip - .Process.MainModule.BaseAddress
            Clipboard.Clear()
            Clipboard.SetText(Hex(NewEP))
            DumpOpts.EntryPoint = NewEP
            debugger.DumpProcess(DumpOpts)

            Dim ImportRec As New ImportReconstruction.ARImpRec
            With ImportRec
                .Initilize(Application.StartupPath & "\")
                If .FixImports(debugger.Process.Id, DumpOpts.OutputPath, NewEP + debugger.ProcessImageBase, path, True) = True Then
                    MsgBox("Unpacked!" & vbCrLf & "Saved to: " & .GetSavePath)
                Else
                    MsgBox("Auto import reconstruction failed!, Manually rebuilt now!")
                End If
            End With

            .Detach()
            .Terminate()
        End With
    End Sub

   
End Class
