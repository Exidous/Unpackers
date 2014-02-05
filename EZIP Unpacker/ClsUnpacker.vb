Public Class ClsUnpacker
    Public Shared Sub UnpackEZIP(ByRef TheFile As String)
        Dim ImportRec As New ImportReconstruction.ARImpRec
        Dim StartOptions As New NonIntrusive.NIStartupOptions
        Dim DumpOpts As New NonIntrusive.NIDumpOptions
        Dim Debugger As New NonIntrusive.NIDebugger
        Dim SearchOpts As New NonIntrusive.NISearchOptions

        With StartOptions
            .commandLine = ""
            .executable = TheFile
            .resumeOnCreate = False
        End With

        With SearchOpts
            .SearchString = "FFE0????"
            .SearchImage = True
        End With

        Dim lst() As UInteger = {}
        With Debugger
            .Execute(StartOptions)
            .SearchMemory(SearchOpts, lst)
            If lst.Length > 0 Then
                .SetBreakpoint(lst(lst.Length - 1))
                .Continue()
                .StepIntoCalls = False
                .SingleStep()
                With DumpOpts
                    .ChangeEP = True
                    .EntryPoint = Debugger.Context.Eip - Debugger.Process.MainModule.BaseAddress
                    .OutputPath = Strings.Left(TheFile, TheFile.Length - 4) & "_dump.exe"
                    .PerformDumpFix = True
                End With
                .DumpProcess(DumpOpts)
                With ImportRec
                    .Initilize(Application.StartupPath & "\")
                    If .FixImports(Debugger.Process.Id, DumpOpts.OutputPath, DumpOpts.EntryPoint + Debugger.ProcessImageBase, TheFile, True) Then
                        MsgBox("Successfully unpacked! Saved to:" & Environment.NewLine & .GetSavePath)
                    Else
                        MsgBox("Manually rebuild imports!")
                    End If
                End With
            End If
        End With
    End Sub
End Class
