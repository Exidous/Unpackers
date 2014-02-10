Public Class ClsUnpacker
    Public Shared Sub UnpackAsPack(ByRef TheFile As String)
        Dim StartOpts As New NonIntrusive.NIStartupOptions
        Dim DumpOpts As New NonIntrusive.NIDumpOptions
        Dim SearchOpts As New NonIntrusive.NISearchOptions
        Dim ImportRec As New ImportReconstruction.ARImpRec
        Dim Debugger As New NonIntrusive.NIDebugger

        With StartOpts
            .commandLine = ""
            .executable = TheFile
            .resumeOnCreate = False
        End With

        With SearchOpts
            .SearchString = "7508B801000000"
            .SearchImage = True
            .MaxOccurs = 1
        End With

        With Debugger
            .Execute(StartOpts)
            Dim lst() As UInteger = {}
            .SearchMemory(SearchOpts, lst)
            If lst.Length > 0 Then
                .SetBreakpoint(lst(0))
                .Continue()
                .SingleStep(3)
                Dim NEWEP As UInteger = Debugger.Context.Eip - Debugger.Process.MainModule.BaseAddress
                With DumpOpts
                    .EntryPoint = NEWEP
                    .OutputPath = Strings.Left(TheFile, TheFile.Length - 4) & "_dump.exe"
                    .PerformDumpFix = True
                End With
                .DumpProcess(DumpOpts)
                With ImportRec
                    .Initilize(Application.StartupPath & "\")
                    If .FixImports(Debugger.Process.Id, DumpOpts.OutputPath, NEWEP + Debugger.ProcessImageBase, TheFile, True) Then
                        MsgBox("Successfully unpacked! Saved to:" & Environment.NewLine & .GetSavePath)
                    Else
                        MsgBox("Import rec failed, manually rebuild imports!")
                    End If
                End With
            End If
            .Detach.Terminate()
        End With
    End Sub
End Class
