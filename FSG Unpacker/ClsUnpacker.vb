Public Class ClsUnpacker

    Public Shared Sub UnpackFSG(ByRef TheProgram As String)
        Dim Debugger As New NonIntrusive.NIDebugger
        Dim StartOpts As New NonIntrusive.NIStartupOptions
        Dim DumpOpts As New NonIntrusive.NIDumpOptions
        Dim ImportRec As New ImportReconstruction.ARImpRec
        Dim SearchOpts As New NonIntrusive.NISearchOptions

        With StartOpts
            .commandLine = ""
            .executable = TheProgram
            .resumeOnCreate = False
        End With

        With Debugger
            .StepIntoCalls = False
            .Execute(StartOpts)
            .SetBreakpoint(.Context.Eip)
            .ClearBreakpoint(.Context.Eip)

            Dim Result() As UInteger = {}

                With SearchOpts
                    .SearchString = "78 F3 75 03 FF 63 0C"
                    .SearchImage = True
                    .MaxOccurs = 1
                End With

            .SearchMemory(SearchOpts, Result)

            If Result.Length > 0 Then
            Else
                MsgBox("Are you sure its protected with FSG?")
                End
            End If


            .SetBreakpoint((Result(0) + &H4))

            .Continue()

            .SingleStep()

            With DumpOpts
                .ChangeEP = True
                .EntryPoint = Debugger.Context.Eip - Debugger.Process.MainModule.BaseAddress
                .OutputPath = Strings.Left(TheProgram, TheProgram.Length - 4) & "_dump.exe"
                .PerformDumpFix = True
            End With

            .DumpProcess(DumpOpts)

            With ImportRec
                .Initilize(Application.StartupPath & "\")
                If .FixImports(Debugger.Process.Id, DumpOpts.OutputPath, DumpOpts.EntryPoint + Debugger.ProcessImageBase, TheProgram, True) = True Then
                    MsgBox("Successfully unpacked FSG! Saved To:" & Environment.NewLine & .GetSavePath)
                Else
                    MsgBox("Auto Rebuild Imports Failed!, Manually rebuild now!")
                End If

            End With
            .Detach.Terminate()
        End With

    End Sub
End Class
