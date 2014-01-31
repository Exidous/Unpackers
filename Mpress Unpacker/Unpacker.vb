Public Class Unpacker
    Public Shared Sub UnpackMpress(ByRef TheFile As String)
        Dim debugger As New NonIntrusive.NIDebugger
        Dim DumpOpts As New NonIntrusive.NIDumpOptions
        Dim ImportRec As New ImportReconstruction.ARImpRec
        Dim StartOpts As New NonIntrusive.NIStartupOptions
        Dim SearchOpts As New NonIntrusive.NISearchOptions

        With StartOpts
            .commandLine = ""
            .executable = TheFile
            .resumeOnCreate = False
        End With

       
        With debugger
            .Execute(StartOpts)
            .SingleStep()
            With SearchOpts
                .SearchString = "E9 ?? ?? ?? ?? ?? ?? ?? ?? 00 00 00 00 ?? ?? 00 00"
                .SearchImage = True
                .MaxOccurs = 1


            End With
            'E9 ?? ?? ?? ?? ?? ?? ?? ?? 00 00 00 00 ?? ?? 00 00
            Dim list() As UInteger = {}
            .SearchMemory(SearchOpts, list)
            If list.Length > 0 Then
                .SetBreakpoint(list(0))
            Else
                MsgBox("Fail are you sure its protected with Mpress?!")
                Exit Sub
            End If
TryAgain:
            .Continue()
            With SearchOpts
                .SearchString = "E9 ?? ?? FF FF ?? ?? ?? ?? FF 25 ?? ?? ?? 00 FF 25 ?? ?? ?? 00"
                .SearchImage = True
                .MaxOccurs = 1
            End With
            list = {}
            .SearchMemory(SearchOpts, list)
            If list.Length > 0 Then
                .SetBreakpoint(list(0))
                .Continue()
            Else
                .SetBreakpoint(.Context.Eip)
                .SingleStep()
                GoTo TryAgain
            End If
            .SingleStep()

            With DumpOpts
                .ChangeEP = True
                .EntryPoint = debugger.Context.Eip
                .OutputPath = Strings.Left(TheFile, TheFile.Length - 4) & "_dump.exe"
            End With
            .DumpProcess(DumpOpts)
            With ImportRec
                .Initilize(Application.StartupPath & "\")
                If .FixImports(debugger.Process.Id, DumpOpts.OutputPath, DumpOpts.EntryPoint, TheFile, True) = True Then
                    MsgBox("Successfully unpacked! Saved to: " & Environment.NewLine & .GetSavePath)
                Else
                    MsgBox("Import Reconstruction failed, Manually rebuild now!")
                End If
            End With
        End With
    End Sub
End Class
