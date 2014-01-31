
Public Class Unpacker

    Dim debugger As New NonIntrusive.NIDebugger()
    Public Sub UnpackePE(path As String)
        Dim opts As New NonIntrusive.NIStartupOptions()
        opts.executable = path
        opts.resumeOnCreate = False

        Dim y As UInteger
        Dim opCode(2) As Byte
        opCode(1) = &HFF
        opCode(2) = &HE0

        debugger.StepIntoCalls = False

        Dim dumpOpts As New NonIntrusive.NIDumpOptions()
        dumpOpts.ChangeEP = True
        dumpOpts.OutputPath = Strings.Left(path, path.Length - 4) & "_dump.exe"
        dumpOpts.PerformDumpFix = True

        Dim newEP As UInteger

        debugger.Execute(opts) _
            .ReadDWORD(debugger.Context.Eip + 1, y) _
            .SetBreakpoint(y) _
            .Continue() _
            .SingleStep(3) _
            .SetBreakpoint(debugger.Context.Ecx) _
            .Continue() _
            .Until(AddressOf FoundJMP, AddressOf debugger.SingleStep) _
            .SingleStep()

        newEP = debugger.Context.Eip - debugger.Process.MainModule.BaseAddress

        dumpOpts.EntryPoint = newEP

        debugger.DumpProcess(dumpOpts)

        Clipboard.Clear()
        'set clipboard OEP/RVA
        Clipboard.SetText(Hex(newEP))

        Dim ImportFixer As New ImportReconstruction.ARImpRec
        ImportFixer.Initilize(Application.StartupPath & "\")

        If ImportFixer.FixImports(debugger.Process.Id, dumpOpts.OutputPath, newEP + debugger.ProcessImageBase, path, True) = True Then
            MsgBox("Successfully unpacked! Saved to: " & Environment.NewLine & ImportFixer.GetSavePath)
        Else
            MsgBox("Import Reconstruction failed, Manually rebuild now!")
        End If

        debugger.Detach().Terminate()

    End Sub

    Public Function FoundJMP()
        If debugger.Context.Eip > &H405240 Then
            Dim i As Integer = 0
        End If
        Dim data() As Byte = debugger.GetInstrOpcodes()
        If (data.Length < 2) Then
            Return False
        End If
        If (data(0) = &HFF And (data(1) = &HE0)) Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
