Public Class ClsUnpacker
    Public Shared Sub UnpackSecuPack(ByRef TheFile As String)
        Dim Debugger As New NonIntrusive.NIDebugger
        Dim StartOpts As New NonIntrusive.NIStartupOptions
        With StartOpts
            .commandLine = ""
            .executable = TheFile
            .resumeOnCreate = False
        End With
        Dim Unpacked As String = Strings.Left(TheFile, TheFile.LastIndexOf("\")) & "\Enter.exe"
        With Debugger
            .Execute(StartOpts)
TryAgain:
            .SetBreakpoint(.FindProcAddress("kernelbase.dll", "WriteFile"))
            .Continue()
            Dim f As Boolean = False
            Do Until f = True
                .StepIntoCalls = False
                .SingleStep()
                Dim opcd = .GetInstrOpcodes()
                If opcd.Length = 3 Then
                    If opcd(0) = &HC2 Then
                        If opcd(1) = &H14 Then
                            If opcd(2) = &H0 Then
                                f = True
                                .SingleStep()
                                .SingleStep()
                            End If
                        End If

                    End If

                End If
            Loop
            If FileIO.FileSystem.FileExists(Unpacked) Then
                .Detach.Terminate()
                FileIO.FileSystem.CopyFile(Unpacked, Strings.Left(TheFile, TheFile.Length - 4) & "_unpacked.exe")
                Process.Start("CMD.exe", "/c taskkill /im Enter.exe /f").WaitForExit()
                FileIO.FileSystem.DeleteFile(Unpacked)
                MsgBox("Successfully unpacked! Saved to:" & Strings.Left(TheFile, TheFile.Length - 4) & "_unpacked.exe")
            Else
                If FileIO.FileSystem.FileExists(Application.StartupPath & "\Enter.exe") Then
                    .Detach.Terminate()
                    Unpacked = Strings.Left(TheFile, TheFile.Length - 4) & "_unpacked.exe"
                    FileIO.FileSystem.CopyFile(Application.StartupPath & "\Enter.exe", Unpacked)
                    Process.Start("CMD.exe", "/c taskkill /im Enter.exe /f").WaitForExit()
                    FileIO.FileSystem.DeleteFile(Application.StartupPath & "\Enter.exe")
                    MsgBox("Successfully unpacked! Saved to:" & Environment.NewLine & Unpacked)
                Else
                    MsgBox("Unpacking Failed!?")
                    End If
                End If
        End With

    End Sub
End Class
