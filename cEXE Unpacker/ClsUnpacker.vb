Public Class ClsUnpacker
    Public Shared Sub UnpackcEXE(ByRef TheFile As String)
        Dim Debugger As New NonIntrusive.NIDebugger
        Dim StartOpts As New NonIntrusive.NIStartupOptions
        With StartOpts
            .commandLine = ""
            .executable = TheFile
            .resumeOnCreate = False
        End With

        With Debugger
            .Execute(StartOpts)
TryAgain:
            .SetBreakpoint(.FindProcAddress("kernel32.dll", "CreateProcessA"))
            .Continue()
            Dim z As UInteger = 0
            Dim fff As String = ""
            .ReadString(.Context.Eax, 1000, System.Text.Encoding.ASCII, fff)
            Dim SH() As String = Split(fff, """")
            Dim TmpFile As String = SH(1)
            TmpFile = Trim(TmpFile)
            Dim Fname As String = Strings.Right(TheFile, TheFile.Length - TheFile.LastIndexOf("\") - 1)
            Dim DelPath As String = Strings.Left(TheFile, TheFile.LastIndexOf("\")) & "\"
            Dim OutFile As String = Replace(TheFile, ".exe", "_unpacked.exe")
            Try
                FileIO.FileSystem.CopyFile(Trim(TmpFile), Trim(OutFile))
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            Dim Unpacked As String = OutFile
            If FileIO.FileSystem.FileExists(Unpacked) Then
                .Detach.Terminate()
                MsgBox("Successfully unpacked! Saved to:" & Unpacked & Environment.NewLine & Environment.NewLine & _
                       "Open Taskmanager and end all *.tmp processes, then delete the *.tmp files in the program directory!")
            Else
                .Detach.Terminate()
                MsgBox("Unpacking Failed!?")
            End If

        End With
    End Sub

End Class
