Imports System.Runtime.InteropServices

Public Class Unpacker
    Dim debugger As New NonIntrusive.NIDebugger()

    <DllImport("ARImpRec.dll", CallingConvention:=CallingConvention.StdCall, EntryPoint:="SearchAndRebuildImports@28", CharSet:=CharSet.Ansi)> _
    Public Shared Function SearchAndRebuildImports(IRProcessId As UInteger, IRNameOfDumped As String, IROEP As UInt32, IRSaveOEPToFile As UInt32, ByRef IRIATRVA As UInt32, ByRef IRIATSize As UInt32, _
    IRWarning As IntPtr) As UInteger
    End Function

    Public Sub UnpackePetite(path As String)
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
            .Execute(SO)
            .InstallHardVEH()
            .SingleStep(5)
            Dim BPAddr As Long = .Context.Esp
            Dim myVal As UInteger = 0 '= 4

            Dim bytess() As Byte = {}
            .ReadData(BPAddr, 4, bytess)
            Array.Reverse(bytess)
            myVal = BitConverter.ToUInt32(bytess, 0)

            'MsgBox(Hex(myVal))
            'MsgBox(Hex(.Context.Eip))
            '.InstallHardVEH()
            .SetHardBreakPoint(.Context.Eip, NonIntrusive.HWBP_MODE.MODE_LOCAL, NonIntrusive.HWBP_TYPE.TYPE_READWRITE, NonIntrusive.HWBP_SIZE.SIZE_4)
            .Continue()
            '  MsgBox(.LastBreak.Context.Eip.ToString("X8"))
            'add the hardware bp == myVal
            'single step 2 times
            'we are at oep
            'dump process & rebuild Imports :D

            Dim NewEP As UInteger
            'NewEP = debugger.Context.Eip - debugger.Process.MainModule.BaseAddress
            NewEP = debugger.LastBreak.BreakAddress - debugger.Process.MainModule.BaseAddress
            .SetBreakpoint(debugger.LastBreak.BreakAddress + debugger.GetInstrLength())

            DumpOpts.EntryPoint = NewEP
            debugger.DumpProcess(DumpOpts)

            Dim iatStart As UInt32 = 0
            Dim iatSize As UInt32 = 0

            Dim errorPtr As IntPtr = Marshal.AllocHGlobal(1000)

            Try
                Dim result As Integer = SearchAndRebuildImports(debugger.Process.Id, DumpOpts.OutputPath, NewEP + debugger.ProcessImageBase, 0, iatStart, iatSize, errorPtr)
                Dim errorMessage As String = Marshal.PtrToStringAnsi(errorPtr)
                Marshal.FreeHGlobal(errorPtr)
            Catch ex As Exception

            End Try


            Dim Npath As String = Strings.Left(path, path.Length - 4) & "_.exe"

            If FileIO.FileSystem.FileExists(Npath) Then
                FileIO.FileSystem.DeleteFile(DumpOpts.OutputPath)
                FileIO.FileSystem.CopyFile(Npath, Strings.Left(Npath, Npath.LastIndexOf("\")) & "\Unpacked.exe")
                FileIO.FileSystem.DeleteFile(Npath)
                MsgBox("Unpacked!" & vbCrLf & "Saved to: " & Strings.Left(Npath, Npath.LastIndexOf("\")) & "\Unpacked.exe")
            Else
                MsgBox("Auto import reconstruction failed!, Manually rebuilt now!")
            End If

            .Detach()
            .Terminate()
        End With
    End Sub

   
End Class
