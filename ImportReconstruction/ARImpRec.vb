Imports System.Runtime.InteropServices

Public Class ARImpRec
    <DllImport("ARImpRec.dll", CallingConvention:=CallingConvention.StdCall, EntryPoint:="SearchAndRebuildImports@28", CharSet:=CharSet.Ansi)> _
    Public Shared Function SearchAndRebuildImports(IRProcessId As UInteger, IRNameOfDumped As String, IROEP As UInt32, IRSaveOEPToFile As UInt32, ByRef IRIATRVA As UInt32, ByRef IRIATSize As UInt32, _
    IRWarning As IntPtr) As UInteger
    End Function

    <DllImport("ARImpRec.dll", CallingConvention:=CallingConvention.StdCall, EntryPoint:="SearchAndRebuildImportsIATOptimized@28", CharSet:=CharSet.Ansi)> _
    Public Shared Function SearchAndRebuildImportsIATOptimized(IRProcessId As UInteger, IRNameOfDumped As String, IROEP As UInt32, IRSaveOEPToFile As UInt32, ByRef IRIATRVA As UInt32, ByRef IRIATSize As UInt32, _
IRWarning As IntPtr) As UInteger
    End Function

    <DllImport("ARImpRec.dll", CallingConvention:=CallingConvention.StdCall, EntryPoint:="SearchAndRebuildImportsNoNewSection@28", CharSet:=CharSet.Ansi)> _
    Public Shared Function SearchAndRebuildImportsNoNewSection(IRProcessId As UInteger, IRNameOfDumped As String, IROEP As UInt32, IRSaveOEPToFile As UInt32, ByRef IRIATRVA As UInt32, ByRef IRIATSize As UInt32, _
IRWarning As IntPtr) As UInteger
    End Function

    Private Shared SavedTo As String
    Private Sub FixFileAllignment(ByRef TheFile As String)
        Dim FileArray() As Byte = FileIO.FileSystem.ReadAllBytes(TheFile)
        Dim AddrToEP As Long = BuildBytes(FileArray(60), FileArray(61), FileArray(62), FileArray(63))
        Dim AddrToAllignment As Long = AddrToEP + 59
        Dim Allignment As String = ""
        Dim AllignmentVal As Long
        Dim SectionAlignment As String = ""
        Dim SectionAlignmentVal As Long

        Dim i As Integer = 4
        Do Until i = 0
            Dim tmp As String = Hex(FileArray(AddrToAllignment + i))
            If tmp.Length = 1 Then tmp = "0" & tmp
            Allignment = Allignment & tmp
            i -= 1
        Loop

        i = 4
        Do Until i = 0
            Dim tmp As String = Hex(FileArray((AddrToAllignment - 4) + i))
            If tmp.Length = 1 Then tmp = "0" & tmp
            SectionAlignment = SectionAlignment & tmp
            i -= 1
        Loop
        SectionAlignmentVal = SectionAlignment
        AllignmentVal = Allignment
        If AllignmentVal < SectionAlignmentVal Then
            Dim ff() As Char = SectionAlignment
            i = 1
            Dim b As Integer = ff.Length - 2
            Do Until i = 5
                Dim TmpVar As Long = ff(b) & ff(b + 1)
                FileArray(AddrToAllignment + i) = "&h" & TmpVar
                i += 1
                b -= 2
            Loop
            FileIO.FileSystem.WriteAllBytes(TheFile, FileArray, False)
        End If
    End Sub
    Sub Initilize(ByVal MyPath As String)
        If FileIO.FileSystem.FileExists(MyPath & "ARImpRec.dll") Then
        Else
            FileIO.FileSystem.WriteAllBytes(MyPath & "ARImpRec.dll", My.Resources.ARImpRec, False)
        End If
    End Sub
    Function GetSavePath()
        Return SavedTo
    End Function
    Function FixImports(ByVal ProcID As UInteger, ByVal DumpPath As String, ByVal IROEP As UInteger, ByVal OrignalFile As String, Optional FixFileAlignment As Boolean = True)
        FixHeadderFlags(OrignalFile, DumpPath)
        If SearchAndRebuildImports(ProcID, DumpPath, IROEP) = False Then
            If SearchAndRebuildImportsNoNewSection(ProcID, DumpPath, IROEP) = False Then
                If SearchAndRebuildImportsIATOptimized(ProcID, DumpPath, IROEP) = False Then
                    Return False
                End If
            End If
        End If

        Dim Npath As String = Strings.Left(DumpPath, DumpPath.Length - 4) & "_.exe"
        Dim ReCheckCount As Integer = 0
ReCheck:

        If FileIO.FileSystem.FileExists(Npath) Then
            FileIO.FileSystem.DeleteFile(DumpPath)
            CleanupFiles(Npath, FixFileAlignment)
        Else

            If ReCheckCount >= 2 Then
                Return False
            End If

            If ReCheckCount = 0 Then
                If SearchAndRebuildImportsNoNewSection(ProcID, DumpPath, IROEP) = True Then
                    ReCheckCount += 1
                    GoTo ReCheck
                Else
                    ReCheckCount += 1
                    If SearchAndRebuildImportsIATOptimized(ProcID, DumpPath, IROEP) = True Then
                        ReCheckCount += 1
                        GoTo ReCheck
                    Else
                        Return False
                    End If
                End If
            Else
                If SearchAndRebuildImportsIATOptimized(ProcID, DumpPath, IROEP) = True Then
                    ReCheckCount += 1
                    GoTo ReCheck
                Else
                    Return False
                End If
            End If

            Return False
            'MsgBox("Auto import reconstruction failed!, Manually rebuilt now!")
        End If
        Return True
    End Function
    Private Function BuildBytes(ByRef Byt1 As Byte, byt2 As Byte, byt3 As Byte, byt4 As Byte)
        Dim Rebuilt As String = ""
        Dim Tmp As String = ""
        Tmp = Hex(byt4)
        If Tmp.Length = 1 Then Tmp = "0" & Tmp
        Rebuilt = Tmp
        Tmp = Hex(byt3)
        If Tmp.Length = 1 Then Tmp = "0" & Tmp
        Rebuilt = Rebuilt & Tmp
        Tmp = Hex(byt2)
        If Tmp.Length = 1 Then Tmp = "0" & Tmp
        Rebuilt = Rebuilt & Tmp
        Tmp = Hex(Byt1)
        If Tmp.Length = 1 Then Tmp = "0" & Tmp
        Rebuilt = Rebuilt & Tmp
        Dim val As Long = "&h" & Rebuilt
        Return val
    End Function
    Private Function SearchAndRebuildImports(ByRef ProcID, ByRef DumpPath, ByRef IROEP)
        Dim iatStart As UInt32 = 0
        Dim iatSize As UInt32 = 0
        Dim errorPtr As IntPtr = GetErrorPtr()
        Try
            Dim result As Integer = SearchAndRebuildImports(ProcID, DumpPath, IROEP, 1, iatStart, iatSize, errorPtr)
            Dim errorMessage As String = Marshal.PtrToStringAnsi(errorPtr)
            Marshal.FreeHGlobal(errorPtr)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Sub FixHeadderFlags(ByRef OrignalFile As String, ByRef DumpFile As String)
        Dim Orignal() As Byte = FileIO.FileSystem.ReadAllBytes(OrignalFile)
        Dim Dump() As Byte = FileIO.FileSystem.ReadAllBytes(DumpFile)
        Dim AddrToEPOG As Long = BuildBytes(Orignal(60), Orignal(61), Orignal(62), Orignal(63))
        Dim AddrToEPDU As Long = BuildBytes(Dump(60), Dump(61), Dump(62), Dump(63))
        Dim NumOfSections As Long = Orignal(AddrToEPOG + 7) & Orignal(AddrToEPOG + 6)
        'MsgBox(Hex(Orignal(AddrToEPOG + 7)))
        'MsgBox(Orignal(AddrToEPOG + 8))
        Dim i As Integer = 0
        Dim Pass As Integer = 1
        Do Until i = NumOfSections
            Dim SectionLength As Long = (&H27 * (Pass))
            If SectionLength > &H27 Then SectionLength = (&H28 * i) + &H27
            Dim Tmp As Long = AddrToEPOG + &HF8
            Dim Tmp2 As Long = AddrToEPDU + &HF8
            Dim SectFlag As Long = Orignal(Tmp + SectionLength)
            ' MsgBox(Hex(Dump(Tmp2 + SectionLength)))
            Dump(Tmp2 + SectionLength) = SectFlag
            Pass += 1
            i += 1
        Loop
        FileIO.FileSystem.WriteAllBytes(DumpFile, Dump, False)
    End Sub
    Private Function SearchAndRebuildImportsNoNewSection(ByRef ProcID, ByRef DumpPath, ByRef IROEP)
        Dim errorPtr As IntPtr
        Dim iatSize As UInteger
        Dim iatStart As UInteger
        Try
            iatStart = 0
            iatSize = 0
            errorPtr = GetErrorPtr()
            Dim result As Integer = SearchAndRebuildImportsNoNewSection(ProcID, DumpPath, IROEP, 0, iatStart, iatSize, errorPtr)
            Dim errorMessage As String = Marshal.PtrToStringAnsi(errorPtr)
            Marshal.FreeHGlobal(errorPtr)
            Return True
        Catch exx As Exception
            Return False
        End Try
    End Function
    Private Function GetErrorPtr()
        Return Marshal.AllocHGlobal(1000)
    End Function
    Private Function SearchAndRebuildImportsIATOptimized(ByRef ProcID, ByRef DumpPath, ByRef IROEP)
        Dim errorPtr As IntPtr
        Dim iatSize As UInteger
        Dim iatStart As UInteger
        Try
            iatSize = 0
            iatStart = 0
            errorPtr = GetErrorPtr()
            Dim result As Integer = SearchAndRebuildImportsIATOptimized(ProcID, DumpPath, IROEP, 0, iatStart, iatSize, errorPtr)
            Dim errorMessage As String = Marshal.PtrToStringAnsi(errorPtr)
            Marshal.FreeHGlobal(errorPtr)
            ' ReCheckCount += 1
            'GoTo ReCheck
        Catch exxx As Exception
            Return False
        End Try
        Return True
    End Function
    Private Sub CleanupFiles(ByVal NewPath As String, ByRef FixAlignment As Boolean)
        Dim Unpacked As String = Replace(NewPath, "dump_", "unpacked")
        Try
            FileIO.FileSystem.DeleteFile(Unpacked)
        Catch ex As Exception
        End Try

        FileIO.FileSystem.CopyFile(NewPath, Unpacked)
        FileIO.FileSystem.DeleteFile(NewPath)
        SavedTo = Unpacked
        If FixAlignment = True Then FixFileAllignment(SavedTo)
    End Sub
End Class
