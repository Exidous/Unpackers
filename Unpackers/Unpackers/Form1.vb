Public Class Form1
    Dim Debugger As New NonIntrusive.NIDebugger

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Sto As New NonIntrusive.NIStartupOptions

        With OpenFileDialog1
            .Title = "Select Packed .exe"
            .FileName = ""
            .ShowDialog()
        End With
        If OpenFileDialog1.FileName <> "" Then
            With Sto
                .commandLine = ""
                .executable = OpenFileDialog1.FileName
                .resumeOnCreate = False
            End With
            With Debugger
                .Execute(Sto)
            End With
            Dim Paker As String = CheckPacker(OpenFileDialog1.FileName)
            Debugger.Detach.Terminate()
            If Paker = "NotFound" Then
                MsgBox("Packer not detected! cannot continue :(")
            Else
                If Paker = "Mpress" Then Mpress_Unpacker.Unpacker.UnpackMpress(OpenFileDialog1.FileName)
                If Paker = "PeCompact" Then PECompactUnpacker.Unpacker.UnpackePE(OpenFileDialog1.FileName)
                If Paker = "UPX" Then UPX_Unpacker.ClsUnpacker.UnpackUPX(OpenFileDialog1.FileName)
                If Paker = "FSG" Then FSG_Unpacker.ClsUnpacker.UnpackFSG(OpenFileDialog1.FileName)
                If Paker = "AsPack" Then AsPack_Unpacker.ClsUnpacker.UnpackAsPack(OpenFileDialog1.FileName)
                If Paker = "PeTite" Then Petite_Unpacker.Unpacker.UnpackePetite(OpenFileDialog1.FileName)
                If Paker = "EZIP" Then EZIP_Unpacker.ClsUnpacker.UnpackEZIP(OpenFileDialog1.FileName)
            End If
            MsgBox("Thanks bye!")
            End
        End If
    End Sub


    Function CheckPacker(ByRef TheFile As String)
        Dim sigs As [String]() = New [String](6) {}
        sigs(0) = "60BE??????008DBE??????FF"
        ' UPX
        sigs(1) = "B8????????5064FF35000000006489250000000033C08908504543"
        ' PECompact
        sigs(2) = "B8????????68????????64????????????64????????????669C6050"
        'PETite
        sigs(3) = "619455A4B680FF1373F933C9FF13731633C0FF13731FB68041B010FF1312C073FA75"
        'FSG 2.0
        sigs(4) = "60e8000000005805????00008b3003f02bc08bfe66adc1e00c8bc850ad2bc803f18bc85751498a44390688043175f6"
        'Mpress
        sigs(5) = "60E803000000E9EB045D4555C3E801"
        'AsPack
        sigs(6) = "E919320000E97C2A0000E919240000E9FF230000E91E2E0000E9882E0000E92C"
        Dim sigMatch As Integer = 0
        For x As Integer = 0 To sigs.Length - 1
            Dim opts As New NonIntrusive.NISearchOptions
            opts.SearchString = sigs(x)
            opts.MaxOccurs = 1
            opts.SearchImage = True
            Dim Arran() As UInteger = {}
            Debugger.SearchMemory(opts, Arran)
            If Arran.Length > 0 Then
                sigMatch = x
                Exit For
            Else
                sigMatch = 99
            End If
        Next
        
        If sigMatch = 4 Then Return "Mpress"
        If sigMatch = 1 Then Return "PeCompact"
        If sigMatch = 0 Then Return "UPX"
        If sigMatch = 3 Then Return "FSG"
        If sigMatch = 5 Then Return "AsPack"
        If sigMatch = 2 Then Return "PeTite"
        If sigMatch = 6 Then Return "EZIP"
        Return "NotFound"

    End Function
End Class
