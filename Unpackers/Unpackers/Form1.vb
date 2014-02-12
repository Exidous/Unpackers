Imports Un4seen.Bass

Public Class Form1
    Dim Debugger As New NonIntrusive.NIDebugger
    Dim stream As Integer = 0
    Private IsFormBeingDragged As Boolean = False
    Private MouseDownX As Integer
    Private MouseDownY As Integer
    Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MyBase.MouseDown

        If e.Button = Windows.Forms.MouseButtons.Left Then
            IsFormBeingDragged = True
            MouseDownX = e.X
            MouseDownY = e.Y
        End If
    End Sub

    Private Sub Form1_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MyBase.MouseUp

        If e.Button = MouseButtons.Left Then
            IsFormBeingDragged = False
        End If
    End Sub

    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MyBase.MouseMove

        If IsFormBeingDragged Then
            Dim temp As Point = New Point()

            temp.X = Me.Location.X + (e.X - MouseDownX)
            temp.Y = Me.Location.Y + (e.Y - MouseDownY)
            Me.Location = temp
            temp = Nothing
        End If
    End Sub

    Function CheckPacker(ByRef TheFile As String)
        Dim sigs As [String]() = New [String](9) {}
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
        'Ezip
        sigs(7) = "E9A6000000"
        'NeoLite
        sigs(8) = "558BEC83C4F0535657"
        'SecuPack
        sigs(9) = "558BEC81EC0C02????56BE0401????8D85F8FEFFFF56506A??FF15541040??8A8DF8FEFFFF33D284C98D85F8FEFFFF7416"
        'cEXE
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
        If sigMatch = 7 Then Return "NeoLite"
        If sigMatch = 8 Then Return "SecuPack"
        If sigMatch = 9 Then Return "cEXE"
        Return "NotFound"

    End Function

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Bass.BASS_StreamFree(stream)
        ' free BASS
        Bass.BASS_Free()


        If FileIO.FileSystem.FileExists("bass.dll") Then
            Try
                FileIO.FileSystem.DeleteFile("bass.dll")
            Catch ex As Exception

            End Try
        End If

        If FileIO.FileSystem.FileExists("Bass.Net.dll") Then
            Try
                FileIO.FileSystem.DeleteFile("Bass.Net.dll")
            Catch ex As Exception

            End Try
        End If
        If FileIO.FileSystem.FileExists("rept.it") Then
            Try
                FileIO.FileSystem.DeleteFile("rept.it")
            Catch ex As Exception

            End Try

        End If
    End Sub

    Private Sub Form1_DragDrop(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each path In files
            Unpak(path)
            Exit For
        Next
    End Sub

    Private Sub Form1_DragEnter(sender As System.Object, e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.AllowDrop = True
        Me.TopMost = True
        If FileIO.FileSystem.FileExists("bass.dll") Then
        Else
            FileIO.FileSystem.WriteAllBytes("bass.dll", My.Resources.bass, False)
        End If

        If FileIO.FileSystem.FileExists("Bass.Net.dll") Then
        Else
            FileIO.FileSystem.WriteAllBytes("Bass.Net.dll", My.Resources.Bass_Net, False)
        End If
        If FileIO.FileSystem.FileExists("rept.it") Then
        Else
            FileIO.FileSystem.WriteAllBytes("rept.it", My.Resources.reptinator, False)
        End If

        Un4seen.Bass.BassNet.Registration("Exidous@Team-REPT.com", "2X2342821182322")

        If Bass.BASS_Init(-1, 44100, BASSInit.BASS_DEVICE_DEFAULT, IntPtr.Zero) Then
            stream = Bass.BASS_MusicLoad("rept.it", 0, 0, BASSFlag.BASS_MUSIC_FT2MOD, 44100)
            If stream <> 0 Then
                Bass.BASS_ChannelPlay(stream, True)
            End If
        End If
    End Sub

    Sub unpak(Optional Path As String = "")
        Dim Sto As New NonIntrusive.NIStartupOptions
        If Path = "" Then
            With OpenFileDialog1
                .Title = "Select Packed .exe"
                .FileName = ""
                .ShowDialog()
            End With
        Else
            OpenFileDialog1.FileName = Path
        End If
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
                Me.TopMost = False
                MsgBox("Packer not detected! cannot continue :(")
            Else
                Me.TopMost = False
                If Paker = "Mpress" Then Mpress_Unpacker.Unpacker.UnpackMpress(OpenFileDialog1.FileName)
                If Paker = "PeCompact" Then PECompactUnpacker.Unpacker.UnpackePE(OpenFileDialog1.FileName)
                If Paker = "UPX" Then UPX_Unpacker.ClsUnpacker.UnpackUPX(OpenFileDialog1.FileName)
                If Paker = "FSG" Then FSG_Unpacker.ClsUnpacker.UnpackFSG(OpenFileDialog1.FileName)
                If Paker = "AsPack" Then AsPack_Unpacker.ClsUnpacker.UnpackAsPack(OpenFileDialog1.FileName)
                If Paker = "PeTite" Then Petite_Unpacker.Unpacker.UnpackePetite(OpenFileDialog1.FileName)
                If Paker = "EZIP" Then EZIP_Unpacker.ClsUnpacker.UnpackEZIP(OpenFileDialog1.FileName)
                If Paker = "NeoLite" Then NeoLite_Unpacker.ClsUnpacker.UnpackNeoLite(OpenFileDialog1.FileName)
                If Paker = "SecuPack" Then SecuPack.ClsUnpacker.UnpackSecuPack(OpenFileDialog1.FileName)
                If Paker = "cEXE" Then cEXE_Unpacker.ClsUnpacker.UnpackcEXE(OpenFileDialog1.FileName)
            End If
            MsgBox("Thanks bye!")
            End
        End If
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Unpak()
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        End
    End Sub

    Private Sub Label1_MouseDown(sender As Object, e As MouseEventArgs) Handles Label1.MouseDown
        Label1.Image = My.Resources.btn_big_unpk_3
    End Sub


    Private Sub Label1_MouseHover(sender As Object, e As EventArgs) Handles Label1.MouseHover
        Label1.Image = My.Resources.btn_big_unpk_2
    End Sub

    Private Sub Label1_MouseLeave(sender As Object, e As EventArgs) Handles Label1.MouseLeave
        Label1.Image = My.Resources.btn_big_unpk_1
    End Sub

    Private Sub Label1_MouseUp(sender As Object, e As MouseEventArgs) Handles Label1.MouseUp
        Label1.Image = My.Resources.btn_big_unpk_1
    End Sub

    Private Sub Label2_MouseDown(sender As Object, e As MouseEventArgs) Handles Label2.MouseDown
        Label2.Image = My.Resources.btn_exit_3
    End Sub


    Private Sub Label2_MouseHover(sender As Object, e As EventArgs) Handles Label2.MouseHover
        Label2.Image = My.Resources.btn_exit_2
    End Sub

    Private Sub Label2_MouseLeave(sender As Object, e As EventArgs) Handles Label2.MouseLeave
        Label2.Image = My.Resources.btn_exit_1
    End Sub

    Private Sub Label2_MouseUp(sender As Object, e As MouseEventArgs) Handles Label2.MouseUp
        Label2.Image = My.Resources.btn_big_unpk_1
    End Sub
End Class

