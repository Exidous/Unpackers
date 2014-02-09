Imports Un4seen.Bass

Public Class Form1
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
            unpak(path)
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
            'BassNet.register("test", "test")

            stream = Bass.BASS_MusicLoad("rept.it", 0, 0, BASSFlag.BASS_MUSIC_FT2MOD, 44100)
            If stream <> 0 Then
                Bass.BASS_ChannelPlay(stream, True)
            End If
        End If
    End Sub

    Sub unpak(Optional FileName As String = "")
        If FileName = "" Then
            With OpenFileDialog1
                .Title = "Select program packed w/ Petite"
                .FileName = ""
                .ShowDialog()
            End With
        Else
            OpenFileDialog1.FileName = FileName
        End If
        If OpenFileDialog1.FileName <> "" Then
            Dim PE As New Unpacker
            Unpacker.UnpackePetite(OpenFileDialog1.FileName)
        End If
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Call unpak()
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
