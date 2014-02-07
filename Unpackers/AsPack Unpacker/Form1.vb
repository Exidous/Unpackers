Imports Un4seen.Bass

Public Class Form1
    Dim stream As Integer = 0
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With OpenFileDialog1
            .FileName = ""
            .Title = "Select program protected with AsPack"
            .ShowDialog()
        End With
        If OpenFileDialog1.FileName <> "" Then
            ClsUnpacker.UnpackAsPack(OpenFileDialog1.FileName)
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

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

        Un4seen.Bass.BassNet.Registration("Exidous2008@gmail.com", "2X2343021152222")

        If Bass.BASS_Init(-1, 44100, BASSInit.BASS_DEVICE_DEFAULT, IntPtr.Zero) Then
            'BassNet.register("test", "test")

            stream = Bass.BASS_MusicLoad("rept.it", 0, 0, BASSFlag.BASS_MUSIC_FT2MOD, 44100)
            If stream <> 0 Then
                Bass.BASS_ChannelPlay(stream, True)
            End If
        End If
    End Sub
End Class
