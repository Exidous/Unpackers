Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With OpenFileDialog1
            .Title = "Select Program Compressed With NeoLite"
            .FileName = ""
            .ShowDialog()
        End With

        If OpenFileDialog1.FileName <> "" Then
            ClsUnpacker.UnpackNeoLite(OpenFileDialog1.FileName)
        End If
    End Sub
End Class
