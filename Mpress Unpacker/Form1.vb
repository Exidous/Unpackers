Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        With OpenFileDialog1
            .Title = "Select program protected with Mpress"
            .FileName = ""
            .ShowDialog()
        End With

        If OpenFileDialog1.FileName <> "" Then Call Unpacker.UnpackMpress(OpenFileDialog1.FileName)


    End Sub
End Class
