Public Class Form1

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
End Class
