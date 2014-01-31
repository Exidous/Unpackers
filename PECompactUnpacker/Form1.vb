Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()

        Dim unp As New Unpacker()

        unp.UnpackePE(OpenFileDialog1.FileName)


    End Sub
End Class
