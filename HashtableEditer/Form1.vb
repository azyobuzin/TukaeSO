Public Class Form1

    Private Sub OpenButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenButton.Click
        If OpenFileDialog1.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            Using fs As New IO.FileStream(OpenFileDialog1.FileName, IO.FileMode.Open)
                Dim bf As New Runtime.Serialization.Formatters.Binary.BinaryFormatter
                HashtableEditGrid1.Hashtable = bf.Deserialize(fs)
            End Using
        End If
    End Sub

    Private Sub SaveButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveButton.Click
        If SaveFileDialog1.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
            Using fs As New IO.FileStream(SaveFileDialog1.FileName, IO.FileMode.Create)
                Dim bf As New Runtime.Serialization.Formatters.Binary.BinaryFormatter
                bf.Serialize(fs, HashtableEditGrid1.Hashtable)
            End Using
        End If
    End Sub
End Class
