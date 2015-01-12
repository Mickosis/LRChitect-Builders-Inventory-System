Public Class Form1

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim connection As New SqlClient.SqlConnection
        Dim command As New SqlClient.SqlCommand
        Dim adaptor As New SqlClient.SqlDataAdapter
        Dim dataset As New DataSet

        connection.ConnectionString = ("Data Source=(LocalDB)\v11.0;AttachDbFilename=C:\Users\Miguel Rigunay\Desktop\LRC\LRChitect System\WindowsApplication2\LRChitect.mdf;Integrated Security=True")
        command.CommandText = " SELECT * FROM [programaccess] WHERE username = '" & TextBox1.Text & "'AND password = '" & TextBox2.Text & "';"
        connection.Open()

        command.Connection = connection
        adaptor.SelectCommand = command
        adaptor.Fill(dataset, "0")

        Dim count = dataset.Tables(0).Rows.Count

        If count > 0 Then
            Form2.Show()
            Me.Hide()
        Else
            MsgBox("Incorrect Login Credentials. Please try again.", MsgBoxStyle.OkOnly, msgboxtitle)
            TextBox1.Clear()
            TextBox2.Clear()
        End If
        connection.Close()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.CharacterCasing = CharacterCasing.Lower
        TextBox2.CharacterCasing = CharacterCasing.Lower
    End Sub

    Private Sub Button1_KeyDown(sender As Object, e As KeyEventArgs) Handles Button1.KeyDown

    End Sub


    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Form5.Show()
        Me.Hide()
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class
