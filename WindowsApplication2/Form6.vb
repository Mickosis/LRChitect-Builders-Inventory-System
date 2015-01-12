Public Class Form6

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim connection As New SqlClient.SqlConnection
        Dim command As New SqlClient.SqlCommand
        Dim adaptor As New SqlClient.SqlDataAdapter
        Dim dataset As New DataSet

        connection.ConnectionString = ("Data Source=(LocalDB)\v11.0;AttachDbFilename=C:\Users\Miguel Rigunay\Desktop\LRC\LRChitect System\WindowsApplication2\LRChitect.mdf;Integrated Security=True")
        command.CommandText = " SELECT * FROM [programaccess] WHERE username = 'root1' AND password = '" & TextBox2.Text & "';"
        connection.Open()

        command.Connection = connection
        adaptor.SelectCommand = command
        adaptor.Fill(dataset, "0")

        Dim count = dataset.Tables(0).Rows.Count
        connection.Close()

        If count > 0 Then
            DBConn()
            SQLSTR = "INSERT INTO programaccess (username, password) VALUES ('" & Form5.TextBox1.Text & "', '" & Form5.TextBox2.Text & "')"
            alterDB()
            MsgBox("Registration Successful. Please login now", , msgboxtitle)
            Form1.Show()
            Me.Hide()


        Else
            MsgBox("Incorrect Login Credentials. Please try again.", MsgBoxStyle.OkOnly, msgboxtitle)
            TextBox1.Clear()
            TextBox2.Clear()
        End If


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        Form5.Show()

    End Sub
End Class