Imports System.Data.SqlClient

Public Class Form3



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Old password field is empty", MsgBoxStyle.OkOnly, msgboxtitle)
            TextBox1.Clear()
        ElseIf TextBox2.Text = "" Then
            MsgBox("New password field is empty", MsgBoxStyle.OkOnly, msgboxtitle)
            TextBox2.Clear()
        ElseIf TextBox3.Text = "" Then
            MsgBox("Repeat password field is empty", MsgBoxStyle.OkOnly, msgboxtitle)
            TextBox3.Clear()
        Else
        If TextBox2.Text <> TextBox3.Text Then
            MsgBox("New passwords must match", MsgBoxStyle.OkOnly, msgboxtitle)
            TextBox2.Clear()
            TextBox3.Clear()
        ElseIf TextBox3.Text = TextBox1.Text Then
            MsgBox("New password matched old password", MsgBoxStyle.OkOnly, msgboxtitle)
            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()

        Else
            Dim connection As New SqlClient.SqlConnection
            Dim command As New SqlClient.SqlCommand
            Dim adaptor As New SqlClient.SqlDataAdapter
            Dim dataset As New DataSet

                connection.ConnectionString = ("Data Source=(LocalDB)\v11.0;AttachDbFilename=C:\Users\Miguel Rigunay\Desktop\LRC\LRChitect System\WindowsApplication2\LRChitect.mdf;Integrated Security=True")
            command.CommandText = " SELECT * FROM [programaccess] WHERE username = '" & Form1.TextBox1.Text & "'AND password = '" & TextBox1.Text & "';"
            connection.Open()

            command.Connection = connection
            adaptor.SelectCommand = command
            adaptor.Fill(dataset, "0")

            Dim count = dataset.Tables(0).Rows.Count

            If count > 0 Then
                DBConn()
                SQLSTR = " UPDATE [programaccess] SET password = '" & TextBox3.Text & "' WHERE username = '" & Form1.TextBox1.Text & "'"
                alterDB()
                    MsgBox("Password changed successfully. Please login again.", MsgBoxStyle.OkOnly, msgboxtitle)
                    Form1.TextBox1.Clear()
                    Form1.TextBox2.Clear()
                Form1.Show()
                Me.Close()

            Else
                MsgBox("Password does not match with original password.", MsgBoxStyle.OkOnly, msgboxtitle)
                TextBox1.Clear()
            End If

            connection.Close()
        End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.CharacterCasing = CharacterCasing.Lower
        TextBox2.CharacterCasing = CharacterCasing.Lower
        TextBox3.CharacterCasing = CharacterCasing.Lower
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim confirm As DialogResult = MsgBox("Are you sure you want to go back?", MsgBoxStyle.YesNo, msgboxtitle)
        If confirm = Windows.Forms.DialogResult.Yes Then
            Form2.Show()
            Me.Hide()
        End If

    End Sub
End Class