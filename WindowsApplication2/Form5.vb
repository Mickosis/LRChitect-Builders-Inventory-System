Public Class Form5

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Please input a username", MsgBoxStyle.OkOnly, msgboxtitle)
            TextBox1.Clear()
        ElseIf TextBox2.Text = "" Then
            MsgBox("Please indicate a password", MsgBoxStyle.OkOnly, msgboxtitle)
            TextBox2.Clear()
        Else
            Form6.Show()
            Me.Hide()
        End If


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim confirm As DialogResult = MsgBox("Are you sure you want to go back?", MsgBoxStyle.YesNo, msgboxtitle)
        If confirm = Windows.Forms.DialogResult.Yes Then
            Form1.Show()
            Me.Hide()
        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.CharacterCasing = CharacterCasing.Lower
        TextBox2.CharacterCasing = CharacterCasing.Lower
    End Sub
End Class