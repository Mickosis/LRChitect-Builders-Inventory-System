Public Class Form2


    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Form3.Show()
        Me.Hide()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form4.Show()
        Form4.materialNameTextBox.Clear()
        Form4.amountTextBox.Clear()
        Form4.serialNumTextBox.Clear()
        Form4.specsTextBox.Clear()
        Form4.DateTimePicker1.ResetText()
        Form4.DateTimePicker2.ResetText()
        Form4.TextBox1.Text = "C:\Users\Miguel Rigunay\Desktop\LRC\LRChitect System\WindowsApplication2\Resources\Default.jpg"
        Form4.PictureBox1.ImageLocation = "C:\Users\Miguel Rigunay\Desktop\LRC\LRChitect System\WindowsApplication2\Resources\Default.jpg"
        Me.Hide()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim result As DialogResult = MsgBox("Are you sure you want to log-out?", MsgBoxStyle.YesNo, msgboxtitle)
        If result = Windows.Forms.DialogResult.Yes Then
            Form1.Show()
            Form1.TextBox1.Clear()
            Form1.TextBox2.Clear()
            Me.Hide()
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Form3.Show()
        Me.Hide()
        Form3.TextBox1.Clear()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form7.TextBox1.Clear()
        Form7.TextBox3.Clear()
        Form7.DateTimePicker1.ResetText()
        Form7.DateTimePicker2.ResetText()
        Form7.ListView1.Clear()
        Form7.DateTimePicker1.Value = New Date(Now.Year, Now.Month, 1)
        Form7.Show()
        Me.Hide()
    End Sub
End Class