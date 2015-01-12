Imports System.IO
Imports System.Data.SqlClient

Public Class Form4



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim confirm As DialogResult = MsgBox("Are you sure you want to go back?", MsgBoxStyle.YesNo, msgboxtitle)
        If confirm = Windows.Forms.DialogResult.Yes Then
        Form2.Show()
            Me.Hide()
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim confirm As DialogResult = MsgBox("Are all information correct?", MsgBoxStyle.YesNo, msgboxtitle)
        If confirm = Windows.Forms.DialogResult.Yes Then
                SQLCONN.Close()

                If materialNameTextBox.Text = ("") Then
                    MsgBox("Please input a material name", MsgBoxStyle.OkOnly, msgboxtitle)
                    materialNameTextBox.Clear()
                ElseIf amountTextBox.Text = ("") Then
                    MsgBox("Please input an amount", MsgBoxStyle.OkOnly, msgboxtitle)
                    amountTextBox.Clear()
                ElseIf TextBox2.Text = ("") Then
                    MsgBox("Please input a price", MsgBoxStyle.OkOnly, msgboxtitle)
                ElseIf Not ValidateNumeric(TextBox2.Text) Then
                    MsgBox("Please input a proper value for price", MsgBoxStyle.OkOnly, msgboxtitle)
                    amountTextBox.Clear()
                ElseIf Not ValidateNumeric(amountTextBox.Text) Then
                    MsgBox("Please input a proper value for amount", MsgBoxStyle.OkOnly, msgboxtitle)
                    amountTextBox.Clear()
                ElseIf specsTextBox.Text = ("") Then
                    MsgBox("Please add a specification, or description", MsgBoxStyle.OkOnly, msgboxtitle)
                    specsTextBox.Clear()
                ElseIf DateTimePicker2.Value.Ticks <= DateTimePicker1.Value.Ticks Then
                    MsgBox("Date received is earlier than date ordered", MsgBoxStyle.OkOnly, msgboxtitle)
                Else

                    Dim sourcepath As String = TextBox1.Text
                    Dim DestPath As String = "C:\LRChitect\"
                    Dim folderpath As String = specsTextBox.Text
                    If Not Directory.Exists(DestPath) Then
                        Directory.CreateDirectory(DestPath)
                    End If
                    Directory.CreateDirectory("C:\LRChitect\" + folderpath)
                    Dim file = New FileInfo(TextBox1.Text)
                    file.CopyTo(Path.Combine(DestPath, specsTextBox.Text, file.Name), True)
                    DBConn()
                    SQLCMD.Parameters.AddWithValue("@name", Path.Combine(DestPath, specsTextBox.Text, file.Name))
                    Dim ms As New MemoryStream()
                    PictureBox1.Image.Save(ms, PictureBox1.Image.RawFormat)
                    Dim data As Byte() = ms.GetBuffer()
                    Dim p As New SqlParameter("@photo", SqlDbType.Image)
                    p.Value = data
                    SQLCMD.Parameters.Add(p)
                    DateTimePicker1.Value.ToString("yyyy-MM-dd")
                    DateTimePicker2.Value.ToString("yyyy-MM-dd")
                    SQLSTR = "INSERT INTO Transactions (materialname, amount, price, serialnumber, specifications, dateordered, datereceived, receipt, path) VALUES ('" & materialNameTextBox.Text & "', '" & amountTextBox.Text & "', '" & TextBox2.Text & "', '" & serialNumTextBox.Text & "', '" & specsTextBox.Text & "', '" & DateTimePicker1.Value & "', '" & DateTimePicker2.Value & "', @photo, @name)"
                    alterDB()
                    MsgBox("Input successful!", , msgboxtitle)
                    SQLCONN.Close()
                    SQLCMD.Parameters.Clear()

                    amountTextBox.Clear()
                    materialNameTextBox.Clear()
                    serialNumTextBox.Clear()
                    specsTextBox.Clear()
                    TextBox2.Clear()
                    DateTimePicker1.ResetText()
                    DateTimePicker2.ResetText()
                PictureBox1.ImageLocation = "C:\Users\Miguel Rigunay\Desktop\LRC\LRChitect System\WindowsApplication2\Resources\Default.jpg"
                TextBox1.Text = "C:\Users\Miguel Rigunay\Desktop\LRC\LRChitect System\WindowsApplication2\Resources\Default.jpg"
                End If


        End If

    End Sub

    Private Function ValidateNumeric(strText As String) _
As Boolean
        ValidateNumeric = CBool(strText = "" _
            Or strText = "-" _
            Or strText = "-." _
            Or strText = "." _
            Or IsNumeric(strText))
    End Function


    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged

    End Sub

    Private Sub materialNameTextBox_TextChanged(sender As Object, e As EventArgs) Handles materialNameTextBox.TextChanged

    End Sub

    Private Sub amountTextBox_TextChanged(sender As Object, e As EventArgs) Handles amountTextBox.TextChanged

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        amountTextBox.Clear()
        materialNameTextBox.Clear()
        serialNumTextBox.Clear()
        specsTextBox.Clear()
        DateTimePicker1.ResetText()
        DateTimePicker2.ResetText()
        PictureBox1.ImageLocation = "C:\Users\Miguel Rigunay\Desktop\WindowsApplication2\WindowsApplication2\Resources\Default.jpg"

    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim fdlg As OpenFileDialog = New OpenFileDialog()
        fdlg.Title = "Choose a Profile Photo"
        fdlg.Filter = "Picture Files(*.jpg;*.jpeg;*.png;*.bmp;*.gif)|*.jpg;*.jpeg;*.png;*.bmp;*.gif"
        fdlg.FilterIndex = 2
        fdlg.RestoreDirectory = True
        If fdlg.ShowDialog() = DialogResult.OK Then
            If File.Exists(fdlg.FileName) = False Then
                MessageBox.Show("Sorry, The File You Specified Does Not Exist.", msgboxtitle)
            Else
                PictureBox1.ImageLocation = fdlg.FileName
            End If

        End If

        TextBox1.Text = fdlg.FileName.ToString
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        amountTextBox.Clear()
        materialNameTextBox.Clear()
        serialNumTextBox.Clear()
        specsTextBox.Clear()
        DateTimePicker1.ResetText()
        TextBox2.Clear()
        DateTimePicker2.ResetText()
        PictureBox1.ImageLocation = "C:\Users\Miguel Rigunay\Desktop\LRChitect System\WindowsApplication2\Resources\Default.jpg"
        TextBox1.Text = "C:\Users\Miguel Rigunay\Desktop\LRChitect System\WindowsApplication2\Resources\Default.jpg"
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class