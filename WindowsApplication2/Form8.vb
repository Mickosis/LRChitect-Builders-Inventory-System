Imports System.IO
Imports System.Data.SqlClient

Public Class Form8

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim confirm As DialogResult = MsgBox("Are you sure you want to go back?", MsgBoxStyle.YesNo, msgboxtitle)
        If confirm = Windows.Forms.DialogResult.Yes Then
            Form2.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox1.Clear()
        amountTextBox.Clear()
        materialNameTextBox.Clear()
        serialNumTextBox.Clear()
        specsTextBox.Clear()
        DateTimePicker1.ResetText()
        DateTimePicker2.ResetText()
        PictureBox1.Image = Nothing
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim confirm As DialogResult = MsgBox("Are all information correct?", MsgBoxStyle.YesNo, msgboxtitle)
        If confirm = Windows.Forms.DialogResult.Yes Then

            If materialNameTextBox.Text = ("") Then
                MsgBox("Please input a material name", MsgBoxStyle.OkOnly, msgboxtitle)
                materialNameTextBox.Clear()
            ElseIf amountTextBox.Text = ("") Then
                MsgBox("Please input an amount", MsgBoxStyle.OkOnly, msgboxtitle)
                amountTextBox.Clear()
            ElseIf Not ValidateNumeric(amountTextBox.Text) Then
                MsgBox("Please input a proper value for amount", MsgBoxStyle.OkOnly, msgboxtitle)
                amountTextBox.Clear()
            ElseIf Not ValidateNumeric(serialNumTextBox.Text) Then
                MsgBox("Please input a proper serial number", MsgBoxStyle.OkOnly, msgboxtitle)
                serialNumTextBox.Clear()
            ElseIf specsTextBox.Text = ("") Then
                MsgBox("Please a specification or description", MsgBoxStyle.OkOnly, msgboxtitle)
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
                If TextBox1.Text.Contains(TextBox1.Text) = True Then
                    DateTimePicker1.Value.ToString("yyyy-MM-dd")
                    DateTimePicker2.Value.ToString("yyyy-MM-dd")
                    SQLSTR = "UPDATE Transactions SET materialname= '" & materialNameTextBox.Text & "', amount = '" & amountTextBox.Text & "', serialnumber = '" & serialNumTextBox.Text & "', specifications = '" & specsTextBox.Text & "', dateordered = '" & DateTimePicker1.Value & "', datereceived = '" & DateTimePicker2.Value & "' WHERE transactionnumber = '" & TextBox2.Text & "'"
                    alterDB()
                    MsgBox("Update successful!", , msgboxtitle)
                    SQLCONN.Close()
                Else
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
                    SQLSTR = "UPDATE Transactions SET materialname= '" & materialNameTextBox.Text & "', amount = '" & amountTextBox.Text & "', serialnumber = '" & serialNumTextBox.Text & "', specifications = '" & specsTextBox.Text & "', dateordered = '" & DateTimePicker1.Value & "', datereceived = '" & DateTimePicker2.Value & "', receipt = @photo, path = @name  WHERE transactionnumber = '" & TextBox2.Text & "'"
                    alterDB()
                    MsgBox("Update successful!", , msgboxtitle)
                    SQLCONN.Close()
                    PictureBox1.ImageLocation = "C:\Users\Miguel Rigunay\Desktop\LRC\LRChitect System\WindowsApplication2\Resources\Default.jpg"
                    TextBox1.Text = "C:\Users\Miguel Rigunay\Desktop\LRC\LRChitect System\WindowsApplication2\Resources\Default.jpg"
                    ms.Dispose()
                End If


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

    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class