Imports System.Data.SqlClient
Imports System.IO

Public Class Form7

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Button5.Enabled = True
            DateTimePicker1.Value.ToString("yyyy-MM-dd")
            DateTimePicker2.Value.ToString("yyyy-MM-dd")
            DBConn()
            SQLSTR = "SELECT * FROM Transactions WHERE materialname LIKE '%" & TextBox1.Text & "%' AND serialnumber LIKE '" & TextBox3.Text & "%' AND dateordered >= '" & DateTimePicker1.Value & "' AND datereceived <= '" & DateTimePicker2.Value & "' "
            readDB()
            ListView1.Clear()
            ListView1.GridLines = True
            ListView1.FullRowSelect = True
            ListView1.View = View.Details
            ListView1.Columns.Add("Transaction Number", 60)
            ListView1.Columns.Add("Material Name", 125)
            ListView1.Columns.Add("Amount", 60)
            ListView1.Columns.Add("Serial Number", 125)
            ListView1.Columns.Add("Specifications", 125)
            ListView1.Columns.Add("Date Ordered", 70)
            ListView1.Columns.Add("Date Received", 70)
            ListView1.Columns.Add("Price", 60)
            While (SQLDR.Read())
                With ListView1.Items.Add(SQLDR("transactionnumber"))
                    .subitems.add(SQLDR("materialname"))
                    .subitems.add(SQLDR("amount"))
                    .subitems.add(SQLDR("serialnumber"))
                    .subitems.add(SQLDR("specifications"))
                    .subitems.add(SQLDR("dateordered"))
                    .subitems.add(SQLDR("datereceived"))
                    .subitems.add(SQLDR("path"))
                    .subitems.add(SQLDR("price"))

                End With
            End While

            SQLDR.Close()
        Catch ex As SqlException
            MessageBox.Show(ex.Message, "MySQL Error: " & ex.Number, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, msgboxtitle)
        End Try

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Clear()
        TextBox3.Clear()
        DateTimePicker1.ResetText()
        DateTimePicker2.ResetText()
        ListView1.Clear()
        RadioButton2.Checked = True
        DateTimePicker1.Value = New Date(Now.Year, Now.Month, 1)

    End Sub

    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = New Date(Now.Year, Now.Month, 1)
        ListView1.Clear()
        Button5.Enabled = False


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim confirm As DialogResult = MsgBox("Are you sure you want to go back?", MsgBoxStyle.YesNo, msgboxtitle)
        If confirm = Windows.Forms.DialogResult.Yes Then
            Form2.Show()
            Me.Hide()
        End If

    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub

    Private Sub ListView1_MouseDown(ByVal sender As Object, _
        ByVal e As MouseEventArgs) Handles ListView1.MouseDown

        Dim selection As ListViewItem = ListView1.GetItemAt(e.X, e.Y)

        If Not (selection Is Nothing) Then
            PictureBox1.Image = System.Drawing.Image.FromFile _
                (selection.SubItems(7).Text)
        End If

        Button4.Enabled = True


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form8.Show()
        Me.Close()
        PictureBox1.Image = Nothing
        DateTimePicker1.Value.ToString("yyyy-MM-dd")
        DateTimePicker2.Value.ToString("yyyy-MM-dd")
        If Not ListView1.SelectedItems.Count = 0 Then
            With ListView1.SelectedItems.Item(0)
                Form8.materialNameTextBox.Text = .SubItems(1).Text
                Form8.amountTextBox.Text = .SubItems(2).Text
                Form8.serialNumTextBox.Text = .SubItems(3).Text
                Form8.specsTextBox.Text = .SubItems(4).Text
                Form8.DateTimePicker1.Value = .SubItems(5).Text
                Form8.DateTimePicker2.Value = .SubItems(6).Text
                Form8.TextBox1.Text = .SubItems(7).Text
                Form8.PictureBox1.Image = System.Drawing.Image.FromFile _
                    (.SubItems(7).Text)
                Form8.TextBox2.Text = .SubItems(0).Text
                Form8.TextBox3.Text = .SubItems(8).Text

            End With
        End If


    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Try
            Button5.Enabled = True
            DateTimePicker1.Value.ToString("yyyy-MM-dd")
            DateTimePicker2.Value.ToString("yyyy-MM-dd")
            DBConn()
            SQLSTR = "SELECT * FROM Transactions"
            readDB()
            ListView1.Clear()
            ListView1.GridLines = True
            ListView1.FullRowSelect = True
            ListView1.View = View.Details
            ListView1.Columns.Add("Transaction Number", 60)
            ListView1.Columns.Add("Material Name", 125)
            ListView1.Columns.Add("Amount", 60)
            ListView1.Columns.Add("Serial Number", 125)
            ListView1.Columns.Add("Specifications", 125)
            ListView1.Columns.Add("Date Ordered", 70)
            ListView1.Columns.Add("Date Received", 70)
            ListView1.Columns.Add("Price", 60)
            While (SQLDR.Read())
                With ListView1.Items.Add(SQLDR("transactionnumber"))
                    .subitems.add(SQLDR("materialname"))
                    .subitems.add(SQLDR("amount"))
                    .subitems.add(SQLDR("serialnumber"))
                    .subitems.add(SQLDR("specifications"))
                    .subitems.add(SQLDR("dateordered"))
                    .subitems.add(SQLDR("datereceived"))
                    .subitems.add(SQLDR("path"))
                    .subitems.add(SQLDR("price"))

                End With
            End While
            SQLDR.Close()
        Catch ex As SqlException
            MessageBox.Show(ex.Message, "MySQL Error: " & ex.Number, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, msgboxtitle)
        End Try
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        ListView1.Clear()
        Button5.Enabled = False
        PictureBox1.Image = Nothing
        Button4.Enabled = False

    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim saveFileDialog1 As New SaveFileDialog
        saveFileDialog1.Filter = "Excel File|*.xlsx"
        saveFileDialog1.Title = "Save an Excel File"
        saveFileDialog1.ShowDialog()
        If saveFileDialog1.FileName <> "" Then
            saveExcelFile(saveFileDialog1.FileName)
        End If
    End Sub

    Public Sub saveExcelFile(ByVal FileName As String)

        Dim xls As Object
        Dim sheet As Object

        xls = CreateObject("Excel.Application")
        xls.screenupdating = True
        xls.Visible = True


        Dim xlWorkSheet As Object = xls.workbooks.add
        sheet = xls.ActiveWorkbook.ActiveSheet
        Dim row As Integer = 2
        Dim col As Integer = 1
        sheet.Cells(1, "A").Value = "TransactionID"
        sheet.Cells(1, "B").Value = "Material Name"
        sheet.Cells(1, "C").Value = "Amount"
        sheet.Cells(1, "D").Value = "Serial Number"
        sheet.Cells(1, "E").Value = "Specifications"
        sheet.Cells(1, "F").Value = "Date Ordered"
        sheet.Cells(1, "G").Value = "Date Delivered"
        sheet.Cells(1, "I").Value = "Price"
        sheet.Cells(1, "J").Value = "Sum Total"

        For Each item As ListViewItem In ListView1.Items
            For i As Integer = 0 To item.SubItems.Count - 1
                sheet.Columns("A:Z").AutoFit()
                sheet.Cells(row, col).Value = item.SubItems(i).Text
                sheet.Columns("H").Value = ""

                col = col + 1
            Next
            row += 1
            col = 1
        Next


        sheet.Cells(row, "J").Value = "=SUM(I:I)"

        xlWorkSheet.SaveAs(FileName)
        xls = Nothing

    End Sub


End Class