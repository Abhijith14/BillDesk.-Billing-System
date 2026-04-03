Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form2
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Cursor = Cursors.WaitCursor
        Dim data_obj As Object = load_data()
        Me.Cursor = Cursors.Default
        MsgBox("Data Loaded Successfully!", MsgBoxStyle.Information, "BillDesk")
    End Sub

    Function load_data() As Object
        Dim datagrid_count As Integer = DataGridView1.Rows.Count
        Dim rowcount As Integer
        Try
            Dim path As String = My.Settings("ExcelPath")
            Using MyConnection As New System.Data.OleDb.OleDbConnection(DbHelper.GetConnectionString(path))
                Dim MyCommand As New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
                MyCommand.TableMappings.Add("Table", "Net-informations.com")
                Dim DtSet As New System.Data.DataSet
                Dim Dttbl As New DataTable()
                MyCommand.Fill(Dttbl)
                rowcount = Dttbl.Rows.Count()
                DtSet.Tables.Add(Dttbl)
                DataGridView1.DataSource = DtSet.Tables(0)

                For i As Integer = 0 To 5
                    Me.DataGridView1.Columns(i).ReadOnly = True
                Next
            End Using
        Catch ex As System.Data.OleDb.OleDbException
            My.Settings("ExcelPath") = ""
            My.Settings.Save()
            Form1.Button1.Enabled = False
            Form1.Button2.Enabled = False
            Form1.Button4.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
        Return {datagrid_count, rowcount}
    End Function

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.DoubleBuffered = True
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Try
            Dim path As String = My.Settings("ExcelPath")
            Using MyConnection As New System.Data.OleDb.OleDbConnection(DbHelper.GetConnectionString(path))
                Using cmd As New System.Data.OleDb.OleDbCommand("select * from [Sheet1$] where [Enrollment No#] LIKE ?", MyConnection)
                    cmd.Parameters.AddWithValue("?", TextBox2.Text & "%")
                    Dim MyCommand As New System.Data.OleDb.OleDbDataAdapter(cmd)
                    MyCommand.TableMappings.Add("Table", "Net-informations.com")
                    Dim DtSet As New System.Data.DataSet
                    Dim Dttbl As New DataTable()
                    MyCommand.Fill(Dttbl)
                    DtSet.Tables.Add(Dttbl)
                    DataGridView1.DataSource = DtSet.Tables(0)
                End Using
            End Using
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Try
            Dim path As String = My.Settings("ExcelPath")
            Using MyConnection As New System.Data.OleDb.OleDbConnection(DbHelper.GetConnectionString(path))
                Using cmd As New System.Data.OleDb.OleDbCommand("select * from [Sheet1$] where [Student Name] LIKE ?", MyConnection)
                    cmd.Parameters.AddWithValue("?", "%" & TextBox1.Text & "%")
                    Dim MyCommand As New System.Data.OleDb.OleDbDataAdapter(cmd)
                    MyCommand.TableMappings.Add("Table", "Net-informations.com")
                    Dim DtSet As New System.Data.DataSet
                    Dim Dttbl As New DataTable()
                    MyCommand.Fill(Dttbl)
                    DtSet.Tables.Add(Dttbl)
                    DataGridView1.DataSource = DtSet.Tables(0)
                End Using
            End Using
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub

    Function SaveData(r As Integer, c As Integer)
        Dim value As Integer = DataGridView1.Rows(r).Cells(c).Value
        Dim xlApp As Excel.Application = Nothing
        Dim xlWorkBook As Excel.Workbook = Nothing
        Dim xlWorkSheet As Excel.Worksheet = Nothing
        Try
            xlApp = New Excel.Application
            Dim path As String = My.Settings("ExcelPath")
            xlWorkBook = xlApp.Workbooks.Open(path)
            xlWorkSheet = xlWorkBook.Worksheets("Sheet1")
            xlWorkSheet.Cells(r + 2, c + 1) = value
            xlWorkBook.Close()
            xlApp.Quit()
        Catch ex As Exception
            MsgBox("Error saving data: " & ex.Message)
        Finally
            If xlWorkSheet IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet)
            If xlWorkBook IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook)
            If xlApp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
            xlWorkSheet = Nothing
            xlWorkBook = Nothing
            xlApp = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If e.RowIndex >= 0 Then
            SaveData(e.RowIndex, e.ColumnIndex)
            MsgBox("Updated.")
        End If
    End Sub
End Class
