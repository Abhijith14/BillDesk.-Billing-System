Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form2
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim data_obj As Object = load_data()

        MsgBox("Data Loaded Successfully !")
    End Sub

    Function load_data() As Object

        Dim datagrid_count As Integer = DataGridView1.Rows.Count
        Dim rowcount As Integer
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim DtSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = My.Settings("ExcelPath") '"C:\Users\ABHIJITH UDAYAKUMAR\OneDrive\Documents\Billing System User Files\data.xlsx"
            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.16.0;Data Source='" + path + "';Extended Properties=Excel 8.0;")
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
            MyCommand.TableMappings.Add("Table", "Net-informations.com")
            DtSet = New System.Data.DataSet
            Dim Dttbl As New DataTable()
            MyCommand.Fill(Dttbl)

            rowcount = Dttbl.Rows.Count()

            DtSet.Tables.Add(Dttbl)
            DataGridView1.DataSource = DtSet.Tables(0)

            Me.DataGridView1.Columns(0).ReadOnly = True
            Me.DataGridView1.Columns(1).ReadOnly = True
            Me.DataGridView1.Columns(2).ReadOnly = True
            Me.DataGridView1.Columns(3).ReadOnly = True
            Me.DataGridView1.Columns(4).ReadOnly = True
            Me.DataGridView1.Columns(5).ReadOnly = True

            'Dim Sqlq2 As System.Data.OleDb.OleDbDataAdapter
            'Dim Sqlq2_set As System.Data.DataSet
            'Sqlq2 = New System.Data.OleDb.OleDbDataAdapter("select DISTINCT(Class) from [Sheet1$]", MyConnection)
            'Sqlq2.TableMappings.Add("Table", "Net-informations.com")
            'Sqlq2_set = New System.Data.DataSet
            'Dim Dttbl2 As New DataTable()
            'Sqlq2.Fill(Dttbl2)

            'Sqlq2_set.Tables.Add(Dttbl2)


            'ComboBox1.DataSource = Sqlq2_set.Tables(0).DefaultView
            'ComboBox1.DisplayMember = "Class"
            'ComboBox1.BindingContext = Me.BindingContext


            'ComboBox2.DataSource = DtSet.Tables(0).DefaultView
            'ComboBox2.DisplayMember = DtSet.Tables(0).Columns(0).ToString()
            'ComboBox2.BindingContext = Me.BindingContext

            'ComboBox2.Text = ""


            MyConnection.Close()

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

    Private Sub Button2_Click(sender As Object, e As EventArgs)
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form3.Hide()
        Me.Location = New Point(Form1.Left + 240, Form1.Top + 50)
        'load_data()
        Console.WriteLine(My.Settings("ExcelPath"))

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim DtSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = My.Settings("ExcelPath")
            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.16.0;Data Source='" + path + "';Extended Properties=Excel 8.0;")
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$] where [Enrollment No#] LIKE '" + TextBox2.Text + "%'", MyConnection)
            MyCommand.TableMappings.Add("Table", "Net-informations.com")
            DtSet = New System.Data.DataSet
            Dim Dttbl As New DataTable()
            MyCommand.Fill(Dttbl)

            DtSet.Tables.Add(Dttbl)
            DataGridView1.DataSource = DtSet.Tables(0)

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim DtSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = My.Settings("ExcelPath")
            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.16.0;Data Source='" + path + "';Extended Properties=Excel 8.0;")
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$] where [Student Name] LIKE '%" + TextBox1.Text + "%'", MyConnection)
            MyCommand.TableMappings.Add("Table", "Net-informations.com")
            DtSet = New System.Data.DataSet
            Dim Dttbl As New DataTable()
            MyCommand.Fill(Dttbl)

            DtSet.Tables.Add(Dttbl)
            DataGridView1.DataSource = DtSet.Tables(0)

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub

    Function SaveData(r As Integer, c As Integer)


        Dim value As Integer = DataGridView1.Rows(r).Cells(c).Value

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        Dim path As String = My.Settings("ExcelPath")
        xlWorkBook = xlApp.Workbooks.Open(path)
        xlWorkSheet = xlWorkBook.Worksheets("Sheet1")
        xlWorkSheet.Cells(r + 2, c + 1) = value

        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

    End Function


    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged

        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            SaveData(e.RowIndex, e.ColumnIndex)
            MsgBox("Updated.")
        End If

    End Sub
End Class