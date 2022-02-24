Public Class Form3
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Location = New Point(Form1.Left + 240, Form1.Top + 50)
    End Sub

    Function CalculateDue()
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim path As String = My.Settings("ExcelPath")
            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.16.0;Data Source='" + path + "';Extended Properties=Excel 8.0;")

            Dim dt As New DataTable()
            Using cmd As New System.Data.OleDb.OleDbCommand("select [Amount Due] from [Sheet1$]", MyConnection)
                MyConnection.Open()
                Using dr As System.Data.OleDb.OleDbDataReader = cmd.ExecuteReader()
                    dt.Load(dr)
                End Using
            End Using

            Dim sumVar As Integer = 0

            For rowIndex As Integer = 0 To dt.Rows.Count - 1
                sumVar = sumVar + CInt(dt.Rows(rowIndex)(0))
            Next

            Label3.Text = FormatNumber(sumVar, 0)

            MyConnection.Close()

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

    End Function


End Class