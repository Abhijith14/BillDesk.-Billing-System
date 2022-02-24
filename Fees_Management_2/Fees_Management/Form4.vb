Public Class Form4
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Location = New Point(Form1.Left + 240, Form1.Top + 50)


        If String.IsNullOrEmpty(My.Settings("ExcelPath")) Then
            Console.WriteLine("EMPTY")
        Else
            TextBox1.Text = My.Settings("ExcelPath")
        End If

        reload_inv()


        'Dim s As String = "00001" 'My.Settings("ExcelPath")
        'Dim i As Integer = 15
        'Console.WriteLine(CStr(CInt(s) + i).PadLeft(5, "0"c)) 'outputs 0002


    End Sub


    Function reload_inv()
        If String.IsNullOrEmpty(My.Settings("invoice_latest")) Then
            My.Settings("invoice_latest") = "00001"
            My.Settings.Save()
        End If

        TextBox2.Text = My.Settings("invoice_latest")
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try

            OpenFileDialog1.Title = "Select the Excel Data File"
            OpenFileDialog1.FileName = String.Empty
            OpenFileDialog1.ShowHelp = False
            OpenFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"

            Dim _result As DialogResult = OpenFileDialog1.ShowDialog()

            If _result.Equals(DialogResult.Cancel) Then
                Console.WriteLine("CANCELL")
            Else

                Console.WriteLine(My.Settings("ExcelPath"))
                TextBox1.Text = OpenFileDialog1.FileName

                My.Settings("ExcelPath") = OpenFileDialog1.FileName
                My.Settings.Save()

            End If




        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        My.Settings("ExcelPath") = ""
        My.Settings.Save()

        TextBox1.Text = ""

        Form1.Button1.Enabled = False
        Form1.Button2.Enabled = False
        Form1.Button4.Enabled = False


        My.Settings("invoice_latest") = "00001"
        My.Settings.Save()


        MsgBox("Cleared Data !")

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Form5.TextboxOnlyNumbers(sender)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        My.Settings("invoice_latest") = TextBox2.Text
        My.Settings.Save()

    End Sub
End Class