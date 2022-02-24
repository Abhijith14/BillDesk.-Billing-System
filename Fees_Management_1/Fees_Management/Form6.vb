Public Class Form6
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If String.IsNullOrEmpty(My.Settings("ExcelPath")) Then
            Console.WriteLine("Done!")
        Else
            Try
                Form2.load_data()
                Console.WriteLine("Doneaaaa!")
                Form5.FillDetails()
                Form3.CalculateDue()
            Catch ex As Exception
                My.Settings("ExcelPath") = ""
                My.Settings.Save()

                Form1.Button1.Enabled = False
                Form1.Button2.Enabled = False
                Form1.Button4.Enabled = False
            End Try
        End If
        Timer1.Stop()
        Me.Hide()
        Form1.Show()
    End Sub

    Private Sub wait(ByVal seconds As Integer)
        For i As Integer = 0 To seconds * 100
            System.Threading.Thread.Sleep(10)
            Application.DoEvents()
        Next
    End Sub
End Class