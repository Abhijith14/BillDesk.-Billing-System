Public Class Form6
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If String.IsNullOrEmpty(My.Settings("ExcelPath")) Then
            Console.WriteLine("Done!")
            'wait(2)
            'For i As Integer = 0 To 1000000000
            ' Console.WriteLine()
            ' Next
        Else
            Form2.load_data()
            Form5.FillDetails()
            Form3.CalculateDue()
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