Imports Fees_Management.numtoword
Public Class Form7
    Function FillReceipt(Rno As String, Name As String, Class_Div As String, Date_r As String, Father As String, Adno As String, Paid As String, Balance As String, LateF As String)
        Label1.Text = Rno
        Label2.Text = Name
        Label3.Text = Class_Div
        Label4.Text = Date_r.Split(" ")(0)
        Label5.Text = Father
        Label6.Text = Adno

        Dim Total As Integer = Integer.Parse(Paid) + Integer.Parse(Balance)

        Label11.Text = Total.ToString + ".00"
        Label13.Text = Paid + ".00"
        Label15.Text = Balance + ".00"

        Label12.Text = LateF + ".00"
        Label14.Text = LateF + ".00"




        Dim Final_due As Integer = Total + Integer.Parse(LateF)
        Label18.Text = Final_due.ToString + ".00"

        Dim Final_paid As Integer = Integer.Parse(Paid) + Integer.Parse(LateF)
        Label19.Text = Final_paid.ToString + ".00"

        Label20.Text = Balance + ".00"


        Label21.Text = Final_paid.ToString + ".00"

        Label22.Text = numtoword.GetNumberWords(Decimal.Parse(Final_paid)) + " Rupees Only"


        Label23.Text = Date_r
    End Function

    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
    End Sub

End Class