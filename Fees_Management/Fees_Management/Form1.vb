Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dash_icon As New Bitmap(New Bitmap(My.Resources.dash_icon), 30, 30)
        Dim stud_icon As New Bitmap(New Bitmap(My.Resources.stud_data), 30, 30)
        Dim bill_icon As New Bitmap(New Bitmap(My.Resources.billing), 30, 30)
        Dim sett_icon As New Bitmap(New Bitmap(My.Resources.settings), 30, 30)

        Button1.Image = dash_icon
        Button1.ImageAlign = ContentAlignment.MiddleCenter

        Button2.Image = stud_icon
        Button2.ImageAlign = ContentAlignment.MiddleCenter

        Button4.Image = bill_icon
        Button4.ImageAlign = ContentAlignment.MiddleCenter

        Button3.Image = sett_icon
        Button3.ImageAlign = ContentAlignment.MiddleCenter



        If String.IsNullOrEmpty(My.Settings("ExcelPath")) Then
            Button1.Enabled = False
            Button2.Enabled = False
            Button4.Enabled = False
        Else
            Form3.Show()
        End If

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        Form6.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form2.Hide()
        Form4.Hide()
        Form5.Hide()
        Form3.Show()

        Form3.CalculateDue()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form3.Hide()
        Form4.Hide()
        Form5.Hide()
        Form2.Show()
    End Sub

    Private Sub Form1_LocationChanged(sender As Object, e As EventArgs) Handles MyBase.LocationChanged, GroupBox2.MouseHover, GroupBox1.MouseHover

        If Form2.Visible Then
            Form2.Location = New Point(Me.Left + 240, Me.Top + 50)
            Form2.BringToFront()

        ElseIf Form3.Visible Then
            Form3.Location = New Point(Me.Left + 240, Me.Top + 50)
            Form3.BringToFront()

        ElseIf Form4.Visible Then
            Form4.Location = New Point(Me.Left + 240, Me.Top + 50)
            Form4.BringToFront()

        ElseIf Form5.Visible Then
            Form5.Location = New Point(Me.Left + 240, Me.Top + 50)
            Form5.BringToFront()
        End If

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form2.Hide()
        Form3.Hide()
        Form5.Hide()
        Form4.reload_inv()
        Form4.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form2.Hide()
        Form3.Hide()
        Form4.Hide()
        Form5.TextBox5.Text = Form5.generate_Invoice()
        Form5.Show()

    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Form6.Close()
    End Sub
End Class
