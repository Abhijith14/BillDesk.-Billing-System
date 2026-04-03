Public Class Form1
    Private _currentChildForm As Form = Nothing
    Private _activeButton As Button = Nothing

    ' UI accent color for active nav button — soft light blue tint
    Private ReadOnly ActiveColor As Color = Color.FromArgb(218, 232, 252)
    Private ReadOnly ActiveBorderColor As Color = Color.FromArgb(53, 95, 221)
    Private ReadOnly DefaultColor As Color = Color.White
    Private ReadOnly HoverColor As Color = Color.FromArgb(238, 244, 255)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.DoubleBuffered = True
        Me.WindowState = FormWindowState.Maximized

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

        ' Add hover effects to all nav buttons
        For Each btn As Button In New Button() {Button1, Button2, Button3, Button4}
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 2
            btn.FlatAppearance.BorderColor = Color.FromArgb(220, 225, 235)
            btn.Cursor = Cursors.Hand

            AddHandler btn.MouseEnter, Sub(s, ev)
                                           Dim b = DirectCast(s, Button)
                                           If b IsNot _activeButton Then
                                               b.BackColor = HoverColor
                                           End If
                                       End Sub
            AddHandler btn.MouseLeave, Sub(s, ev)
                                           Dim b = DirectCast(s, Button)
                                           If b IsNot _activeButton Then
                                               b.BackColor = DefaultColor
                                           End If
                                       End Sub
        Next

        ' Initialize data (moved from Form6)
        If String.IsNullOrEmpty(My.Settings("ExcelPath")) Then
            Button1.Enabled = False
            Button2.Enabled = False
            Button4.Enabled = False
        Else
            Try
                Form2.load_data()
                Form5.FillDetails()
                Form3.CalculateDue()
            Catch ex As Exception
                My.Settings("ExcelPath") = ""
                My.Settings.Save()
                Button1.Enabled = False
                Button2.Enabled = False
                Button4.Enabled = False
            End Try
            ShowChildForm(Form3)
            SetActiveButton(Button1)
        End If
    End Sub

    Private Sub SetActiveButton(btn As Button)
        ' Reset previous active button
        If _activeButton IsNot Nothing Then
            _activeButton.BackColor = DefaultColor
            _activeButton.FlatAppearance.BorderColor = Color.FromArgb(220, 225, 235)
            _activeButton.ForeColor = Color.FromArgb(53, 95, 221)
        End If

        ' Highlight the new active button — soft tint with blue left border
        btn.BackColor = ActiveColor
        btn.FlatAppearance.BorderColor = ActiveBorderColor
        btn.ForeColor = Color.FromArgb(30, 60, 160)
        _activeButton = btn
    End Sub

    Private Sub ShowChildForm(childForm As Form)
        If _currentChildForm IsNot Nothing Then
            _currentChildForm.Hide()
        End If

        childForm.TopLevel = False
        childForm.FormBorderStyle = FormBorderStyle.None
        childForm.Dock = DockStyle.Fill

        If Not contentPanel.Controls.Contains(childForm) Then
            contentPanel.Controls.Add(childForm)
        End If

        childForm.BringToFront()
        childForm.Show()
        _currentChildForm = childForm
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        If MessageBox.Show("Are you sure you want to exit?", "BillDesk", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Me.Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        SetActiveButton(Button1)
        ShowChildForm(Form3)
        Form3.CalculateDue()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        SetActiveButton(Button2)
        ShowChildForm(Form2)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        SetActiveButton(Button3)
        Form4.reload_inv()
        ShowChildForm(Form4)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        SetActiveButton(Button4)
        Form5.TextBox5.Text = Form5.generate_Invoice()
        ShowChildForm(Form5)
    End Sub
End Class
