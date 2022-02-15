Imports System.Data.OleDb

Public Class Form5
    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Location = New Point(Form1.Left + 240, Form1.Top + 50)
        TextBox5.Text = generate_Invoice()
        'FillDetails()
    End Sub

    Function UpdateData(admn As String, bal As String)
        Try
            Dim con As New OleDb.OleDbConnection
            con.ConnectionString = "provider=Microsoft.ACE.OLEDB.16.0;Data Source='" + My.Settings("ExcelPath") + "';Extended Properties=Excel 8.0;"
            con.Open()

            Dim query As String = "UPDATE [Sheet1$] SET [Amount Due] = @n1 WHERE [Enrollment No#] = @n2"


            Using cmd = New OleDbCommand(query, con)
                cmd.Parameters.AddWithValue("@n1", bal)
                cmd.Parameters.AddWithValue("@n2", admn)
                cmd.ExecuteNonQuery()
            End Using


            UpdateData = True
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            UpdateData = False
        End Try
    End Function


    Function return_length()
        Dim rowcount As Integer
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim DtSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = My.Settings("ExcelPath") '"C:\Users\ABHIJITH UDAYAKUMAR\OneDrive\Documents\Billing System User Files\data.xlsx"
            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.16.0;Data Source='" + path + "';Extended Properties=Excel 8.0;")
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("Select * from [Sheet2$] where Invoice_No = '" + TextBox5.Text + "'", MyConnection)
            MyCommand.TableMappings.Add("Table", "Net-informations.com")
            DtSet = New System.Data.DataSet
            Dim Dttbl As New DataTable()
            MyCommand.Fill(Dttbl)

            rowcount = Dttbl.Rows.Count()

            return_length = rowcount

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

    End Function


    Function SaveExcel(sender As Object, e As EventArgs, k As Boolean)
        Dim pram As OleDbParameter
        Dim dr As DataRow
        Dim olecon As OleDbConnection
        Dim olecomm As OleDbCommand
        Dim olecomm1 As OleDbCommand
        Dim oleadpt As OleDbDataAdapter
        Dim ds As DataSet
        Try
            olecon = New OleDbConnection
            olecon.ConnectionString = "provider=Microsoft.ACE.OLEDB.16.0;Data Source='" + My.Settings("ExcelPath") + "';Extended Properties=Excel 8.0;"
            olecomm = New OleDbCommand
            'olecomm.CommandText = "Select * from [Sheet2$] where Invoice_No = '" + TextBox5.Text + "'"
            olecomm.CommandText = "Select * from [Sheet2$]"
            olecomm.Connection = olecon
            olecomm1 = New OleDbCommand
            olecomm1.CommandText = "Insert into [Sheet2$] " &
                "(Invoice_No, Invoice_Date, Admission_No, Student_Name, Class, Division, Father, Mother, Balance, Paid, Total, LateFee) values " &
                "(@invno, @invdate, @admnno, @studname, @class, @division, @father, @mother, @bal, @paid, @tot, @late)"
            olecomm1.Connection = olecon
            pram = olecomm1.Parameters.Add("@invno", OleDbType.VarChar)
            pram.SourceColumn = "Invoice_No"
            pram = olecomm1.Parameters.Add("@invdate", OleDbType.VarChar)
            pram.SourceColumn = "Invoice_Date"
            pram = olecomm1.Parameters.Add("@admnno", OleDbType.VarChar)
            pram.SourceColumn = "Admission_No"
            pram = olecomm1.Parameters.Add("@studname", OleDbType.VarChar)
            pram.SourceColumn = "Student_Name"
            pram = olecomm1.Parameters.Add("@class", OleDbType.VarChar)
            pram.SourceColumn = "Class"
            pram = olecomm1.Parameters.Add("@division", OleDbType.VarChar)
            pram.SourceColumn = "Division"
            pram = olecomm1.Parameters.Add("@father", OleDbType.VarChar)
            pram.SourceColumn = "Father"
            pram = olecomm1.Parameters.Add("@mother", OleDbType.VarChar)
            pram.SourceColumn = "Mother"
            pram = olecomm1.Parameters.Add("@bal", OleDbType.VarChar)
            pram.SourceColumn = "Balance"
            pram = olecomm1.Parameters.Add("@paid", OleDbType.VarChar)
            pram.SourceColumn = "Paid"
            pram = olecomm1.Parameters.Add("@tot", OleDbType.VarChar)
            pram.SourceColumn = "Total"
            pram = olecomm1.Parameters.Add("@late", OleDbType.VarChar)
            pram.SourceColumn = "LateFee"

            oleadpt = New OleDbDataAdapter(olecomm)
            ds = New DataSet
            olecon.Open()
            oleadpt.Fill(ds, "Sheet2")

            If return_length() = 0 Then
                If IsNothing(ds) = False Then
                    dr = ds.Tables(0).NewRow
                    dr("Invoice_No") = TextBox5.Text
                    dr("Invoice_Date") = DateTimePicker1.Value
                    dr("Admission_No") = TextBox1.Text
                    dr("Student_Name") = TextBox2.Text
                    dr("Class") = TextBox4.Text
                    dr("Division") = TextBox3.Text
                    dr("Father") = TextBox7.Text
                    dr("Mother") = TextBox6.Text
                    dr("Balance") = TextBox9.Text
                    dr("Paid") = TextBox8.Text
                    dr("Total") = TextBox10.Text
                    dr("LateFee") = TextBox11.Text

                    If UpdateData(TextBox1.Text, TextBox9.Text) Then
                        ds.Tables(0).Rows.Add(dr)
                        oleadpt = New OleDbDataAdapter
                        oleadpt.InsertCommand = olecomm1
                        Dim i As Integer = oleadpt.Update(ds, "Sheet2")
                        'MessageBox.Show(i & " row affected")
                        MsgBox("Record Saved")


                        My.Settings("invoice_latest") = My.Settings("invoice_latest") + 1
                        My.Settings.Save()


                        If k = True Then
                            Button1_Click(sender, e)
                        End If

                    End If
                End If
            Else
                MsgBox("Invoice Number Already Exists !!")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            olecon.Close()
            olecon = Nothing
            olecomm = Nothing
            oleadpt = Nothing
            ds = Nothing
            dr = Nothing
            pram = Nothing
        End Try

    End Function

    Function FillDetails()
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim path As String = My.Settings("ExcelPath")
            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.16.0;Data Source='" + path + "';Extended Properties=Excel 8.0;")

            Dim cmd As New System.Data.OleDb.OleDbCommand("select * from [Sheet1$]", MyConnection)
            Dim da As New System.Data.OleDb.OleDbDataAdapter(cmd)
            Dim dt As New DataSet
            da.Fill(dt)

            Dim column1 As New AutoCompleteStringCollection
            Dim column2 As New AutoCompleteStringCollection
            Dim i As Integer

            Dim admnno As String
            Dim studname As String
            Dim base As String
            For i = 0 To dt.Tables(0).Rows.Count - 1
                base = "-" + dt.Tables(0).Rows(i)(3).ToString + "-" + dt.Tables(0).Rows(i)(2).ToString + "-" + dt.Tables(0).Rows(i)(4).ToString + "-" + dt.Tables(0).Rows(i)(5).ToString + "-" + dt.Tables(0).Rows(i)(6).ToString
                admnno = dt.Tables(0).Rows(i)(0).ToString + "-" + dt.Tables(0).Rows(i)(1).ToString + base
                studname = dt.Tables(0).Rows(i)(1).ToString + "-" + dt.Tables(0).Rows(i)(0).ToString + base
                column1.Add(admnno)
                column2.Add(studname)
            Next

            TextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource
            TextBox1.AutoCompleteCustomSource = column1
            TextBox1.AutoCompleteMode = AutoCompleteMode.Suggest

            TextBox2.AutoCompleteSource = AutoCompleteSource.CustomSource
            TextBox2.AutoCompleteCustomSource = column2
            TextBox2.AutoCompleteMode = AutoCompleteMode.Suggest

            MyConnection.Close()


        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Function

    Function FillSearchDetails()
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim path As String = My.Settings("ExcelPath")
            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.16.0;Data Source='" + path + "';Extended Properties=Excel 8.0;")

            Dim cmd As New System.Data.OleDb.OleDbCommand("select * from [Sheet2$]", MyConnection)
            Dim da As New System.Data.OleDb.OleDbDataAdapter(cmd)
            Dim dt As New DataSet
            da.Fill(dt)

            Dim column1 As New AutoCompleteStringCollection
            Dim column2 As New AutoCompleteStringCollection
            Dim column3 As New AutoCompleteStringCollection
            Dim i As Integer

            Dim invno As String
            Dim admnno As String
            Dim studname As String
            Dim base As String



            For i = 0 To dt.Tables(0).Rows.Count - 1
                base = "$" + dt.Tables(0).Rows(i)(4).ToString + "$" + dt.Tables(0).Rows(i)(5).ToString + "$" + dt.Tables(0).Rows(i)(6).ToString + "$" + dt.Tables(0).Rows(i)(7).ToString + "$" + dt.Tables(0).Rows(i)(8).ToString + "$" + dt.Tables(0).Rows(i)(9).ToString + "$" + dt.Tables(0).Rows(i)(11).ToString
                invno = dt.Tables(0).Rows(i)(0).ToString + "$" + dt.Tables(0).Rows(i)(1).ToString + "$" + dt.Tables(0).Rows(i)(2).ToString + "$" + dt.Tables(0).Rows(i)(3).ToString + base
                admnno = dt.Tables(0).Rows(i)(2).ToString + "$" + dt.Tables(0).Rows(i)(3).ToString + "$" + dt.Tables(0).Rows(i)(0).ToString + "$" + dt.Tables(0).Rows(i)(1).ToString + base
                studname = dt.Tables(0).Rows(i)(3).ToString + "$" + dt.Tables(0).Rows(i)(2).ToString + "$" + dt.Tables(0).Rows(i)(0).ToString + "$" + dt.Tables(0).Rows(i)(1).ToString + base
                column1.Add(invno)
                column2.Add(admnno)
                column3.Add(studname)
            Next

            TextBox5.AutoCompleteSource = AutoCompleteSource.CustomSource
            TextBox5.AutoCompleteCustomSource = column1
            TextBox5.AutoCompleteMode = AutoCompleteMode.Suggest

            TextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource
            TextBox1.AutoCompleteCustomSource = column2
            TextBox1.AutoCompleteMode = AutoCompleteMode.Suggest

            TextBox2.AutoCompleteSource = AutoCompleteSource.CustomSource
            TextBox2.AutoCompleteCustomSource = column3
            TextBox2.AutoCompleteMode = AutoCompleteMode.Suggest

            MyConnection.Close()


        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Function

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If CheckBox1.Checked = True Then
            If CountCharacter(TextBox1.Text, "$") = 10 Then
                Dim strArr() As String
                strArr = TextBox1.Text.Split("$")

                TextBox1.Text = strArr(0)
                TextBox2.Text = strArr(1)
                TextBox5.Text = strArr(2)
                DateTimePicker1.Value = strArr(3)
                TextBox4.Text = strArr(4)
                TextBox3.Text = strArr(5)
                TextBox7.Text = strArr(6)
                TextBox6.Text = strArr(7)
                TextBox9.Text = strArr(8)
                TextBox8.Text = strArr(9)
                TextBox11.Text = strArr(10)


            End If
        Else
            If CountCharacter(TextBox1.Text, "-") = 6 Then
                Dim strArr() As String
                strArr = TextBox1.Text.Split("-")

                TextBox1.Text = strArr(0)
                TextBox2.Text = strArr(1)
                TextBox4.Text = strArr(2)
                TextBox3.Text = strArr(3)
                TextBox7.Text = strArr(4)
                TextBox6.Text = strArr(5)
                TextBox9.Text = FormatNumber(CDec(strArr(6)), 0)
                TextBox10.Text = TextBox9.Text

                TextBox8.Enabled = True

            End If
        End If

    End Sub

    Public Function CountCharacter(ByVal value As String, ByVal ch As Char) As Integer
        Dim cnt As Integer = 0
        For Each c As Char In value
            If c = ch Then
                cnt += 1
            End If
        Next
        Return cnt
    End Function

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If CheckBox1.Checked = True Then
            If CountCharacter(TextBox2.Text, "$") = 10 Then
                Dim strArr() As String
                strArr = TextBox2.Text.Split("$")

                TextBox2.Text = strArr(0)
                TextBox1.Text = strArr(1)
                TextBox5.Text = strArr(2)
                DateTimePicker1.Value = strArr(3)
                TextBox4.Text = strArr(4)
                TextBox3.Text = strArr(5)
                TextBox7.Text = strArr(6)
                TextBox6.Text = strArr(7)
                TextBox9.Text = strArr(8)
                TextBox8.Text = strArr(9)
                TextBox11.Text = strArr(10)


            End If
        Else


            If CountCharacter(TextBox2.Text, "-") = 6 Then
                Dim strArr() As String
                strArr = TextBox2.Text.Split("-")

                TextBox2.Text = strArr(0)
                TextBox1.Text = strArr(1)
                TextBox4.Text = strArr(2)
                TextBox3.Text = strArr(3)
                TextBox7.Text = strArr(4)
                TextBox6.Text = strArr(5)
                TextBox9.Text = FormatNumber(CDec(strArr(6)), 0)
                TextBox10.Text = TextBox9.Text

                TextBox8.Enabled = True


            End If
        End If
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        TextboxOnlyNumbers(sender)

        Try

            If String.IsNullOrEmpty(TextBox10.Text) Then
                Console.WriteLine("Search Mode")
            Else
                Dim fees As Decimal = CDec(TextBox8.Text)
                Dim balance As Decimal = CDec(TextBox10.Text)
                Dim New_balance As Decimal = balance - fees
                TextBox9.Text = New_balance.ToString


                If New_balance >= 0 Then
                    Console.Write("POSITIVE")
                Else
                    MsgBox("Balance Cant be Negative")
                    TextBox9.Text = TextBox10.Text
                    TextBox8.Text = ""
                End If
            End If


        Catch ex As System.InvalidCastException
            Console.WriteLine("EMPTY!!!!!")
            TextBox9.Text = TextBox10.Text
        End Try


    End Sub

    Public Sub TextboxOnlyNumbers(ByRef objTxtBox As TextBox)

        ' ONLY allow numbers
        If Not IsNumeric(objTxtBox.Text) Then

            ' Don't process things like too many backspaces
            If objTxtBox.Text.Length > 0 Then

                MsgBox("Numerical Values only!")

                Try
                    ' If something bad was entered delete the last character
                    objTxtBox.Text = objTxtBox.Text.Substring(0, objTxtBox.Text.Length - 1)

                    ' Put the cursor and the END of the corrected number
                    objTxtBox.Select(objTxtBox.Text.Length + 1, 1)

                Catch ex As Exception
                End Try
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()

        TextBox5.Text = generate_Invoice()

        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()
        TextBox11.Clear()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        SaveExcel(sender, e, True)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If CheckBox1.Checked = False Then
            SaveExcel(sender, e, False)
        End If

        Form7.FillReceipt(TextBox5.Text, TextBox2.Text, TextBox4.Text + "-" + TextBox3.Text, DateTimePicker1.Value.ToString, TextBox7.Text, TextBox1.Text, TextBox8.Text, TextBox9.Text, TextBox11.Text)
        Form7.Show()

        Button1_Click(sender, e)

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = False Then
            TextBox5.ReadOnly = True
            TextBox8.ReadOnly = False
            TextBox5.AutoCompleteMode = AutoCompleteMode.None

            Button2.Visible = True
            'Button1.Visible = True
            Button3.Text = "Save and Print"

            CheckBox2.Checked = False
            CheckBox2.Enabled = True

            Button1_Click(sender, e)

            FillDetails()

        Else
            TextBox5.ReadOnly = False
            TextBox8.ReadOnly = True
            TextBox10.Text = ""

            Button2.Visible = False
            'Button1.Visible = False
            Button3.Text = "Print"

            CheckBox2.Checked = True
            CheckBox2.Enabled = False

            FillSearchDetails()


        End If
    End Sub


    Function generate_Invoice()
        Dim start_str As String = "GCS-"
        Dim latest As Integer

        If My.Settings("invoice_latest") = 0 Then
            My.Settings("invoice_latest") = 9027
            My.Settings.Save()
        End If


        latest = My.Settings("invoice_latest")

        generate_Invoice = start_str + (latest + 1).ToString

    End Function

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If CountCharacter(TextBox5.Text, "$") = 10 Then
            Dim strArr() As String
            strArr = TextBox5.Text.Split("$")

            Console.WriteLine(strArr(8))


            TextBox5.Text = strArr(0)
            DateTimePicker1.Value = strArr(1)
            TextBox1.Text = strArr(2)
            TextBox2.Text = strArr(3)
            TextBox4.Text = strArr(4)
            TextBox3.Text = strArr(5)
            TextBox7.Text = strArr(6)
            TextBox6.Text = strArr(7)
            TextBox9.Text = strArr(8)
            TextBox8.Text = strArr(9)
            TextBox11.Text = strArr(10)


        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Label14.Visible = True
            TextBox11.Visible = True
        Else
            Label14.Visible = False
            TextBox11.Visible = False
        End If
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        TextboxOnlyNumbers(sender)
    End Sub
End Class