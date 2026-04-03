Module DbHelper
    Function GetConnectionString(path As String) As String
        Dim providers() As String = {"Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0"}
        For Each prov In providers
            Try
                Dim cs As String = "provider=" & prov & ";Data Source='" & path & "';Extended Properties=Excel 8.0;"
                Dim con As New System.Data.OleDb.OleDbConnection(cs)
                con.Open()
                con.Close()
                Return cs
            Catch ex As Exception
                ' Try next provider
            End Try
        Next
        Throw New Exception(
            "Cannot connect to the Excel data file." & vbCrLf & vbCrLf &
            "The Microsoft Access Database Engine is not installed on this machine." & vbCrLf &
            "Please download and install one of the following:" & vbCrLf &
            "  - Microsoft Access Database Engine 2016 Redistributable" & vbCrLf &
            "  - Microsoft Access Database Engine 2010 Redistributable" & vbCrLf & vbCrLf &
            "Search 'Microsoft Access Database Engine Redistributable' on Microsoft's download center."
        )
    End Function
End Module
