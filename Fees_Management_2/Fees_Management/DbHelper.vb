Imports Microsoft.Win32

Module DbHelper
    Function GetConnectionString(path As String) As String
        Dim providers() As String = {"Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0"}
        For Each prov In providers
            If IsProviderRegistered(prov) Then
                Return "provider=" & prov & ";Data Source='" & path & "';Extended Properties=Excel 8.0;"
            End If
        Next
        Throw New Exception(
            "The Microsoft Access Database Engine is not installed on this machine." & vbCrLf & vbCrLf &
            "Please download and install one of the following:" & vbCrLf &
            "  - Microsoft Access Database Engine 2016 Redistributable" & vbCrLf &
            "  - Microsoft Access Database Engine 2010 Redistributable" & vbCrLf & vbCrLf &
            "Search 'Microsoft Access Database Engine Redistributable' on Microsoft's download center."
        )
    End Function

    Private Function IsProviderRegistered(providerName As String) As Boolean
        Try
            Dim key As RegistryKey = Registry.ClassesRoot.OpenSubKey(providerName)
            If key IsNot Nothing Then
                key.Close()
                Return True
            End If
        Catch
        End Try
        Return False
    End Function
End Module
