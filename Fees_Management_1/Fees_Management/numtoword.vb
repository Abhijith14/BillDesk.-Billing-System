Public Class numtoword
    Public Const MAX_DECIMAL_VALUE As Decimal = 2147483647.2147483647D
    Public Shared DecimalSeparator As String = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator
    Public Shared GroupSeparator As String = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator
    Public Shared SpaceString As String = " "
    Public Shared AndString As String = "and"
    Public Shared DashString As String = "-"
    Public Shared DecimalString As String = "point"
    Public Shared NegativeString As String = "negative"

    Protected Sub New()
    End Sub


    Public Enum RootNumbers
        Zero
        One
        Two
        Three
        Four
        Five
        Six
        Seven
        Eight
        Nine
        Ten
        Eleven
        Twelve
        Thirteen
        Fourteen
        Fifteen
        Sixteen
        Seventeen
        Eighteen
        Nineteen
        Twenty
        Thirty = 30
        Forty = 40
        Fifty = 50
        Sixty = 60
        Seventy = 70
        Eighty = 80
        Ninety = 90
        Hundred = 100
        Thousand = 1000
        Million = 1000000
        Billion = 1000000000
    End Enum

    Public Shared Function GetRootNumberWord(ByVal number As Integer) As String
        Return [Enum].GetName(GetType(RootNumbers), number)
    End Function

    Public Shared Function GetNumberWords(ByVal number As Integer) As String
        If number = 0 Then Return GetRootNumberWord(0)

        If number < 0 Then
            Return NegativeString & SpaceString & GetNumberWords(System.Math.Abs(number))
        End If

        Dim result As New System.Text.StringBuilder
        Dim digitIndex As Integer = 9
        While digitIndex > 1
            Dim digitValue As Integer = CInt(10 ^ digitIndex)
            If number \ digitValue > 0 Then
                result.Append(GetNumberWords(number \ digitValue))
                result.Append(SpaceString)
                result.Append(GetRootNumberWord(digitValue))
                result.Append(SpaceString)
                number = number Mod digitValue
            End If

            If digitIndex = 9 Then
                digitIndex = 6
            ElseIf digitIndex = 6 Then
                digitIndex = 3
            ElseIf digitIndex = 3 Then
                digitIndex = 2
            Else
                digitIndex = 0
            End If
        End While

        If number > 0 Then
            If result.Length > 0 Then
                result.Append(AndString)
                result.Append(SpaceString)
            End If

            If number < 20 Then
                result.Append(GetRootNumberWord(number))
            Else
                result.Append(GetRootNumberWord((number \ 10) * 10))
                Dim modTen As Integer = number Mod 10
                If modTen > 0 Then
                    result.Append(DashString)
                    result.Append(GetRootNumberWord(modTen))
                End If
            End If
        End If

        Return result.ToString
    End Function

End Class
