Attribute VB_Name = "modFormat"
Option Explicit

' Format a currency value representing a size in bytes into a string for display
Const cibytesLimit As Currency = 512
Const ciKLimit As Currency = cibytesLimit * 1024
Const ciMLimit As Currency = ciKLimit * 1024
Public Function FormatSize(iSize As Currency) As String
Dim fSize As Double
Dim sPostFix As String
    If iSize > cibytesLimit Then
        If iSize > ciMLimit Then
            fSize = iSize / 1024 / 1024 / 1024
            sPostFix = "G"
        ElseIf iSize > ciKLimit Then
            fSize = iSize / 1024 / 1024
            sPostFix = "M"
        ElseIf iSize > cibytesLimit Then
            fSize = iSize / 1024
            sPostFix = "K"
        Else
            fSize = iSize
            sPostFix = ""
        End If
        If fSize > 99.9 Then
            FormatSize = Format(fSize, "0") & sPostFix
        ElseIf fSize > 9.99 Then
            FormatSize = Format(fSize, "0.0") & sPostFix
        ElseIf fSize > 0.999 Then
            FormatSize = Format(fSize, "0.00") & sPostFix
        Else
            FormatSize = Format(fSize, "0.00") & sPostFix
        End If
    Else
        FormatSize = iSize
    End If
End Function

