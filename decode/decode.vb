Imports System.Text.RegularExpressions

Module decode
    Declare Function SetConsoleOutputCP Lib "kernel32" (ByVal wCodePageID As Integer) As Integer
    Const CodePageUTF8 = 65001
    Private Function DecodeString(ByVal from As String) As String
        Const marker = "~"
        Dim encodedByte As Regex = New Regex(marker & "[0-9a-fA-F][0-9a-fA-F]")
        Dim utf8 As New System.Text.UTF8Encoding ' no BOM (byte order mark)
        Dim bytes() As Byte = {}
        Dim wasDecoded As Boolean = False
        Dim b As Byte

        Dim f As Long = 1
        Do While f <= Len(from)
            If encodedByte.IsMatch(Mid(from, f, 3)) Then
                b = CByte("&h" & Mid(from, f + 1, 2))
                wasDecoded = True
                f += 3
            Else
                b = AscW(Mid(from, f, 1))
                f += 1
            End If
            ' Append b to bytes:
            ReDim Preserve bytes(UBound(bytes) + 1)
            bytes(UBound(bytes)) = b
        Loop
        If wasDecoded Then
            DecodeString = utf8.GetString(bytes)
        Else
            DecodeString = from
        End If
        Exit Function
    End Function

    Sub Main()
        ' Decode the command line arguments, and output the results.
        SetConsoleOutputCP(CodePageUTF8) ' Don't change characters to 'best fit' ASCII.
        Dim args() As String = Environment.GetCommandLineArgs()
        Dim decoded As String
        Dim codePoint As String
        For a As Integer = 1 To args.GetUpperBound(0)
            decoded = DecodeString(args(a))
            Console.WriteLine(decoded)
            ' To output Unicode code points:
            ' For d As Integer = 0 To Len(decoded) - 1
            '   codePoint = Hex(AscW(decoded(d)))
            '   Do While Len(codePoint) < 4 ' 0-pad to at least 4 digits
            '     codePoint = "0" & codePoint
            '   Loop
            '   If (d > 0) Then
            '     Console.Write(" ")
            '   End If
            '   Console.Write("U+" & codePoint)
            ' Next
            ' Console.WriteLine("")
        Next
    End Sub
End Module
