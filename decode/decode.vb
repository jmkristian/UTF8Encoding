Imports coding

Module Decode
    Declare Function SetConsoleOutputCP Lib "kernel32" (ByVal wCodePageID As Integer) As Integer
    Const CodePageUTF8 = 65001
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
