Imports System.Text.RegularExpressions
Public Module Coding
    Const marker = "~"
    Const special = marker & Chr(34) & "\%"

    Public Function EncodeString(ByVal instring As String) As String
        ' Description:  This function encodes characters so they are correctly passed
        '               during Shell and ShellExecute calls.  Double-quotation marks ("), backslash (\),
        '               and percent sign (%) are trapped and converted to "~<hex_value>"
        ' Returns:      encoded string
        ' Revision:     7/25/2020:  #1605, port VBA code from John W6JMK
        ' Revision:     7/29/2020:  UTF-8 encoding


        Dim utf8 As New System.Text.UTF8Encoding ' no BOM (byte order mark)
        Dim bytes As Byte()
        Dim i As Long
        Dim n As Integer
        Dim NewString As String
        Dim OneChar                         ' Resolves as a variant

        NewString = ""
        bytes = utf8.GetBytes(instring)
        For i = 0 To bytes.GetUpperBound(0) ' loop through each byte
            n = CInt(bytes(i))
            OneChar = ChrW(n)
            If n < 16 Then                  ' for single digit Hex numbers, pad with a "0"
                NewString = NewString & marker & "0" & Hex(n)

                ' for non-ASCII characters, add the marker and Hex value
            ElseIf n < 32 Or n > 126 Or InStr(special, OneChar) Then
                NewString = NewString & marker & Hex(n)

            Else                            ' for everything else, just add the character
                NewString = NewString & OneChar
            End If
        Next
        Return NewString
    End Function

    Public Function DecodeString(ByVal from As String) As String
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
End Module
