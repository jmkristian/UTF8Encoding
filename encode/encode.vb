Imports coding
Module Encode
    Sub Main()
        ' Encode the command line arguments.
        Dim clArgs() As String = Environment.GetCommandLineArgs()
        For i As Integer = 1 To clArgs.GetUpperBound(0)
            Console.WriteLine(EncodeString(clArgs(i)))
        Next
    End Sub
End Module
