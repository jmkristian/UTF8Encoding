Imports NUnit.Framework
Imports encode
Imports decode

Namespace Tests
    Public Class Coding

        Dim testCases As New Collections.Hashtable

        <SetUp>
        Public Sub Setup()
            testCases("nonASCII") = {
                {ChrW(&H7F), "~7f"},
                {ChrW(&H100), "~c4~80"},
                {ChrW(&H1000), "~E1~80~80"},
                {ChrW(&H3432), "~E3~90~B2"},
                {ChrW(&HFF38), "~ef~bc~b8"},
                {"Jos" & ChrW(&HE9), "Jos~c3~A9"},
                {ChrW(&HFEFF) & "A", "~eF~Bb~bfA"},
                {ChrW(&H201C) & "Word" & ChrW(&H201D), "~e2~80~9CWord~E2~80~9d"}}
            testCases("specials") = {
                {"""Quotation""", "~22Quotation~22"},
                {"c:\Users", "c:~5cUsers"},
                {"90%", "90~25"},
                {"~5", "~7E5"}}
        End Sub

        <TestCase("nonASCII")>
        <TestCase("specials")>
        Public Sub Encode(t As String)
            Dim testCase = testCases(t)
            For i As Long = 0 To testCase.GetUpperBound(0)
                Assert.AreEqual(LCase(testCase(i, 1)), LCase(EncodeString(testCase(i, 0))))
            Next
        End Sub

        <TestCase("nonASCII")>
        <TestCase("specials")>
        Public Sub Decode(t As String)
            Dim testCase = testCases(t)
            For i As Long = 0 To testCase.GetUpperBound(0)
                Assert.AreEqual(testCase(i, 0), DecodeString(testCase(i, 1)))
            Next
        End Sub

    End Class
End Namespace
