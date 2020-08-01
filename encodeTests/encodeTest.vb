Imports NUnit.Framework
Imports coding
Namespace Tests
    Public Class Coding

        Dim nonASCII = {
             {ChrW(&H7F), "~7f"},
             {ChrW(&H100), "~c4~80"},
             {ChrW(&H1000), "~E1~80~80"},
             {"Jos" & ChrW(&HE9), "Jos~c3~A9"},
             {ChrW(&H201C) & "Word" & ChrW(&H201D), "~e2~80~9CWord~E2~80~9d"}}
        Dim specials = {
            {"""Quotation""", "~22Quotation~22"},
            {"C:\Users", "C:~5cUsers"},
            {"90%", "90~25"}}

        <SetUp>
        Public Sub Setup()
        End Sub

        <Test>
        Public Sub EncodeNonASCII()
            For i As Long = 0 To nonASCII.GetUpperBound(0)
                Assert.AreEqual(LCase(nonASCII(i, 1)), LCase(EncodeString(nonASCII(i, 0))))
            Next
        End Sub

        <Test>
        Public Sub DecodeNonASCII()
            For i As Long = 0 To nonASCII.GetUpperBound(0)
                Assert.AreEqual(nonASCII(i, 0), DecodeString(nonASCII(i, 1)))
            Next
        End Sub

        <Test>
        Public Sub EncodeSpecials()
            For i As Long = 0 To specials.GetUpperBound(0)
                Assert.AreEqual(LCase(specials(i, 1)), LCase(EncodeString(specials(i, 0))))
            Next
        End Sub

        <Test>
        Public Sub DecodeSpecials()
            For i As Long = 0 To specials.GetUpperBound(0)
                Assert.AreEqual(specials(i, 0), DecodeString(specials(i, 1)))
            Next
        End Sub

    End Class
End Namespace
