Imports System.IO

Public Class LinkBuilder
    Dim base As String = ""
    Dim gotos As String = ""
    Public Sub New()

    End Sub
    'Public Sub New(ByRef baseStringIn As String, ByRef gotoString As String)
    '    base = baseStringIn
    '    gotos = gotoString
    'End Sub



    Public Function buildLink(ByRef baseStringIn As String, ByRef gotoString As String) As String


        Dim NothingOverHere As String = "There was no file by that name"
        Dim fileName As String = baseStringIn.Substring(baseStringIn.LastIndexOf("\"), 8) & "SLD.pdf" ' you get the "\" infront
        Dim LinkToReturn = gotoString & fileName
        ' the reason for checking twice, is that this is a class 
        ' that could be used in another program at some point
        ' thats why I made it fully featured
        ' if the actual word document exists
        If File.Exists(baseStringIn) Then

            'if the link to the pdf exists
            If File.Exists(LinkToReturn) Then
                Return LinkToReturn
            Else
                ' if the inital file does not exist
                Return NothingOverHere
            End If
        Else
            ' if the inital file does not exist
            Return NothingOverHere
        End If


    End Function
End Class
