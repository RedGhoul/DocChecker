Imports System
Imports System.IO
Imports Microsoft.Office.Interop.Word

Public Class LoggerInHtml

    Dim placeHolder As String = ""
    Dim writer As StreamWriter
    Dim numOfFilesNotFound As Integer = 0
    Dim numOfbackupwasalreadythere As Integer = 0
    Dim numOfLinksMade As Integer = 0
    Dim numOfLinksMadeBadly As Integer = 0
    Dim numOfTextNotFound As Integer = 0
    Dim numOfChangedFiles As Integer = 0
    Dim ErrorLogWriter As HTML
    Dim numOfTitleRowsAllowed As Integer = 0
    Dim numOfDocs As Integer = 0
    Dim BarPairs As ArrayList = New ArrayList()
    Dim tempdocname As String = ""
    '              where your going to save everything, What file type you want to print to, Title of the error doc, Sub title of the error doc
    Public Sub New(ByRef locationOfLog As String, ByRef fileType As String, ByRef title As String, ByRef SourceFileLocation As String, ByRef same As Boolean) ', ByRef dates As String)
        ' have to put this some where, not here in memeory
        ' cause you need to save it so it can keep writing to the same file

        Dim r As Random = New Random

        'assigning that HTML writer to 
        ErrorLogWriter = New HTML(title, SourceFileLocation, locationOfLog)

        If same = True Then

            placeHolder = locationOfLog & "\WordDocCheckLOG" & ".html"
            'Dim writer2 As StreamWriter = New StreamWriter(placeHolder)
            'writer2.WriteLine(" ")
            'writer2.Close()
        Else

            Select Case fileType
                Case ".txt"
                    placeHolder = locationOfLog & "\WordDocCheckLOG.txt"
                    Dim writer2 As StreamWriter = New StreamWriter(placeHolder)
                    writer2.WriteLine(" ")
                    writer2.Close()
                Case ".html"
                    placeHolder = locationOfLog & "\WordDocCheckLOG.html"
                    Dim writer2 As StreamWriter = New StreamWriter(placeHolder)
                    writer2.WriteLine(" ")
                    writer2.Close()
                Case Else
                    MessageBox.Show("Would you like to re-start ?")
            End Select
        End If

    End Sub

    'This one is redundant! keeping here just cause we may need it in the furture
    Public Function recInTxt(ByRef callsign As String, ByRef fileinfo As ArrayList) As Boolean
        'usually this here would be done in the Makehear function
        writer = File.AppendText(placeHolder)
        ' this should also record each link that it modified
        writer.WriteLine("The file:" & Convert.ToString(fileinfo(0)))
        writer.WriteLine("Link to the file:" & Convert.ToString(fileinfo(1)))
        Select Case callsign
            Case "backupwasalreadythere" ' there was no link in the file so I put a link in it 
                writer.WriteLine("This was seen at: " & Date.Now)
                writer.WriteLine("-----------------------------------")
                numOfbackupwasalreadythere = numOfbackupwasalreadythere + 1
            Case "nofilewasfound" ' there was no link in the file so I put a link in it 
                writer.WriteLine("This was seen at: " & Date.Now)
                writer.WriteLine("-----------------------------------")
                numOfFilesNotFound = numOfFilesNotFound + 1
            Case "nolinksomadealink" ' there was no link in the file so I put a link in it 
                writer.WriteLine("This file did not have a hyper link. This is the hyperlink that was used:")
                writer.WriteLine(Convert.ToString(fileinfo(2)))
                writer.WriteLine("This was done at: " & Date.Now)
                writer.WriteLine("-----------------------------------")
                numOfLinksMade = numOfLinksMade + 1
            Case "nothingtolinkto" 'there was nolink and there was no pdf to link to
                writer.WriteLine("did not have a hyper link. Tried to use the link below")
                writer.WriteLine(Convert.ToString(fileinfo(2)))
                writer.WriteLine(", but the file did not exist. So I just left it the way it is.")
                writer.WriteLine("done at: " & Date.Now)
                writer.WriteLine("-----------------------------------")
                numOfLinksMadeBadly = numOfLinksMadeBadly + 1
            Case "notextwasfoundinfile" 'the "Link to SLD" was not present in the document usually happens if it is not found
                writer.WriteLine("The text wanted to find was no in this doc")
                writer.WriteLine("link temp link used instead was: " & Convert.ToString(fileinfo(2)))
                writer.WriteLine("done at: " & Date.Now)
                writer.WriteLine("-----------------------------------")
                numOfTextNotFound = numOfTextNotFound + 1
            Case "backupwasalreadythere" ' there was no link in the file so I put a link in it 
                writer.WriteLine("This was seen at: " & Date.Now)
                numOfbackupwasalreadythere = numOfbackupwasalreadythere + 1
            Case Else
        End Select
        Return True
    End Function

    Public Function recInHtml(ByRef callsign As String, ByRef fileinfo As ArrayList) As Boolean
        'usually this here would be done in the Makeheader function
        ' but your making the header over here
        If numOfTitleRowsAllowed < 1 Then
            ErrorLogWriter.makeTitleRowTable("File Name", "Link to Old File", "Link to New File", "SLD that was linked to", "Replace Old File with New")
        End If

        Select Case callsign
            'This one is kin of redundant! keeping here just cause we may need it in the furture
            Case "backupwasalreadythere" ' there was no link in the file so I put a link in it
                'name of the file\\ name of the DOC plus the link to the file\\ The change that was made\\ Time it was made
                ErrorLogWriter.makeRowTable(fileinfo(0), "", fileinfo(0) & ".doc", fileinfo(1), "There was something wrong with your dir", "") ',Date.Now)
                numOfbackupwasalreadythere = numOfbackupwasalreadythere + 1
            Case "nofilewasfound" ' there was no link in the file so I put a link in it 
                'name of the file\\ name of the DOC plus the link to the file\\ The change that was made\\ Time it was made
                ErrorLogWriter.makeRowTable(fileinfo(0), "", fileinfo(0) & ".doc", fileinfo(1), "SLD Doc file missing ", "") ', Date.Now)
                numOfFilesNotFound = numOfFilesNotFound + 1
            Case "nolinksomadealink" ' there was no link in the file so I put a link in it 
                'this is if everything went file
                'name of the file\\ name of the DOC plus the link to the file\\ The change that was made\\ Time it was made
                If tempdocname <> fileinfo(0) Then
                    ErrorLogWriter.makeRowTable(fileinfo(0), "", fileinfo(0) & ".doc", fileinfo(1), "SLD linked to ", fileinfo(2), fileinfo(3)) ', Date.Now)
                End If
                tempdocname = fileinfo(0)
                numOfLinksMade = numOfLinksMade + 1
            Case "nothingtolinkto" 'there was nolink and there was no pdf to link to
                'this is if everything went file
                'name of the file\\ name of the DOC plus the link to the file\\ The change that was made\\ Time it was made
                ErrorLogWriter.makeRowTable(fileinfo(0), "", fileinfo(0) & ".doc", fileinfo(1), "SLD PDF file missing ", "") ', Date.Now)
                numOfLinksMadeBadly = numOfLinksMadeBadly + 1
            Case "notextwasfoundinfile" 'the "Link to SLD" was not present in the document usually happens if it is not found
                'this is if everything went file
                'name of the file\\ name of the DOC plus the link to the file\\ The change that was made\\ Time it was made
                ErrorLogWriter.makeRowTable(fileinfo(0), "", fileinfo(0) & ".doc", fileinfo(1), "Link to SLD text missing", "") ', Date.Now)
                numOfTextNotFound = numOfTextNotFound + 1
            Case Else
                ' should put something here just incase
        End Select
        numOfTitleRowsAllowed = numOfTitleRowsAllowed + 1
        Return True
    End Function

    Public Function close(ByRef value As String, ByRef filesin As Integer) As Boolean
        If value.ToUpper = "HTMLCLOSE" Then
            ' do the final tallies over here
            ' close the html header here
            Dim rightMan As StreamWriter = File.AppendText(placeHolder)
            rightMan.WriteLine(ErrorLogWriter.makeRows(filesin))
            'writer.WriteLine(ErrorLogWriter.makeagraph())
            rightMan.Close()
            Return True
        Else
            ' do some stuff here at some point
            Return True
        End If

        Return True
    End Function

    Public Function shutDown(ByRef value As String, ByRef NumQs As Integer, Optional ByRef Numchanges As Integer = 0, Optional ByRef t As Boolean = True) As Boolean
        If value.ToUpper = "HTMLCLOSE" Then
            ' do the final tallies over here
            ' close the html header here

            writer = New StreamWriter(placeHolder)
            'writer.AutoFlush = True

            BarPairs.Clear()

            BarPairs.Add("Files that have been Scanned for Errors: , " & NumQs.ToString())
            If numOfbackupwasalreadythere > 0 Then
                BarPairs.Add("dir errors: , " & numOfbackupwasalreadythere.ToString())
            End If
            If numOfFilesNotFound > 0 Then
                BarPairs.Add("SLD Doc files that were missing, " & numOfFilesNotFound.ToString())
            End If
            BarPairs.Add("Changed Files: , " & Numchanges.ToString())
            BarPairs.Add("New Links Made: , " & numOfLinksMade.ToString())
            BarPairs.Add("SLDs that were Missing: ," & numOfLinksMadeBadly.ToString())
            BarPairs.Add("'" & "Link to SLD" & "'" & "or" & "'" & "Link to Single Line Diagram" & "'" & " text missing: , " & numOfTextNotFound.ToString())

            writer.WriteLine(ErrorLogWriter.makeDoc(ErrorLogWriter.stats(BarPairs), t))
            'writer.WriteLine(ErrorLogWriter.makeDoc("stats"))
            'writer.WriteLine(ErrorLogWriter.makeagraph())

            writer.Close()


            Return True
        Else
            ' do some stuff here at some point
            Return True
        End If

        Return True
    End Function
End Class
