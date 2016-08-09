Imports System
Imports System.IO
Imports Microsoft.Office.Interop.Word

Public Class Worker
    Dim logFileLocaition As String = ""
    Dim lookUpDirectory As String = ""
    Dim TextToFind As String = ""
    Dim File_Path As String = ""
    Dim ProcessedFiles As Integer = 0
    Dim FilesChanged As Integer = 0
    Dim SkippedFiles As Integer = 0
    Dim timecre As Date

    Public Sub New(ByRef log_place As String, ByRef Sourcedir As String, ByRef File_Paths As String, ByRef textToFindIn As String)
        logFileLocaition = log_place
        TextToFind = textToFindIn
        lookUpDirectory = Sourcedir
        File_Path = File_Paths
    End Sub

    Public Function work(ByRef DP1 As DateTimePicker, ByRef DP2 As DateTimePicker, ByRef List As ListBox, ByRef textboxr As TextBox) As Boolean
        'Dim times As String = Format(Date.Now(), "yyyyBdd") & "B" & Format(Date.Now(), "hhBmmBss")
        Dim wordApplication As New Microsoft.Office.Interop.Word.Application
        Dim WordDocument As Microsoft.Office.Interop.Word.Document
        Dim fileInfoHolder As ArrayList = New ArrayList() ' holder that holds the relavent info of the document 
        'Dim lookUpDirectory As String = "C:\Projects\VBAapps\Sandbox\"
        Dim hyperLinkText As String = ""
        Dim loggingDevice As LoggerInHtml = New LoggerInHtml(logFileLocaition, ".html", "Schedule ADF Aduit Report", "Source Folder: " & lookUpDirectory, True) ', times)
        Dim linkingDevice As LinkBuilder = New LinkBuilder()

        Dim fileInfo As FileInfo

        Dim tempStr As String = ""

        For Each dir As String In Directory.GetDirectories(lookUpDirectory, "*-SCH")


            For Each fileIn As String In Directory.GetFiles(dir & "\SCH D\WORKING", "5*-DCA.doc")

                ' we make it right here so it genrates a new one of the logger everytime is goes into the DIR 

                ' getting the inital file info 
                fileInfo = New FileInfo(fileIn)


                'putting the file location into the list box on the GUI
                List.Items.Add(fileIn)

                If fileInfo.LastWriteTime < DP1.Value And fileInfo.LastWriteTime >= DP2.Value Then

                    'getting the numbers to make the file name
                    Dim filename As String = fileIn.Substring(fileIn.LastIndexOf("\"), 11).Replace("\", "")
                    'making the hyper link over here use:
                    '"\\vortex-ho1\net mngt\OpDocs\Agreements\Official\SLD\541555-SLD.pdf"
                    'hyperLinkText = linkingDevice.buildLink(fileIn, "\\vortex-ho1\net mngt\OpDocs\Agreements\Official\SLD\")
                    hyperLinkText = linkingDevice.buildLink(fileIn, "\\vortex-ho1\net mngt\OpDocs\Agreements\Official\SLD")
                    'opening the word doc 
                    Try
                        WordDocument = wordApplication.Documents.Open(fileIn)
                    Catch ex As Exception
                        Continue For
                    End Try




                    'putting everything you need for the logger in ther arraylist
                    fileInfoHolder.Add(filename) ' put in Number code of the file 
                    fileInfoHolder.Add(fileIn) ' the actual location to the file
                    fileInfoHolder.Add(hyperLinkText) ' and the hyperlink that you are going to put into it
                    fileInfoHolder.Add(File_Path & fileIn.Substring(fileIn.LastIndexOf("\"), 11) & "-changed" & ".doc")



                    Dim truthVal As Boolean = True


                    Dim numberOfSLDs As Integer = 0

                    Dim tempp As String = ""

                    Dim textFlag As Boolean = False

                    Try
                        While wordApplication.Selection.Find.Execute(TextToFind) Or wordApplication.Selection.Find.Execute("Link to Single Line Diagram")
                            'checking if the selection has a hyperlink and if the link the linkbuilder made actually results in a file
                            'if it is less then one it has no hyper link in it
                            ' means its cannot equal 1 
                            Dim temp As String = wordApplication.Selection.Range.Text


                            If wordApplication.Selection.Range.Hyperlinks.Count <> 1 Then
                                If hyperLinkText = "There was no file by that name" Then
                                    ' there was no PDF to link to, so you log it and exist
                                    ' then close the doc 
                                    loggingDevice.recInHtml("nothingtolinkto", fileInfoHolder)
                                    'wordApplication.Documents.Close(False)
                                    SkippedFiles = SkippedFiles + 1
                                    truthVal = False
                                    Exit While
                                Else

                                    'you replace the missing hyperlink, to a PDF that exist in the system
                                    'then log it 

                                    wordApplication.ActiveDocument.Hyperlinks.Add(Anchor:=wordApplication.Selection.Range, Address:=hyperLinkText, SubAddress:="", ScreenTip:="", TextToDisplay:="Link to SLD")
                                    loggingDevice.recInHtml("nolinksomadealink", fileInfoHolder)
                                    wordApplication.ActiveDocument.SaveAs(File_Path & fileIn.Substring(fileIn.LastIndexOf("\"), 11) & "-changed" & ".doc")
                                    
                                End If



                            End If

                            numberOfSLDs = numberOfSLDs + 1

                            If numberOfSLDs = 6 Then

                                ' if you find all the text, then you can save it and close the document
                                
                                Exit While

                            End If

                            textFlag = True

                        End While

                        If Not textFlag And truthVal Then

                            loggingDevice.recInHtml("notextwasfoundinfile", fileInfoHolder)

                            textFlag = False
                            truthVal = True

                        End If

                        wordApplication.Documents.Close(False)

                    Catch exx As Exception

                        ' for when things go real bad 
                        wordApplication.Documents.Close(False)

                        loggingDevice.shutDown("htmlclose", ProcessedFiles)

                        Continue For

                    End Try


                    tempStr = ""
                    fileInfoHolder.Clear()

                    ProcessedFiles = ProcessedFiles + 1

                End If

                loggingDevice.shutDown("htmlclose", ProcessedFiles)

                List.Items.Remove(fileIn)
                FilesChanged = Directory.GetFiles(File_Path).Length - 2

            Next
            textboxr.Text = "Processed Files: " & ProcessedFiles.ToString() & "  " & "Changed Files: " & FilesChanged.ToString() & "  " & "Skipped Files: " & SkippedFiles.ToString()



        Next

        loggingDevice.shutDown("htmlclose", ProcessedFiles, FilesChanged, False)
        ProcessedFiles = 0
        FilesChanged = 0
        wordApplication.Quit()
        Return True
    End Function
End Class
