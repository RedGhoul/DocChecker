Imports System
Imports System.IO
Imports Microsoft.Office.Interop.Word
Imports System.Diagnostics

Public Class Form1
    Dim di As DirectoryInfo
    Dim fi As FileInfo
    'Dim fiList As List(Of FileInfo)
    Public Folder_Path As String = ""
    Public LogFile_Path As String = ""
    Public File_Path As String = ""


    Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        'Process.Enabled = False
        DateTimePickerStart.Value = "1/1/1906"

    End Sub

    Private Sub Browse_Click(sender As Object, e As System.EventArgs) Handles Browse.Click
        FolderBrowserDialog1.ShowDialog()
        'TextBox1.Text = "C:\Projects\VBAapps\Sandbox"
        TextBox1.Text = ""
        TextBox1.Text = FolderBrowserDialog1.SelectedPath
        Folder_Path = TextBox1.Text
        If LogFile_Path.Contains("\") And Folder_Path.Contains("\") And File_Path.Contains("\") Then
            Process.Enabled = True
        End If
    End Sub

    Private Sub startapp(ByRef adress As String)

    End Sub
    ' test spec is the folder where everything will be duplicated to, if any changes are needed
    Public Sub Process_Click(sender As System.Object, e As System.EventArgs) Handles Process.Click
        
        Dim Steve As Worker = New Worker(LogFile_Path, Folder_Path, File_Path, "Link to SLD")
        ' you hand in the listbox into the work method
        If Steve.work(DateTimePickerEnd, DateTimePickerStart, ListBox1, TextBoxResults) = True Then
            If CheckOpenOnEnd.Checked = True Then
                System.Diagnostics.Process.Start(LogFile_Path & "\WordDocCheckLOG.html")
            End If

        End If

    End Sub

    Private Sub TestLogSetup_Click(sender As System.Object, e As System.EventArgs) Handles Testlogsetup.Click
        FolderBrowserDialog2.ShowDialog()
        TestlogsetupText.Text = ""
        LogFile_Path = FolderBrowserDialog2.SelectedPath
        TestlogsetupText.Text = LogFile_Path
        If LogFile_Path.Contains("\") And Folder_Path.Contains("\") And File_Path.Contains("\") Then
            Process.Enabled = True
        End If

    End Sub

    Private Sub FilesetupB_Click(sender As System.Object, e As System.EventArgs) Handles NewFilesetup.Click
        FolderBrowserDialog3.ShowDialog()
        NewFilesetupText.Text = ""
        File_Path = FolderBrowserDialog3.SelectedPath
        NewFilesetupText.Text = File_Path
        If LogFile_Path.Contains("\") And Folder_Path.Contains("\") And File_Path.Contains("\") Then
            Process.Enabled = True
        End If
    End Sub

    Private Sub Label6_Click(sender As System.Object, e As System.EventArgs) Handles Label6.Click

    End Sub

    Private Sub TextBox2_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBoxResults.TextChanged

    End Sub

    Private Sub CheckOpenOnEnd_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckOpenOnEnd.CheckedChanged

    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged
        Folder_Path = TextBox1.Text
    End Sub

    Private Sub NewFilesetupText_TextChanged(sender As System.Object, e As System.EventArgs) Handles NewFilesetupText.TextChanged
        File_Path = NewFilesetupText.Text
    End Sub

    Private Sub TestlogsetupText_TextChanged(sender As System.Object, e As System.EventArgs) Handles TestlogsetupText.TextChanged
        LogFile_Path = TestlogsetupText.Text

    End Sub
End Class
