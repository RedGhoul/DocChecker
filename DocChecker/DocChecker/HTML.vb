'Passes a really big string in and out of it's self
Public Class HTML
    Dim titleOfDoc As String = ""
    Dim titleSubHeading1 As String = ""
    Dim titleSubHeading2 As String = ""
    Dim Document As String = "" '<![CDATA[<html></html>]]>.Value
    'still editing the sytles
    Dim htmlStyle As String = <![CDATA[<style> 
   	.chart div {
	  font: 18px sans-serif;
	  background-color: #576565;
	  text-align: right;
	  
	  padding: 5px;
	  margin: 1px;
	  color: white;
	  font-family: 'Arial';
	   
	}
	h2
    {
       font-family: 'Arial';
       color: #484E4E;
    }
    h4
    {
       font-family: 'Arial';
       color: #484E4E;
    }
 
table {
  font-family: 'Arial';
  
  
  border: 1px solid #eee;
  border-bottom: 2px solid #00cccc;
  font-size: 14px;
    
    }
 
  th, td 
  {
    color: #484E4E;
    padding: 1px 4px;
    border-collapse: collapse;
  }
  
   th 
   {
    background: #00cccc;
    color: #fff;
    text-transform: uppercase;
        font-size: 12px;
    }
 
    table tr:nth-child(odd)
    {
        background-color: #eee;
    }
    table tr:nth-child(even)
    {
        background-color: #fff;
    }
  
    input 
    {
    background-color: #00CCCC; 
    border: none;
    color: white;
    font-size: 16px;
    }

}</style> 
]]>.Value
    Dim htmlVBS As String = <![CDATA[<script language="vbscript" type="text/vbscript">
    Function Replace(File)
	dim a 
	a = MsgBox("Do you Really want to Replace the File?",1,"Warning")
	    If a = 1 then
		    dim filesys 
		    dim str
		    str = Split(File, "+") 
		    set filesys=CreateObject("Scripting.FileSystemObject")
            filesys.CopyFile str(1) , str(2)
		    filesys.DeleteFile str(1) 
		    filesys.CopyFile str(0), str(1)
	    End If
    End Function

</script>]]>.Value
    Dim htmlBody As String = ""
    Dim htmlTable As String = ""
    Dim htmlTableFiller As String = ""

    Dim htmlraw As ArrayList = New ArrayList()
    Dim htmlrawOld As ArrayList = New ArrayList()

    Dim htmlstats As String = ""
    Dim compar As String = ""
    Dim rowsCount As Integer = 0
    Dim rowtemp As String = ""

    Sub New(ByRef Title As String, ByRef sourceFileLocation As String, ByRef destinationfolder As String)
        titleOfDoc = "<h2>" & Title & "</h2>" & vbNewLine
        titleSubHeading1 = "<h4>" & sourceFileLocation & "</h4>" & vbNewLine
        titleSubHeading2 = "<h4>" & "Destination Folder: " & destinationfolder & "</h4>" & vbNewLine
    End Sub

    ' the                1st is your schedule, 2nd old doc link, 3rd new doc link, 4th SLD that was linked to, 5th the replace button // the other stuff is the links you may or may not want to put into it
    Function makeTitleRowTable(ByRef C1h1 As String, ByRef C2h2 As String, ByRef C3h3 As String, ByRef C4h4 As String, ByRef c5h5 As String) As Boolean
        Dim templinktext As String = ""
        Dim temp As String = ""
        ' should make it check everything in the future
        temp = "<tr>" & "<th>" & "<a>" & C1h1 & "</a>" & "</th>" & "<th>" & "<a>" & C2h2 & "</a>" & "</th>" & "<th>" & "<a>" & C3h3 & "</a>" & "</th>" & "<th>" & "<a>" & C4h4 & "</a>" & "</th>" & "<th>" & "<a>" & c5h5 & "</a>" & "</th>" & "</tr>" & vbNewLine
        htmlraw.Add(temp)
        Return True
    End Function



    ' the                 1st Name of sech    , 2nd nothing here                      , 3rd The new doc link              ,4th link to new document               ,5th Name of the error or SLD       , 6th the link to the SLD file          , 7th just the button or something
    Function makeRowTable(ByRef C1h1 As String, Optional ByRef C1h1Link As String = "", Optional ByRef C2h2 As String = "", Optional ByRef C2h2Link As String = "", Optional ByRef C3h3 As String = "", Optional ByRef C3h3Link As String = "", Optional ByRef C4h4 As String = "") As Boolean
        Dim templinktext As String = ""
        Dim templinktext2 As String = ""
        Dim tmeplinktext3 As String = ""
        Dim temp As String = ""
        ' name of new file and then name of the old file

        Dim buttons As String = "<input id=" & """AB""" & " " & "type=" & """button""" & " " & "value=" & """Replace""" & " " & "name=" & """run_button""" & " " & "onClick=" & """Replace(" & "'" & C4h4 & "+" & C2h2Link & "+" & C2h2Link.Substring(0, C2h2Link.LastIndexOf(".doc")).Replace("WORKING", "Archive") & "-BackUp.doc" & "'" & ")""" & " > "
        ' should make it check everytding in tde future
        ' you have two links you wanna make 

        If C2h2Link <> "" And C3h3Link <> "" Then
            templinktext = "<a href=""" & C2h2Link & """>" & C2h2 & "</a>"
            templinktext2 = "<a href=""" & C3h3Link & """>" & C3h3 & "</a>" 'link to the SLD 
            tmeplinktext3 = "<a href=""" & C4h4 & """>" & C2h2 & "</a>"
            '"<td>" & tmeplinktext3 & "</td>" & 
            If C4h4 <> "" Then
                temp = "<tr>" & "<td>" & C1h1 & "</td>" & "<td>" & templinktext & "</td>" & "<td>" & tmeplinktext3 & "</td>" & "<td>" & templinktext2 & "</td>" & "<td>" & buttons & "</td>" & "</tr>" & vbNewLine
            Else
                temp = "<tr>" & "<td>" & C1h1 & "</td>" & "<td>" & templinktext & "</td>" & "<td>" & "NULL" & "</td>" & "<td>" & templinktext2 & "</td>" & "<td>" & "NULL" & "</td>" & "</tr>" & vbNewLine
            End If
        ElseIf C3h3Link <> "" Then
            ' if you have only a single link to make 
            templinktext2 = "<a href=""" & C3h3Link & """>" & C3h3 & "</a>" & vbNewLine
            temp = "<tr>" & "<td>" & C1h1 & "</td>" & "<td>" & C3h3 & "</td>" & "<td>" & templinktext2 & "</td>" & "<td>" & C4h4 & "</td>" & "</tr>" & vbNewLine
        ElseIf C2h2Link <> "" Then
            ' if you have only a single link to make 
            templinktext = "<a href=""" & C2h2Link & """>" & C2h2 & "</a>"
            temp = "<tr>" & "<td>" & C1h1 & "</td>" & "<td>" & templinktext & "</td>" & "<td>" & C3h3 & "</td>" & "<td>" & C4h4 & "</td>" & "</tr>" & vbNewLine
        Else
            ' if you have no links to make 
            temp = "<tr>" & "<td>" & C1h1 & "</td>" & "<td>" & C2h2 & "</td>" & "<td>" & C3h3 & "</td>" & "<td>" & C4h4 & "</td>" & "</tr>" & vbNewLine
        End If

            htmlraw.Add(temp)
            Return True
    End Function

    ' pos 0 = the total number of files that it went through
    ' pos 1 = name of bar 1, number assocated with bar 1
    ' and it continues with that pattern, other than the 
    Function makeBarChart(ByRef chartPar As ArrayList) As String
        Dim table As String = ""

        Dim numOfFiles As Integer = Convert.ToInt32(chartPar(0))
        'should put some error checking here in the future

        For index = 1 To chartPar.Count - 1
            Dim stuff As String() = chartPar(index).ToString().Split(",")
            Dim width As String = (((stuff(1)) / numOfFiles) * 1050).ToString()
            If chartPar(index).ToString().Contains("TOTAL") Then
                ' magic number here just for now
                table = table + "<div style=" & "width:" & 900 & "px;" & ">" & stuff(0) & ":" & numOfFiles & "</div>" & vbNewLine
            Else
                table = table + "<div style=" & "width:" & width & "px;" & ">" & stuff(0) & ":" & stuff(1) & "</div>" & vbNewLine
            End If

        Next

        Dim tempTable As String = "<table>" & "<tr>" & "<div class=" & """chart""" & ">" & table & "</div>" & "</tr>" & "</table>" & vbNewLine
        Return tempTable
    End Function

    ' the way that the Data arraylist is defined is by, a key value pair ecoded
    ' in a single string
    'POS 1 = "a data identifer, then a number' this was done so that this is the same type data format they 
    'would have to use with the rest

    Function stats(ByRef data As ArrayList) As String
        Dim temp As String = "<h4>STATS</h4>" & vbNewLine
        'Dim holder As String = "<h3>stats</h3>"
        For index = 0 To data.Count - 1
            Dim stuff As String() = data(index).ToString().Split(",")
            temp = temp + "<h4>" & stuff(0) & stuff(1) & "</h4>" & vbNewLine
        Next

        Return temp
    End Function

    Function makeRows(ByRef value As Integer) As String

        Dim reutrnRow As String = ""

        If htmlraw.Count <> 0 Then
            For Each rows In htmlraw
                If Not rowtemp.Contains(rows.ToString()) Then
                    reutrnRow = reutrnRow + rows.ToString()
                End If

            Next
        End If

        rowtemp = reutrnRow
        Return reutrnRow

    End Function

    Function makeDoc(ByRef subject As String, Optional ByRef t As Boolean = True) As String
        'If htmlraw.Count <> 0 Then
        '    For Each row In htmlraw
        '        htmlTableFiller = htmlTableFiller + row.ToString()
        '    Next
        'Else
        '    htmlTableFiller = "<p> there was no table made here </p>"
        'End If

        Dim reutrnRow As String = ""

        If htmlraw.Count <> 0 Then
            If subject.Contains("0") Then

                For Each rows In htmlraw
                    reutrnRow = reutrnRow + rows.ToString()
                Next

            Else

                For Each rows In htmlraw
                    If Not rowtemp.Contains(rows.ToString()) Then
                        reutrnRow = reutrnRow + rows.ToString()
                    End If

                Next
            End If

        End If

        If t Then
            rowtemp = reutrnRow
            htmlTableFiller = reutrnRow
        End If



        If subject = "" Then
            'htmlTable = "<table>" & "<tr>" & "<th>" & titleOfDoc & titleSubHeading & "</th>" & "<th></th>" & "<th></th>" & "<th></th>" & "<tr>" & htmlTableFiller & "</table>"
            htmlTable = "<table>" & htmlTableFiller & "</table>" & vbNewLine
            htmlBody = "<body>" & htmlTable & "</body>" & vbNewLine
            Document = "<html>" & htmlStyle & htmlBody & "</html>" '<![CDATA[<html></html>]]>.Value& vbNewLine

        ElseIf subject.Contains("table") Then
            ' here do to the error checking I am put in at a later time
            htmlTable = "<table>" & htmlTableFiller & "</table>" & subject & vbNewLine
            htmlBody = "<body>" & titleOfDoc & titleSubHeading1 & titleSubHeading2 & htmlTable & "</body>" & vbNewLine
            Document = "<html>" & htmlStyle & htmlBody & "</html>" & vbNewLine '<![CDATA[<html></html>]]>.Value \

        ElseIf subject.Contains("stats".ToUpper()) Then
            htmlTable = "<table>" & htmlTableFiller & "</table>" & vbNewLine '& htmlstats
            htmlBody = "<body>" & titleOfDoc & titleSubHeading1 & titleSubHeading2 & htmlTable & subject & "</body>" & vbNewLine
            Document = "<html>" & htmlVBS & htmlStyle & htmlBody & "</html>" & vbNewLine '<![CDATA[<html></html>]]>.Value
            'subject = ""
        End If

        Return Document
    End Function

End Class