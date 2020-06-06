Attribute VB_Name = "XML_Module"
Option Explicit
Public Sub XML_SetTag(fileName As String, tag As String)
'Adds new tag to filename.xml
'filename as filename from application.defaultfilepath
'tag as destination string (e.g. to create <xml><tag1></tag1></xml> give xml\tag1)

XML_Sheet.Cells.ClearContents

Dim file As String
Dim i As Integer, j As Integer, x As Integer
Dim loc As Integer, j_save As Integer
Dim tempVal As String

file = Application.DefaultFilePath & fileName

Call XML_Read(file)
Call XML_Itemize(tag)

x = 0
i = 1
j = 1

If XML_Sheet.Cells(1, 1) = "" Then
    XML_Sheet.Cells(1, 1) = "<" & tag & ">"
    XML_Sheet.Cells(1, 2) = "</" & tag & ">"
    x = 1
End If

Do Until x = 1
    
    If j = 1 Then
        loc = XML_FindFirst("", XML_Sheet.Cells(j, 3), i)
    ElseIf XML_Sheet.Cells(j, 3) = "" And loc <> 0 Then
        Exit Sub
    Else
        loc = XML_FindFirst(XML_Sheet.Cells(j - 1, 3), XML_Sheet.Cells(j, 3), i)
    End If
    
    If loc = 0 Then
        
        j_save = j
        
        Do Until XML_Sheet.Cells(j, 3) = ""
            XML_Sheet.Cells(i + 1, 2) = "<" & XML_Sheet.Cells(j, 3) & ">"
            i = i + 1
            j = j + 1
        Loop
        
        j = j - 1
        
        Do Until j = j_save - 1
            XML_Sheet.Cells(i + 1, 2) = "</" & XML_Sheet.Cells(j, 3) & ">"
            i = i + 1
            j = j - 1
        Loop
        
        x = 1
        
    Else
    
        i = loc
        j = j + 1
        
    End If

Loop

Call XML_Write(file)

XML_Sheet.Cells.ClearContents

End Sub
Public Sub XML_SetVal(fileName As String, tag As String, val As Variant, append As Integer)
'Adds new value to tag in filename.xml
'filename as filename from application.defaultfilepath
'tag as destination string (e.g. to create <xml><tag1></tag1></xml> give xml\tag1)
'value as tag value
'append: 1-append value, 0-replace value

XML_Sheet.Cells.ClearContents

Dim file As String
Dim i As Integer, j As Integer, x As Integer
Dim loc As Integer

file = Application.DefaultFilePath & fileName

Call XML_Read(file)
Call XML_Itemize(tag)

x = 0
i = 1
j = 1
Do Until x = 1
        
    If j = 1 Then
        loc = XML_FindFirst("", XML_Sheet.Cells(j, 3), i)
    Else
        loc = XML_FindFirst(XML_Sheet.Cells(j - 1, 3), XML_Sheet.Cells(j, 3), i)
    End If
    
    If loc = 0 Then
        If Left(XML_Sheet.Cells(i + 1, 1), 1) = "<" Or append = 1 Then
            XML_Sheet.Cells(i + 1, 2) = val
        Else
            XML_Sheet.Cells(i + 1, 1) = val
        End If
        x = 1
    Else
        i = loc
        j = j + 1
    End If

Loop

Call XML_Write(file)

XML_Sheet.Cells.ClearContents
        
End Sub
Public Function XML_GetVal(fileName As String, tag As String)
'Returns value from tag in filename.xml
'filename as filename from application.defaultfilepath
'tag as destination string (e.g. to get value from <xml><tag1>value</tag1></xml> give xml\tag1)

XML_Sheet.Cells.ClearContents

Dim file As String
Dim i As Integer, j As Integer, x As Integer
Dim loc As Integer

file = Application.DefaultFilePath & fileName

Call XML_Read(file)
Call XML_Itemize(tag)

x = 0
i = 1
j = 1
Do Until x = 1
    
    If j = 1 Then
        loc = XML_FindFirst("", XML_Sheet.Cells(j, 3), i)
    Else
        loc = XML_FindFirst(XML_Sheet.Cells(j - 1, 3), XML_Sheet.Cells(j, 3), i)
    End If
    
    If loc = 0 Then
        XML_GetVal = XML_Sheet.Cells(i + 1, 1)
        x = 1
    Else
        i = loc
        j = j + 1
    End If

Loop

If Left(XML_GetVal, 1) = "<" Then
    XML_GetVal = ""
End If

XML_Sheet.Cells.ClearContents

End Function
Public Sub XML_GetTag(fileName As String, Info As String)
'prints full XML tree from given tag (tag must be unique)
'filename as filename from application.defaultfilepath
'tag as tag name (not destination string)

XML_Sheet.Cells.ClearContents

Dim file As String, text As String
Dim i As Integer, j As Integer

file = Application.DefaultFilePath & fileName
i = 1

Open file For Input As #1
    
    Do Until EOF(1)
    
        Line Input #1, text
        
        If text = "<" & Info & ">" Then
            XML_Sheet.Cells(i, 1) = text
            Do Until text = "</" & Info & ">"
                Line Input #1, text
                i = i + 1
                XML_Sheet.Cells(i, 1) = text
            Loop
        End If
        
    Loop
    
Close #1

End Sub
Public Function XML_GetValQuick(tag As String)
'gets tag value from printed XML Tree (must call XML_GetTag first)

Dim i As Integer
i = 1

Do Until XML_Sheet.Cells(i, 1) = "<" & tag & ">" Or XML_Sheet.Cells(i, 1) = ""
    i = i + 1
Loop

If XML_Sheet.Cells(i + 1, 1) = "" Then
    XML_GetValQuick = ""
ElseIf Left(XML_Sheet.Cells(i + 1, 1), 1) <> "<" Then
    XML_GetValQuick = XML_Sheet.Cells(i + 1, 1)
End If

End Function
Public Function XML_CheckValQuick(value As String)
'checks if value exists in printed XML Tree (must call XML_GetTag first)

Dim i As Integer
i = 1

Do Until XML_Sheet.Cells(i, 1) = value Or XML_Sheet.Cells(i, 1) = ""
    i = i + 1
Loop

If XML_Sheet.Cells(i, 1) = "" Then
    XML_CheckValQuick = 0
Else
    XML_CheckValQuick = 1
End If

End Function
Public Sub XML_Clear(fileName As String)
'Clears XML

Dim file As String

file = Application.DefaultFilePath & fileName

Open file For Output As #1
    Print #1, ""
Close #1

End Sub
Private Sub XML_Read(file As String)

Dim i As Integer
Dim text As String

i = 0

Open file For Input As #1
    
    Do Until EOF(1)
        i = i + 1
        Line Input #1, text
        XML_Sheet.Cells(i, 1) = text
    Loop
    
Close #1

End Sub
Private Sub XML_Write(file As String)

Dim i As Integer, j As Integer

i = 1

Open file For Output As #1

    Do Until XML_Sheet.Cells(i, 1) = ""
    
        If XML_Sheet.Cells(i, 2) = "" Then
        
            Print #1, XML_Sheet.Cells(i, 1).value
            i = i + 1
            
        Else
            
            j = i
            Do Until XML_Sheet.Cells(j, 2) = ""
                Print #1, XML_Sheet.Cells(j, 2).value
                XML_Sheet.Cells(j, 2) = ""
                j = j + 1
            Loop
            
        End If
        
    Loop
    
Close #1
        

End Sub
Private Sub XML_Itemize(tag As String)

Dim x As Integer, i As Integer

x = 0
i = 1
 
Do Until InStr(tag, "\") = 0

    XML_Sheet.Cells(i, 3) = Left(tag, InStr(tag, "\") - 1)
    tag = Right(tag, Len(tag) - InStr(tag, "\"))
    i = i + 1
    
Loop

XML_Sheet.Cells(i, 3) = tag

End Sub
Private Function XML_FindFirst(superTag As String, tag As String, start As Integer)

Dim i As Integer

i = start

If superTag <> "" Then

    Do Until XML_Sheet.Cells(i, 1) = "</" & superTag & ">" Or XML_Sheet.Cells(i, 1) = ""
        
        If Len(XML_Sheet.Cells(i, 1)) > 1 Then
            If Mid(XML_Sheet.Cells(i, 1), 2, Len(XML_Sheet.Cells(i, 1)) - 2) = tag And Left(XML_Sheet.Cells(i, 1), 1) = "<" Then
                XML_FindFirst = i
                Exit Function
            End If
        End If
        
        i = i + 1
        
    Loop
    
    XML_FindFirst = 0
    
Else

    Do Until XML_Sheet.Cells(i, 1) = ""
        
        If Left(XML_Sheet.Cells(i, 1), 1) = "<" Then
            If Mid(XML_Sheet.Cells(i, 1), 2, Len(XML_Sheet.Cells(i, 1)) - 2) = tag Then
                XML_FindFirst = i
                Exit Function
            End If
        End If
        
        i = i + 1
        
    Loop
    
    XML_FindFirst = 0
    
End If

End Function
