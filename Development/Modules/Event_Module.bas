Attribute VB_Name = "Event_Module"
Option Explicit
Public Function KnowCheck(parameter As String)

KnowCheck = WorksheetFunction.RandBetween(1, 100)

End Function
Public Function PersCheck(parameter As String)

PersCheck = WorksheetFunction.RandBetween(1, 100)

End Function
Public Function PartCheck(parameter As String)

PartCheck = 1

End Function
Public Function ref(scope As String, parameter As String)

If scope = "player" Then
    ref = XML_GetVal("\Resources\Player.xml", parameter)
ElseIf scope = "planet" Then
    ref = Get_Info(parameter)
End If

End Function
Public Sub do_action(action As String, parameters As String)
'Things do can do:
'   Move player
'   Move other characters
'   Do damage to player
'   Do damage to other characters
'   Repair player
'   Repair other characters
'   Gain/Lose Personality Stats
'   Gain/Lose Knowledge
'   Gain/Lose Parts
'   Gain/Lose Opinion

End Sub
Public Sub Import_Event(name As String)

Dim cmpComponent As VBIDE.VBComponent

For Each cmpComponent In ThisWorkbook.VBProject.VBComponents
    If Left(cmpComponent.name, 6) = "Module" Then
        ThisWorkbook.VBProject.VBComponents.Remove cmpComponent
    End If
Next

ThisWorkbook.VBProject.VBComponents.Import (Application.DefaultFilePath & "\Resources\Events\" & name & ".bas")

End Sub

