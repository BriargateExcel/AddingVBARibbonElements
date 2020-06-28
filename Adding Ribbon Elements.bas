Option Explicit

Private Const CommandBarName As String = "Table Builder"

Private Type CodeType
  CustomBar As CommandBar
End Type

Private This As CodeType

Public Sub Auto_Open()

' https://bettersolutions.com/vba/ribbon/face-ids-2003.htm for FaceIDs

    On Error Resume Next
    CommandBars(CommandBarName).Delete
    On Error GoTo 0

    Set This.CustomBar = CommandBars.Add(Name:=CommandBarName)

    BuildButton "RourineToExecute", "Button Caption", 81 ' Capital B
    
    This.CustomBar.Visible = True

End Sub ' Auto_Open

Private Sub BuildButton( _
  ByVal RoutineToExecute As String, _
  ByVal Caption As String, _
  ByVal FaceID As Long)

' Build one button on the command bar

    Dim NewButton As CommandBarButton
    Set NewButton = This.CustomBar.Controls.Add(msoControlButton)
    
    NewButton.OnAction = RoutineToExecute
    NewButton.Caption = Caption
    NewButton.FaceID = FaceID
    
End Sub ' BuildButton

Public Sub Auto_Close()

    On Error Resume Next
    CommandBars(CommandBarName).Delete
    On Error GoTo 0

End Sub ' Auto_Close