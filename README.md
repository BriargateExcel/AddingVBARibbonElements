**Adding Your Own Ribbon Elements to Excel**

Here is some code to put in your VBA project to add your own commands to the Excel Ribbon. 

These new commands appear in the Ribbon as “Add-ins” in a section called “Custom Toolbars.”

![image-20200628094718137](https://github.com/BriargateExcel/AddingVBARibbonElements/blob/master/Add-ins%20Ribbon%20Overview.png)                               

 ![image-20200628094838659](https://github.com/BriargateExcel/AddingVBARibbonElements/blob/master/Add-ins%20Custom%20Toolbar.png)

**Instructions for using:**

1. The code below can go in any VBA Code Module or as a standalone module
2. Give your command bar a name by specifying `CommandBarName`. Spaces are allowed.
3. Specify the routine to execute in `NewButton.OnAction`
4. Specify the caption that appears when the user hovers over the button in `NewButton.Caption`. Spaces are allowed.
5. Specify your button’s appearance in `NewButton.FaceID`. There is an extensive collection of images in https://bettersolutions.com/vba/ribbon/face-ids-2003.htm

**How it works:**

1. Auto_Open executes when you launch your workbook
2. Auto_Open then
    1. Deletes `CommandBarName` if it already exists
    2. Creates `CommandBarName`
    3. Builds the button(s)
    4. Makes `CommandBarName` visible
3. When the user hovers over a button, they see “`CommandBarName: ButtonCaption`”
4. `Auto_Close` executes when you close your workbook to delete the command bar

**How I use it:**

- I have a set of buttons in my Personal.xlsb that are available in all my workbooks
- I have unique buttons for each workbook that allow me to execute routines quickly and easily

```
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
```

