Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Helper class for form Palette for event handling of textboxes.
' 2017-04-19. Gustav Brock, Cactus Data ApS, CPH.
' Version 1.0.0
' License: MIT.

' *

Private Const EventProcedure    As String = "[Event Procedure]"

Private WithEvents ClassTextBox As Access.TextBox
Attribute ClassTextBox.VB_VarHelpID = -1

Public Sub Initialize(ByRef TextBox As Access.TextBox)

    Set ClassTextBox = TextBox
    
    ClassTextBox.OnClick = EventProcedure
    
End Sub

Public Sub Terminate()

    Set ClassTextBox = Nothing

End Sub

Private Sub ClassTextBox_Click()

    ' Select full content.
    ClassTextBox.SelStart = 0
    ClassTextBox.SelLength = Len(ClassTextBox.Value)
    ' Display the clicked value.
    ClassTextBox.Parent!CopyClicked.Value = ClassTextBox.Value
    ' Copy the clicked value to the clipboard.
    DoCmd.RunCommand acCmdCopy

End Sub