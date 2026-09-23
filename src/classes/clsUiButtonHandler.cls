VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUiButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Button As MSForms.CommandButton
Attribute Button.VB_VarHelpID = -1
Public Owner As Object
Public ActionName As String
Private Sub Button_Click()
    CallByName Owner, ActionName, VbMethod
End Sub
