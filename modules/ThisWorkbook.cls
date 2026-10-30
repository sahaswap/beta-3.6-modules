Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' Fit Sheet1 to this screen (Module2 FitSheet1ToWindow).
Private Sub Workbook_Open()
    On Error Resume Next
    Application.Run "FitSheet1ToWindow"
End Sub
