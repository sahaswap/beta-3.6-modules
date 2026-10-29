Attribute VB_Name = "clsCalDay"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' ============================================================
' clsCalDay
' One instance per day cell in frmCalendar's month grid. Wraps
' the dynamically-created CommandButton so its click can be
' handled (MSForms has no control arrays, so each cell needs its
' own WithEvents sink - same trick clsRenameRow uses for rows).
' Kept alive in frmCalendar's m_cells collection.
' ============================================================

Public WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1
Public Parent As Object        ' the frmCalendar that owns this cell
Public TheDate As Date

Private Sub Btn_Click()
    If Not Parent Is Nothing Then Parent.SelectDate TheDate
End Sub

