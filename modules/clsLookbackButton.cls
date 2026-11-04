VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLookbackButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The click of the "Lookback Only" button that modLargeExport adds to
' frmExportMode while it runs (see AddLookbackChoice there). It answers the
' way the form's own buttons do, so the caller reads SelectedMode as usual.
' Part of the large-export tool only, like modLargeExport.

Private WithEvents m_btn As MSForms.CommandButton

Public Sub Attach(ByVal btn As MSForms.CommandButton)
    Set m_btn = btn
End Sub

Private Sub m_btn_Click()
    frmExportMode.SelectedMode = "LOOKBACK"
    frmExportMode.userCancelled = False
    frmExportMode.Hide
End Sub
