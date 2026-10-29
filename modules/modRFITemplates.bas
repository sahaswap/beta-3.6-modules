Attribute VB_Name = "modRFITemplates"
'==================================================================
' modRFITemplates  -  RFI content accessor (centralized)
'
' RFI template content now lives ON THE SHEET (Template_DB), NOT in
' code, so analysts edit wording in one central place alongside the
' Escalation / Non-Escalation / Scenario templates.
'
' Template_DB RFI block layout (row 1 = headers):
'   F        G     H     I         J..S
'   RFI Key  To    Cc    Subject   Q1 .. Q10
' Rows 2.. hold one template each (RFI_Wise, RFI_Airwallex, RFI_Common).
' Blank question cells are simply ignored (Wise/Common use Q1..Q6).
'
' GetRFITemplate(name) reads that block and returns an RFITemplate.
' If Template_DB or the row is missing, Name = "" (caller treats as
' unknown) - so the engine fails safe rather than using stale code text.
'==================================================================
Option Explicit

' --- Shared header / footer strings (right-aligned header + encrypt tag) ---
Public Const RFI_HEADER_LINE1 As String = "Financial Crimes Compliance Unit"
Public Const RFI_HEADER_LINE2 As String = "AML / TM Department"
Public Const RFI_ENCRYPT_TAG  As String = "send cfsbencrypt"

' --- Where the RFI block lives ---
Private Const RFI_SHEET As String = "Template_DB"
Private Const RFI_KEY_COL As Long = 6      ' column F
Private Const RFI_TO_COL As Long = 7       ' column G
Private Const RFI_CC_COL As Long = 8       ' column H
Private Const RFI_SUBJ_COL As Long = 9     ' column I
Private Const RFI_Q1_COL As Long = 10      ' column J (first question)
Private Const RFI_MAX_Q As Long = 10       ' up to Q10 (column S)
Private Const RFI_FIRST_ROW As Long = 2    ' first template row
Private Const RFI_LAST_ROW As Long = 20    ' scan ceiling (plenty of headroom)

' --- One record describing a single RFI template ---
Public Type RFITemplate
    Name As String            ' "Wise" | "Airwallex" | "Common" | "" (unknown)
    ToLine As String
    CcLine As String
    SubjectFormat As String   ' with {CUSTOMER}/{CASE}/{ALERT}
    questions() As String     ' 1-based array of question text
End Type

'------------------------------------------------------------------
' GetRFITemplate
'   PRE:  templateName is "Wise" | "Airwallex" | "Common" (case-insensitive).
'   POST: reads the matching "RFI_<name>" row from Template_DB and returns a
'         populated RFITemplate (Questions 1-based, blanks dropped). Unknown
'         name or missing sheet/row -> Name = ""; never raises.
'------------------------------------------------------------------
Public Function GetRFITemplate(ByVal templateName As String) As RFITemplate
    Dim t As RFITemplate
    Dim ws As Worksheet
    Dim wantKey As String, rowKey As String
    Dim r As Long, c As Long
    Dim qlist() As String, qn As Long
    Dim cellTxt As String

    t.Name = ""
    t.questions = EmptyStrArray()

    wantKey = "RFI_" & UCase$(Trim$(templateName))   ' e.g. RFI_WISE

    On Error GoTo CleanFail
    Set ws = ThisWorkbook.Sheets(RFI_SHEET)

    For r = RFI_FIRST_ROW To RFI_LAST_ROW
        rowKey = UCase$(Trim$(CStr(ws.Cells(r, RFI_KEY_COL).Value)))
        If rowKey = wantKey Then
            ' Friendly display name (strip the RFI_ prefix, Title-ish)
            Select Case UCase$(Trim$(templateName))
                Case "WISE":      t.Name = "Wise"
                Case "AIRWALLEX": t.Name = "Airwallex"
                Case "COMMON":    t.Name = "Common"
                Case Else:        t.Name = Trim$(templateName)
            End Select

            t.ToLine = CStr(ws.Cells(r, RFI_TO_COL).Value)
            t.CcLine = CStr(ws.Cells(r, RFI_CC_COL).Value)
            t.SubjectFormat = CStr(ws.Cells(r, RFI_SUBJ_COL).Value)

            ' collect non-blank questions left-to-right
            ReDim qlist(1 To RFI_MAX_Q)
            qn = 0
            For c = RFI_Q1_COL To (RFI_Q1_COL + RFI_MAX_Q - 1)
                cellTxt = Trim$(CStr(ws.Cells(r, c).Value))
                If cellTxt <> "" Then
                    qn = qn + 1
                    qlist(qn) = cellTxt
                End If
            Next c

            If qn > 0 Then
                ReDim Preserve qlist(1 To qn)
                t.questions = qlist
            End If

            Exit For
        End If
    Next r

    GetRFITemplate = t
    Exit Function

CleanFail:
    t.Name = ""
    t.questions = EmptyStrArray()
    GetRFITemplate = t
End Function

'------------------------------------------------------------------
' EmptyStrArray - a zero-length String() (UBound < LBound) so callers
' that loop 1..UBound simply do nothing.
'------------------------------------------------------------------
Private Function EmptyStrArray() As String()
    EmptyStrArray = Split(vbNullString)
End Function
