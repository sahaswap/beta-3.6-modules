Attribute VB_Name = "frmCalendar"
Attribute VB_Base = "0{DC314EDA-04D2-4C48-851C-ECBF589E000A}{2DE3387B-5596-4E5E-9E56-F33ACC826C4B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Lightweight month calendar. Shown modally via PickDate; returns
' the chosen date. Day cells are built at runtime inside fraDays,
' each wrapped by a clsCalDay that reports its click back here. No
' external ActiveX/OCX controls, so it works on locked-down machines.

Private m_view As Date       ' first day of the month currently shown
Private m_picked As Boolean
Private m_result As Date
Private m_cells As Collection   ' clsCalDay instances (kept alive)

' Show modally seeded from seedText (mm.dd.yyyy, else today).
' Returns True + outDate if a day was picked, False if cancelled.
Public Function PickDate(ByVal seedText As String, ByRef outDate As Date) As Boolean
Dim seed As Date
seed = SeedFromText(seedText)
m_view = DateSerial(Year(seed), Month(seed), 1)
m_picked = False
BuildGrid
Me.Show 1
PickDate = m_picked
If m_picked Then outDate = m_result
End Function

Private Function SeedFromText(ByVal s As String) As Date
Dim parts() As String
parts = Split(Trim(s), ".")
If UBound(parts) = 2 Then
If IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2)) Then
On Error Resume Next
SeedFromText = DateSerial(CInt(parts(2)), CInt(parts(0)), CInt(parts(1)))
If Err.Number = 0 Then Exit Function
On Error GoTo 0
End If
End If
SeedFromText = Date
End Function

Private Sub BuildGrid()
lblMonth.caption = Format$(m_view, "mmmm yyyy")
' Drop the old cell wrappers first so their WithEvents refs are
' released before we remove the buttons they point at.
Set m_cells = New Collection
Dim i As Long
For i = fraDays.Controls.count - 1 To 0 Step -1
fraDays.Controls.Remove fraDays.Controls(i).Name
Next i

Dim firstDow As Long, daysIn As Long, d As Long
firstDow = Weekday(m_view, vbSunday) - 1
daysIn = Day(DateSerial(Year(m_view), Month(m_view) + 1, 0))

Const CW As Single = 30
Const ch As Single = 18
For d = 1 To daysIn
Dim idx As Long, col As Long, row As Long
idx = firstDow + (d - 1)
col = idx Mod 7
row = idx \ 7
Dim Btn As MSForms.CommandButton
Set Btn = fraDays.Controls.Add("Forms.CommandButton.1")
Btn.Left = col * CW
Btn.Top = row * ch
Btn.Width = CW
Btn.Height = ch
Btn.caption = CStr(d)
Dim cd As clsCalDay
Set cd = New clsCalDay
Set cd.Btn = Btn
Set cd.Parent = Me
cd.TheDate = DateSerial(Year(m_view), Month(m_view), d)
m_cells.Add cd
Next d
End Sub

Private Sub btnPrev_Click()
m_view = DateSerial(Year(m_view), Month(m_view) - 1, 1)
BuildGrid
End Sub

Private Sub btnNext_Click()
m_view = DateSerial(Year(m_view), Month(m_view) + 1, 1)
BuildGrid
End Sub

' Called by a clsCalDay when its day button is clicked.
Public Sub SelectDate(ByVal d As Date)
m_result = d
m_picked = True
Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then       ' the X - cancel but keep the instance
m_picked = False
Cancel = True
Me.Hide
End If
End Sub

