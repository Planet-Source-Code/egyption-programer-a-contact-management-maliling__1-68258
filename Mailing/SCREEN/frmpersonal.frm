VERSION 5.00
Begin VB.Form frmpersonal 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10140
   Icon            =   "frmpersonal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   10140
   Begin VB.PictureBox Picture1 
      Height          =   585
      Left            =   -60
      Picture         =   "frmpersonal.frx":0442
      ScaleHeight     =   525
      ScaleWidth      =   10740
      TabIndex        =   30
      Top             =   -60
      Width           =   10800
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Persons"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   315
         TabIndex        =   31
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CheckBox chkpos 
      Appearance      =   0  'Flat
      Caption         =   "Show Position"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   29
      Top             =   1725
      Width           =   1425
   End
   Begin VB.ComboBox cmbname 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmpersonal.frx":3F78
      Left            =   3720
      List            =   "frmpersonal.frx":3F7A
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   28
      Top             =   1350
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.CommandButton cmdtypepos 
      Appearance      =   0  'Flat
      Caption         =   "Type of persone"
      Height          =   315
      Left            =   7852
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2610
      Width           =   1710
   End
   Begin VB.ListBox lst_person 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   7485
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1335
      Width           =   2445
   End
   Begin VB.ListBox lst_typ 
      Appearance      =   0  'Flat
      Height          =   930
      Left            =   7485
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   11
      Top             =   2970
      Width           =   2445
   End
   Begin VB.ComboBox cmb_titel 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmpersonal.frx":3F7C
      Left            =   1590
      List            =   "frmpersonal.frx":3FAA
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1335
      Width           =   1455
   End
   Begin VB.CommandButton cmdCmovefrist 
      Caption         =   "l<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Frist Record"
      Top             =   4185
      Width           =   495
   End
   Begin VB.CommandButton cmdCmoveprevious 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4635
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Previous Record"
      Top             =   4185
      Width           =   495
   End
   Begin VB.CommandButton cmdCmovenext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5145
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Next Record"
      Top             =   4185
      Width           =   495
   End
   Begin VB.CommandButton cmdCMovelast 
      Caption         =   ">l"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "Last Record"
      Top             =   4185
      Width           =   495
   End
   Begin VB.TextBox txtPersonalID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "jhjhjhkjh"
      Top             =   630
      Width           =   945
   End
   Begin VB.TextBox txt_mail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1575
      MaxLength       =   100
      TabIndex        =   7
      Top             =   2760
      Width           =   5550
   End
   Begin VB.TextBox txt_notes 
      Appearance      =   0  'Flat
      Height          =   990
      Left            =   1575
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   3075
      Width           =   5550
   End
   Begin VB.TextBox txt_fax 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1575
      MaxLength       =   24
      TabIndex        =   6
      Top             =   2400
      Width           =   1755
   End
   Begin VB.TextBox txt_pos 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1575
      MaxLength       =   60
      TabIndex        =   3
      Top             =   1695
      Width           =   4200
   End
   Begin VB.TextBox txt_mobile 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   3960
      MaxLength       =   24
      TabIndex        =   5
      Top             =   2040
      Width           =   1755
   End
   Begin VB.TextBox txt_tel 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1575
      MaxLength       =   24
      TabIndex        =   4
      Top             =   2040
      Width           =   1755
   End
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3705
      MaxLength       =   60
      TabIndex        =   2
      Top             =   1350
      Width           =   3390
   End
   Begin VB.ComboBox cmb_company 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1575
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   975
      Width           =   5550
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Personal ID:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   16
      Left            =   420
      TabIndex        =   23
      Top             =   660
      Width           =   990
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   12
      Left            =   945
      TabIndex        =   21
      Top             =   1335
      Width           =   450
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "E-mail:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   15
      TabIndex        =   20
      Top             =   2760
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Persone in company"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   7920
      TabIndex        =   19
      Top             =   990
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   15
      TabIndex        =   18
      Top             =   3105
      Width           =   1380
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Home:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   15
      TabIndex        =   17
      Top             =   2415
      Width           =   1380
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Position:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   15
      TabIndex        =   16
      Top             =   1710
      Width           =   1380
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Mobile"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   2040
      Width           =   585
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "D.Tel.:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   15
      TabIndex        =   14
      Top             =   2085
      Width           =   1380
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3105
      TabIndex        =   13
      Top             =   1395
      Width           =   525
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Company name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   15
      TabIndex        =   12
      Top             =   990
      Width           =   1380
   End
End
Attribute VB_Name = "frmpersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim ws As Workspace
Dim rs As Recordset
Dim rs2 As Recordset
Dim rs1 As Recordset
Dim SqLst As String
Dim SqLst1 As String
Dim PathPro As String
Dim X, d As String
Private Sub cmb_company_Change()
If frmpersonal.cmb_company.Text <> "" Then
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
SqLst = "Select name From Qpersonal"
SqLst = SqLst & " WHERE Companyname = '" & frmpersonal.cmb_company.Text & "'"
frmpersonal.lst_person.Clear
Set rs = db.OpenRecordset(SqLst)

Do While Not rs.EOF
frmpersonal.lst_person.AddItem rs("name")
rs.MoveNext
    Loop
    End If
End Sub

Private Sub cmb_company_Click()
If frmpersonal.cmb_company.Text <> "" Then
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
SqLst = "Select name From Qpersonal"
SqLst = SqLst & " WHERE Companyname = '" & frmpersonal.cmb_company.Text & "'"
frmpersonal.lst_person.Clear
Set rs = db.OpenRecordset(SqLst)

Do While Not rs.EOF
frmpersonal.lst_person.AddItem rs("name")
rs.MoveNext
    Loop
    End If
    frmpersonal.lst_typ.Clear
End Sub

Private Sub cmbname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frmpersonal.search
frmpersonal.cmdCmovenext.Enabled = True
frmpersonal.cmdCmoveprevious.Enabled = True
frmpersonal.txt_name.Visible = True
'frmpersonal.txtAbbreviation.Visible = True
frmpersonal.cmbname.Visible = False
'frmpersonal.cmbabbr.Visible = False
frmpersonal.txtPersonalID.Enabled = False
frmpersonal.txtPersonalID.Locked = True
End If
End Sub
Public Sub PersonalDis()
frmpersonal.txtPersonalID.Text = rs("personalId")

frmpersonal.txt_tel.Text = rs("Tel1") & ""
frmpersonal.txt_fax.Text = rs("fax1") & ""
frmpersonal.cmb_titel.Text = rs("Title") & ""
frmpersonal.txt_pos.Text = rs("Position") & ""
If IsNull(rs("showPos")) Or rs("showPos") = "No" Then
frmpersonal.chkpos.Value = 0
Else
frmpersonal.chkpos.Value = 1
End If
frmpersonal.txt_notes.Text = rs("Notes") & ""
frmpersonal.txt_name.Text = rs("name") & ""
frmpersonal.txt_mail.Text = rs("Email") & ""
frmpersonal.txt_mobile.Text = rs("Tel2") & ""
frmpersonal.cmb_company.Text = rs("Companyname") & ""
y = 0
SqLst = "Select posTypename From Qper"
SqLst = SqLst & " WHERE personalId = " & frmpersonal.txtPersonalID.Text & ""
frmpersonal.lst_typ.Clear
Set rs = db.OpenRecordset(SqLst)
Do While Not rs.EOF
frmpersonal.lst_typ.AddItem rs("posTypename")
frmpersonal.lst_typ.Selected(y) = True
y = y + 1
rs.MoveNext
Loop
SqLst = "Select name From Qpersonal"
SqLst = SqLst & " WHERE Companyname = '" & frmpersonal.cmb_company.Text & "'"
frmpersonal.lst_person.Clear
Set rs = db.OpenRecordset(SqLst)
Do While Not rs.EOF
frmpersonal.lst_person.AddItem rs("name")
rs.MoveNext
Loop
'frmpersonal.txt_name.SetFocus
End Sub

Private Sub cmdcmovefrist_Click()
Moving.Movefrist (Me.Name)
End Sub

Private Sub cmdcMovelast_Click()
Moving.MoveLast (Me.Name)
End Sub

Private Sub cmdcmovenext_Click()
Moving.NextM (Me.Name)
End Sub

Private Sub cmdcmoveprevious_Click()
Moving.Previous (Me.Name)
End Sub

Private Sub cmdtypepos_Click()
If frmpersonal.txtPersonalID.Text = "" Then Exit Sub
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")

     frmpersonal.lst_typ.Clear
    Set rs = db.OpenRecordset("PosType")
     Do While Not rs.EOF
       frmpersonal.lst_typ.AddItem rs("posTypename")
        rs.MoveNext
    Loop
SqLst = "Select posTypename From Qper"
SqLst = SqLst & " WHERE personalId = " & frmpersonal.txtPersonalID.Text & ""

Set rs = db.OpenRecordset(SqLst)

Do While Not rs.EOF
y = 0
Do While y < frmpersonal.lst_typ.ListCount
frmpersonal.lst_typ.ListIndex = y

If rs("posTypename") = frmpersonal.lst_typ.List(frmpersonal.lst_typ.ListIndex) Then
frmpersonal.lst_typ.Selected(y) = True
GoTo e2
Else
End If
y = y + 1

Loop
e2:
rs.MoveNext
Loop
End Sub

Public Sub search()
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
 SqLst = "Select * From Qpersonal"
        SqLst = SqLst & " WHERE name = '" & frmpersonal.cmbname.Text & "'"
        'SqLst = SqLst & " and Companyname = '" & frmCampany.txtCompanyName.Text & "' ORDER BY CompanyId "
        Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
frmpersonal.PersonalDis
Else
MsgBox ("This name is not Found")
Moving.Movefrist (Me.Name)
End If
End Sub


Private Sub Form_Activate()
Add.Mode
frmMain.Arrange Index
If frmpersonal.txtPersonalID.Text = "" Then
display.display (Me.Name)
End If
End Sub
Public Sub searchcode()
On Error GoTo RR
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
 SqLst = "Select * From Qpersonal"
        SqLst = SqLst & " WHERE personalId = " & frmpersonal.txtPersonalID.Text & ""
        'SqLst = SqLst & " and Companyname = '" & frmCampany.txtCompanyName.Text & "' ORDER BY CompanyId "
        Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
frmpersonal.PersonalDis
Else
MsgBox ("Can't be found")
Moving.Movefrist (Me.Name)
End If
RR:
If Err.Number = 3075 Then
MsgBox "Enter any value to comblet search"
'Add.addsearch (ACFRM.Name)
Moving.Movefrist (Me.Name)
End If
End Sub

Private Sub Form_Initialize()
If Right(App.Path, 1) = "\" Then
            PathPro = App.Path & "Mailing.mdb"
        Else
            PathPro = App.Path & "\Mailing.mdb"
        End If
Set db = OpenDatabase(PathPro, True, False, ";pwd=eit")
End Sub

Private Sub Form_Load()
frmMain.hideb
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
frmMain.hideb

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Add.ModeP

End Sub

Private Sub lst_person_Click()
If frmpersonal.lst_person.Text <> "" Then
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
 SqLst = "Select * From Qpersonal"
        SqLst = SqLst & " WHERE name = '" & frmpersonal.lst_person.Text & "'"
        'SqLst = SqLst & " and Companyname = '" & frmCampany.txtCompanyName.Text & "' ORDER BY CompanyId "
        Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
frmpersonal.txtPersonalID.Text = rs("personalId")

frmpersonal.txt_name.SetFocus
frmpersonal.txt_tel.Text = rs("Tel1") & ""
frmpersonal.txt_fax.Text = rs("fax1") & ""
frmpersonal.cmb_titel.Text = rs("Title") & ""
frmpersonal.txt_pos.Text = rs("Position") & ""
frmpersonal.txt_notes.Text = rs("Notes") & ""
frmpersonal.txt_name.Text = rs("name") & ""
If IsNull(rs("showPos")) Or rs("showPos") = "No" Then
frmpersonal.chkpos.Value = 0
Else
frmpersonal.chkpos.Value = 1
End If
frmpersonal.txt_mail.Text = rs("Email") & ""
frmpersonal.txt_mobile.Text = rs("Tel2") & ""
frmpersonal.cmb_company.Text = rs("Companyname") & ""
y = 0
SqLst = "Select posTypename From Qper"
SqLst = SqLst & " WHERE personalId = " & frmpersonal.txtPersonalID.Text & ""
frmpersonal.lst_typ.Clear
Set rs = db.OpenRecordset(SqLst)
Do While Not rs.EOF
frmpersonal.lst_typ.AddItem rs("posTypename")
frmpersonal.lst_typ.Selected(y) = True
y = y + 1
rs.MoveNext
Loop
SqLst = "Select name From Qpersonal"
SqLst = SqLst & " WHERE Companyname = '" & frmpersonal.cmb_company.Text & "'"
frmpersonal.lst_person.Clear
Set rs = db.OpenRecordset(SqLst)
Do While Not rs.EOF
frmpersonal.lst_person.AddItem rs("name")
rs.MoveNext
Loop
Else
End If
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
frmMain.hideb
End Sub

Private Sub txt_fax_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 47 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_mobile_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 47 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_tel_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 47 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtPersonalID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frmpersonal.searchcode
frmpersonal.cmdCmovenext.Enabled = True
frmpersonal.cmdCmoveprevious.Enabled = True
frmpersonal.txt_name.Visible = True
'frmpersonal.txtAbbreviation.Visible = True
frmpersonal.cmbname.Visible = False
'frmpersonal.cmbabbr.Visible = False
frmpersonal.txtPersonalID.Enabled = False
frmpersonal.txtPersonalID.Locked = True

End If
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 47 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub

