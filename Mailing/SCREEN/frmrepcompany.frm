VERSION 5.00
Begin VB.Form frmReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Information for Companies"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Send Mail"
      Height          =   375
      Left            =   2145
      TabIndex        =   23
      Top             =   5010
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3885
      TabIndex        =   22
      Top             =   5010
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   375
      Left            =   3008
      Picture         =   "frmrepcompany.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5010
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search By Words"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   390
      TabIndex        =   15
      Top             =   3810
      Width           =   6210
      Begin VB.OptionButton Option1 
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   5190
         TabIndex        =   20
         Top             =   825
         Width           =   945
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Product "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3660
         TabIndex        =   19
         Top             =   840
         Width           =   990
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Company name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   18
         Top             =   825
         Width           =   1785
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   17
         Top             =   810
         Width           =   1080
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   1230
         TabIndex        =   16
         Top             =   300
         Width           =   3540
      End
   End
   Begin VB.ComboBox Cmbtype 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmrepcompany.frx":0102
      Left            =   2100
      List            =   "frmrepcompany.frx":0104
      TabIndex        =   14
      Top             =   1515
      Width           =   3225
   End
   Begin VB.ComboBox CmbProduct 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmrepcompany.frx":0106
      Left            =   2115
      List            =   "frmrepcompany.frx":0108
      TabIndex        =   13
      Top             =   720
      Width           =   3225
   End
   Begin VB.ComboBox CmbCategoryname 
      Height          =   315
      Left            =   2085
      TabIndex        =   12
      Top             =   360
      Width           =   3225
   End
   Begin VB.Frame Frame2 
      Height          =   1725
      Left            =   630
      TabIndex        =   4
      Top             =   2025
      Width           =   4920
      Begin VB.ComboBox CmbRegion 
         DataField       =   "City"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmrepcompany.frx":010A
         Left            =   1380
         List            =   "frmrepcompany.frx":010C
         TabIndex        =   11
         Top             =   1170
         Width           =   3300
      End
      Begin VB.ComboBox Cmbcountry 
         Height          =   315
         Left            =   1395
         TabIndex        =   10
         Top             =   405
         Width           =   3315
      End
      Begin VB.ComboBox cmbCity 
         DataField       =   "City"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmrepcompany.frx":010E
         Left            =   1395
         List            =   "frmrepcompany.frx":0110
         TabIndex        =   1
         Top             =   780
         Width           =   3315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Country:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   585
         TabIndex        =   8
         Top             =   390
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C7E6FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Region:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   19
         Left            =   690
         TabIndex        =   7
         Top             =   1185
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "City:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   885
         TabIndex        =   5
         Top             =   795
         Width           =   405
      End
   End
   Begin VB.ComboBox cmbCompanyname 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmrepcompany.frx":0112
      Left            =   2100
      List            =   "frmrepcompany.frx":0114
      TabIndex        =   0
      Top             =   1125
      Width           =   3225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "Company type:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   765
      TabIndex        =   9
      Top             =   1575
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "Product name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   14
      Left            =   810
      TabIndex        =   6
      Top             =   765
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "Company name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   660
      TabIndex        =   3
      Top             =   1155
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "Category name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   705
      TabIndex        =   2
      Top             =   315
      Width           =   1260
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCategoryname_Click()
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(App.Path & "\Mailing.mdb")
SqLst1 = "Select Companyname From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
NewMsg.cmbCompanyname.Clear
NewMsg.cmbCompanyname.AddItem "All"
    Do While Not rs1.EOF
       NewMsg.cmbCompanyname.AddItem rs1("Companyname")
       rs1.MoveNext
       Loop

SqLst1 = "Select Productname From QProduct"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
NewMsg.CmbProduct.Clear
NewMsg.CmbProduct.AddItem "All"
    Do While Not rs1.EOF
       NewMsg.CmbProduct.AddItem rs1("Productname")
       rs1.MoveNext
       Loop
NewMsg.Cmbtype.Text = ""
End Sub



Private Sub cmbCity_Click()
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(App.Path & "\Mailing.mdb")
SqLst1 = "Select Region From City"
SqLst1 = SqLst1 & " WHERE city = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
NewMsg.CmbRegion.Clear
 'Set rs = db.OpenRecordset("Category")
    Do While Not rs1.EOF
        NewMsg.CmbRegion.AddItem rs1("Region") & ""
       rs1.MoveNext
    Loop
End Sub

Private Sub cmbCompanyname_Click()
NewMsg.Cmbtype.Text = "All"
NewMsg.CmbCategoryname.Text = "All"
NewMsg.CmbProduct.Text = "All"
End Sub

Private Sub Cmbcountry_Click()
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(App.Path & "\Mailing.mdb")
SqLst1 = "Select city From City"
SqLst1 = SqLst1 & " WHERE country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
NewMsg.cmbCity.Clear
 'Set rs = db.OpenRecordset("Category")
    Do While Not rs1.EOF
        NewMsg.cmbCity.AddItem rs1("city")
       rs1.MoveNext
    Loop
    NewMsg.CmbRegion.Clear
End Sub

Private Sub CmbProduct_Click()
NewMsg.Cmbtype.Text = ""
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(App.Path & "\Mailing.mdb")
SqLst1 = "Select Companyname From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
NewMsg.cmbCompanyname.Clear
NewMsg.cmbCompanyname.AddItem "All"
    Do While Not rs1.EOF
       NewMsg.cmbCompanyname.AddItem rs1("Companyname")
       rs1.MoveNext
       Loop
End Sub

Private Sub Cmbtype_Click()
NewMsg.CmbCategoryname.Text = ""
NewMsg.CmbProduct.Text = ""
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(App.Path & "\Mailing.mdb")
SqLst1 = "Select Companyname From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
NewMsg.cmbCompanyname.Clear
NewMsg.cmbCompanyname.AddItem "All"
    Do While Not rs1.EOF
       NewMsg.cmbCompanyname.AddItem rs1("Companyname")
       rs1.MoveNext
       Loop

End Sub

Private Sub cmdPrint_Click()
If NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\repcompany.rpt")
    frmMain.CrystalReport1.SelectionFormula = ""
    frmMain.CrystalReport1.Destination = 0
    frmMain.CrystalReport1.WindowState = crptMaximized
    frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
'===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\repcompany.rpt")
    frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Categoryname}=" & _
    "'" & NewMsg.CmbCategoryname.Text & "'"
    frmMain.CrystalReport1.Destination = 0
     frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
'============================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
    frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Categoryname}=" & _
    "'" & NewMsg.CmbCategoryname.Text & "'"
    frmMain.CrystalReport1.Destination = 0
    
    frmMain.CrystalReport1.WindowState = crptMaximized
    frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
   Exit Sub
     '=============================================
ElseIf NewMsg.CmbCategoryname.Text = "All" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\repcompany.rpt")
    frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Companyname}=" & _
    "'" & NewMsg.cmbCompanyname.Text & "'"
    frmMain.CrystalReport1.Destination = 0
    frmMain.CrystalReport1.WindowState = crptMaximized
    frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
    '============================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
    frmMain.CrystalReport1.SelectionFormula = ""
    frmMain.CrystalReport1.Destination = 0
    
    frmMain.CrystalReport1.WindowState = crptMaximized
    frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
   Exit Sub
   
'===========================================
ElseIf NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
    frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Productname}=" & _
    "'" & NewMsg.CmbProduct.Text & "'"
    frmMain.CrystalReport1.Destination = 0
     frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
    '=============================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text <> "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
    frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Companyname}=" & _
    "'" & NewMsg.cmbCompanyname.Text & "'"
    frmMain.CrystalReport1.Destination = 0
    frmMain.CrystalReport1.WindowState = crptMaximized
    frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
   '============================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "All" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\type.rpt")
    frmMain.CrystalReport1.SelectionFormula = ""
    frmMain.CrystalReport1.Destination = 0
    
    frmMain.CrystalReport1.WindowState = crptMaximized
    frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
    Exit Sub
    '============================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbtype.Text <> "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\type.rpt")
    frmMain.CrystalReport1.SelectionFormula = "{Qtype.Typename}=" & _
    "'" & NewMsg.Cmbtype.Text & "'"
    frmMain.CrystalReport1.Destination = 0
    frmMain.CrystalReport1.WindowState = crptMaximized
    frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
'=============================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbtype.Text = "All" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\type.rpt")
    frmMain.CrystalReport1.SelectionFormula = "{Qtype.Companyname}=" & _
    "'" & NewMsg.cmbCompanyname.Text & "'"
    frmMain.CrystalReport1.Destination = 0
    frmMain.CrystalReport1.WindowState = crptMaximized
    frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
    '====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\repcompany.rpt")
     frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Categoryname}=" & _
    "'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatcomp.country}=" & _
    "'" & NewMsg.Cmbcountry.Text & "'"
    frmMain.CrystalReport1.Destination = 0
     frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
    '====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
    And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\repcompany.rpt")
     frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Categoryname}=" & _
    "'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatcomp.city}=" & _
    "'" & NewMsg.cmbCity.Text & "'"
    frmMain.CrystalReport1.Destination = 0
     frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
''====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
    NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\repcompany.rpt")
     frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Categoryname}=" & _
    "'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatcomp.Region}=" & _
    "'" & NewMsg.CmbRegion.Text & "'"
    frmMain.CrystalReport1.Destination = 0
     frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
'====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
     frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Productname}=" & _
    "'" & NewMsg.CmbProduct.Text & "'" & " And {Qprocomp.country}=" & _
    "'" & NewMsg.Cmbcountry.Text & "'"
    frmMain.CrystalReport1.Destination = 0
     frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
    '====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
    And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
     frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Productname}=" & _
    "'" & NewMsg.CmbProduct.Text & "'" & " And {Qprocomp.city}=" & _
    "'" & NewMsg.cmbCity.Text & "'"
    frmMain.CrystalReport1.Destination = 0
     frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
''====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
    NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
     frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Productname}=" & _
    "'" & NewMsg.CmbProduct.Text & "'" & " And {Qprocomp.Region}=" & _
    "'" & NewMsg.CmbRegion.Text & "'"
    frmMain.CrystalReport1.Destination = 0
     frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
'====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\type.rpt")
     frmMain.CrystalReport1.SelectionFormula = "{Qtype.Typename}=" & _
    "'" & NewMsg.Cmbtype.Text & "'" & " And {Qtype.country}=" & _
    "'" & NewMsg.Cmbcountry.Text & "'"
    frmMain.CrystalReport1.Destination = 0
     frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
    '====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
    And NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text = "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\type.rpt")
     frmMain.CrystalReport1.SelectionFormula = "{Qtype.Typename}=" & _
    "'" & NewMsg.Cmbtype.Text & "'" & " And {Qtype.city}=" & _
    "'" & NewMsg.cmbCity.Text & "'"
    frmMain.CrystalReport1.Destination = 0
     frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
''====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
    NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text <> "" Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\type.rpt")
     frmMain.CrystalReport1.SelectionFormula = "{Qtype.Typename}=" & _
    "'" & NewMsg.Cmbtype.Text & "'" & " And {Qtype.Region}=" & _
    "'" & NewMsg.CmbRegion.Text & "'"
    frmMain.CrystalReport1.Destination = 0
     frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
      Exit Sub
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(0).Value = True Then
    frmMain.CrystalReport1.ReportFileName = (App.Path & "\repcompany.rpt")
    frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Categoryname}like" & _
    "'*" & NewMsg.Text1.Text & "*'"
    frmMain.CrystalReport1.Destination = 0
     frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(1).Value = True Then

    frmMain.CrystalReport1.ReportFileName = (App.Path & "\repcompany.rpt")
    frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Companyname}like" & _
    "'*" & NewMsg.Text1.Text & "*'"
    frmMain.CrystalReport1.Destination = 0
     frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1

ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(2).Value = True Then
 frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
    frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Productname}like" & _
    "'*" & NewMsg.Text1.Text & "*'"
    frmMain.CrystalReport1.Destination = 0
     frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(3).Value = True Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\type.rpt")
    frmMain.CrystalReport1.SelectionFormula = "{Qtype.Typename}like" & _
    "'*" & NewMsg.Text1.Text & "*'"
    frmMain.CrystalReport1.Destination = 0
     frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\Mailing.mdb")
    frmMain.CrystalReport1.Action = 1

    '=====================================
'ElseIf frmrepReservation.CmbCategoryname.Text = "" And frmrepReservation.cmbCompanyname.Text = "" And _
'    frmrepReservation.CmbHallname.Text = "" And frmrepReservation.txtReservationNo.Text = "" And frmrepReservation.cmbStatus.Text <> "" Then
'    frmMain.CrystalReport1.ReportFileName = (App.Path & "\report\reservation.rpt")
'    frmMain.CrystalReport1.SelectionFormula = "{QResrv.Status}=" & _
'    "'" & frmrepReservation.cmbStatus.Text & "'"
'    frmMain.CrystalReport1.Destination = 0
'     frmMain.CrystalReport1.WindowState = crptMaximized
'     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\planDataBase.mdb")
'    frmMain.CrystalReport1.Action = 1
''===========================================
'ElseIf frmrepReservation.CmbCategoryname.Text = "" And frmrepReservation.cmbCompanyname.Text = "" And _
'    frmrepReservation.CmbHallname.Text = "" And frmrepReservation.txtReservationNo.Text <> "" And frmrepReservation.cmbStatus.Text = "" Then
'    frmMain.CrystalReport1.ReportFileName = (App.Path & "\report\reservation.rpt")
'    frmMain.CrystalReport1.SelectionFormula = "{QResrv.ReservationNo}=" & _
'    "" & frmrepReservation.txtReservationNo.Text & ""
'    frmMain.CrystalReport1.Destination = 0
'     frmMain.CrystalReport1.WindowState = crptMaximized
'     frmMain.CrystalReport1.DataFiles(0) = (App.Path & "\planDataBase.mdb")
'    frmMain.CrystalReport1.Action = 1
''===========================================

End If
End Sub

Private Sub Command1_Click()
Form_Activate
NewMsg.CmbProduct.Text = ""
NewMsg.CmbRegion.Text = ""
NewMsg.Cmbtype.Text = ""
NewMsg.Text1.Text = ""
NewMsg.cmbCity.Text = ""
frmMain.Show
End Sub

Private Sub Command2_Click()
NewMsg.Show
End Sub



Private Sub Option1_Click(Index As Integer)
NewMsg.CmbCategoryname.Text = ""
NewMsg.CmbProduct.Text = ""
NewMsg.cmbCompanyname.Text = ""
NewMsg.Cmbtype.Text = ""
NewMsg.Cmbcountry.Text = ""
NewMsg.cmbCity.Text = ""
NewMsg.CmbRegion.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
NewMsg.CmbCategoryname.Text = ""
NewMsg.CmbProduct.Text = ""
NewMsg.cmbCompanyname.Text = ""
NewMsg.Cmbtype.Text = ""
NewMsg.Cmbcountry.Text = ""
NewMsg.cmbCity.Text = ""
NewMsg.CmbRegion.Text = ""
End Sub
