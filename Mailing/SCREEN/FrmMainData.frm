VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMainData 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10725
   Icon            =   "FrmMainData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   10725
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   0
      Picture         =   "FrmMainData.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   10680
      TabIndex        =   36
      Top             =   -15
      Width           =   10740
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MAIN DATA"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   37
         Top             =   90
         Width           =   1650
      End
   End
   Begin VB.CommandButton cmdCMovelast 
      Caption         =   ">l"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Last Record"
      Top             =   3645
      Width           =   495
   End
   Begin VB.CommandButton cmdCmovenext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5205
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Next Record"
      Top             =   3645
      Width           =   495
   End
   Begin VB.CommandButton cmdCmoveprevious 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4695
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Previous Record"
      Top             =   3645
      Width           =   495
   End
   Begin VB.CommandButton cmdCmovefrist 
      Caption         =   "l<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Frist Record"
      Top             =   3645
      Width           =   495
   End
   Begin TabDlg.SSTab StbMainData 
      Height          =   3000
      Left            =   45
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   585
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   5292
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   706
      ShowFocusRect   =   0   'False
      BackColor       =   16777215
      TabCaption(0)   =   "Companies Categories"
      TabPicture(0)   =   "FrmMainData.frx":3F78
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtCategorycode"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtCategoryname"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lstcat"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Products"
      TabPicture(1)   =   "FrmMainData.frx":3F94
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstprod"
      Tab(1).Control(1)=   "CmbCategoryname"
      Tab(1).Control(2)=   "txtProductID"
      Tab(1).Control(3)=   "txtProductname"
      Tab(1).Control(4)=   "Label1(7)"
      Tab(1).Control(5)=   "Label1(2)"
      Tab(1).Control(6)=   "Label1(3)"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Cities and Countries"
      TabPicture(2)   =   "FrmMainData.frx":3FB0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstcity"
      Tab(2).Control(1)=   "Cmbcountry"
      Tab(2).Control(2)=   "cmbCity"
      Tab(2).Control(3)=   "txtRegion"
      Tab(2).Control(4)=   "txtcitycode"
      Tab(2).Control(5)=   "Label1(6)"
      Tab(2).Control(6)=   "Label1(4)"
      Tab(2).Control(7)=   "Label1(5)"
      Tab(2).Control(8)=   "Label1(17)"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Types"
      TabPicture(3)   =   "FrmMainData.frx":3FCC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstype"
      Tab(3).Control(1)=   "txtTypeID"
      Tab(3).Control(2)=   "txtTypeName"
      Tab(3).Control(3)=   "Label1(9)"
      Tab(3).Control(4)=   "Label1(8)"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Position Type"
      TabPicture(4)   =   "FrmMainData.frx":3FE8
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtTposname"
      Tab(4).Control(1)=   "txtTposID"
      Tab(4).Control(2)=   "lstpostype"
      Tab(4).Control(3)=   "Label1(11)"
      Tab(4).Control(4)=   "Label1(10)"
      Tab(4).ControlCount=   5
      Begin VB.TextBox txtTposname 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73110
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1095
         Width           =   3915
      End
      Begin VB.TextBox txtTposID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73110
         MaxLength       =   3
         TabIndex        =   33
         Top             =   720
         Width           =   630
      End
      Begin VB.ListBox lstpostype 
         Appearance      =   0  'Flat
         Height          =   1200
         Left            =   -68430
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   705
         Width           =   3495
      End
      Begin VB.ListBox lstcity 
         Appearance      =   0  'Flat
         Height          =   1200
         Left            =   -68160
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   900
         Width           =   3495
      End
      Begin VB.ListBox lstprod 
         Appearance      =   0  'Flat
         Height          =   1200
         Left            =   -68280
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   3495
      End
      Begin VB.ListBox lstcat 
         Appearance      =   0  'Flat
         Height          =   1200
         ItemData        =   "FrmMainData.frx":4004
         Left            =   6000
         List            =   "FrmMainData.frx":4006
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   750
         Width           =   4185
      End
      Begin VB.ListBox lstype 
         Appearance      =   0  'Flat
         Height          =   1200
         Left            =   -68160
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   3495
      End
      Begin VB.ComboBox Cmbcountry 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73395
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1230
         Width           =   3120
      End
      Begin VB.ComboBox cmbCity 
         Appearance      =   0  'Flat
         DataField       =   "City"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "FrmMainData.frx":4008
         Left            =   -73380
         List            =   "FrmMainData.frx":400A
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1590
         Width           =   3120
      End
      Begin VB.TextBox txtTypeID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73200
         MaxLength       =   3
         TabIndex        =   30
         Top             =   720
         Width           =   630
      End
      Begin VB.TextBox txtTypeName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73215
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1110
         Width           =   3915
      End
      Begin VB.ComboBox CmbCategoryname 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73320
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1185
         Width           =   3975
      End
      Begin VB.TextBox txtProductID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73320
         MaxLength       =   3
         TabIndex        =   26
         Top             =   840
         Width           =   630
      End
      Begin VB.TextBox txtProductname 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73320
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox txtRegion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73395
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1950
         Width           =   3120
      End
      Begin VB.TextBox txtcitycode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73395
         MaxLength       =   3
         TabIndex        =   21
         Top             =   900
         Width           =   630
      End
      Begin VB.TextBox txtCategoryname 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1935
         MaxLength       =   100
         TabIndex        =   0
         Top             =   1110
         Width           =   3915
      End
      Begin VB.TextBox txtCategorycode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1905
         MaxLength       =   3
         TabIndex        =   13
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C7E6FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Type Pos. name:"
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
         Index           =   11
         Left            =   -74505
         TabIndex        =   35
         Top             =   1095
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C7E6FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Type Pos. ID:"
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
         Index           =   10
         Left            =   -74265
         TabIndex        =   34
         Top             =   765
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C7E6FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Type ID:"
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
         Left            =   -74010
         TabIndex        =   32
         Top             =   765
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C7E6FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Type  name:"
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
         Index           =   8
         Left            =   -74295
         TabIndex        =   31
         Top             =   1095
         Width           =   975
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
         Index           =   7
         Left            =   -74760
         TabIndex        =   29
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C7E6FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Product ID:"
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
         Left            =   -74400
         TabIndex        =   28
         Top             =   840
         Width           =   900
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
         Index           =   3
         Left            =   -74640
         TabIndex        =   27
         Top             =   1560
         Width           =   1140
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
         Index           =   6
         Left            =   -74145
         TabIndex        =   25
         Top             =   1965
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C7E6FE&
         BackStyle       =   0  'Transparent
         Caption         =   "City code:"
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
         Left            =   -74325
         TabIndex        =   24
         Top             =   900
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C7E6FE&
         BackStyle       =   0  'Transparent
         Caption         =   "City name:"
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
         Index           =   5
         Left            =   -74415
         TabIndex        =   23
         Top             =   1620
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C7E6FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Country name:"
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
         Index           =   17
         Left            =   -74730
         TabIndex        =   22
         Top             =   1245
         Width           =   1200
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
         Index           =   1
         Left            =   480
         TabIndex        =   20
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C7E6FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Category ID:"
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
         Left            =   720
         TabIndex        =   19
         Top             =   720
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmMainData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PathPro As String

Private Sub asxToolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
Unload Me
End Sub

Private Sub CmbCategoryname_Change()
If frmMainData.CmbCategoryname.Text <> "" Then
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
SqLst1 = "Select distinct Productname From QProduct"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & frmMainData.CmbCategoryname.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
frmMainData.lstprod.Clear
    Do While Not rs1.EOF
      frmMainData.lstprod.AddItem rs1("Productname")
             rs1.MoveNext
    Loop
    End If
End Sub

Private Sub cmbCategoryname_Click()
If frmMainData.CmbCategoryname.Text <> "" Then
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
SqLst1 = "Select distinct Productname From QProduct"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & frmMainData.CmbCategoryname.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
frmMainData.lstprod.Clear
 If rs1.RecordCount <> 0 Then
    Do While Not rs1.EOF
      frmMainData.lstprod.AddItem rs1("Productname")
             rs1.MoveNext
    Loop
    frmMainData.lstprod.ListIndex = 0
Else
frmMainData.txtProductname.Text = ""
End If
    End If
End Sub

Private Sub cmbCity_Click()
If frmMainData.cmbCity.Text <> "" Then
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
SqLst1 = "Select distinct Region From City"
SqLst1 = SqLst1 & " WHERE city = '" & frmMainData.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
frmMainData.lstcity.Clear
    Do While Not rs1.EOF
        frmMainData.lstcity.AddItem rs1("Region")
       rs1.MoveNext
    Loop
     frmMainData.lstcity.ListIndex = 0
    End If
        
End Sub

Private Sub Cmbcountry_Change()
If frmMainData.Cmbcountry.Text <> "" Then
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
SqLst1 = "Select distinct city From City"
SqLst1 = SqLst1 & " WHERE country = '" & frmMainData.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
frmMainData.cmbCity.Clear
frmMainData.lstcity.Clear
 'Set rs = db.OpenRecordset("Category")
    Do While Not rs1.EOF
        frmMainData.cmbCity.AddItem rs1("city")
        frmMainData.lstcity.AddItem rs1("city")
       rs1.MoveNext
    Loop
    End If
End Sub

Private Sub Cmbcountry_Click()
If frmMainData.Cmbcountry.Text <> "" Then
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
SqLst1 = "Select distinct city From City"
SqLst1 = SqLst1 & " WHERE country = '" & frmMainData.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
frmMainData.cmbCity.Clear
frmMainData.lstcity.Clear
 'Set rs = db.OpenRecordset("Category")
    Do While Not rs1.EOF
        frmMainData.cmbCity.AddItem rs1("city")
        frmMainData.lstcity.AddItem rs1("city")
       rs1.MoveNext
    Loop
    End If
    frmMainData.lstcity.ListIndex = 0
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
Private Sub Form_Activate()
display.display (Me.Name)
Add.Mode
Dim Index As Integer
frmMain.Arrange Index
frmMain.hideb
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

Private Sub lstcat_Click()
If frmMainData.lstcat.Text <> "" Then
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
 SqLst = "Select * From Category"
        SqLst = SqLst & " WHERE Categoryname = '" & frmMainData.lstcat.Text & "'"
        'SqLst = SqLst & " and Companyname = '" & frmCampany.txtCompanyName.Text & "' ORDER BY CompanyId "
        Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
frmMainData.txtCategorycode.Text = rs("CategoryID")
frmMainData.txtCategoryname.Text = rs("Categoryname")
Else
End If
End If
End Sub

Private Sub lstcity_Click()
If frmMainData.lstcity.Text <> "" Then
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
 SqLst = "Select * From City"
        SqLst = SqLst & " WHERE city = '" & frmMainData.lstcity.Text & "'"
        'SqLst = SqLst & " and Companyname = '" & frmCampany.txtCompanyName.Text & "' ORDER BY CompanyId "
        Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
frmMainData.txtcitycode.Text = rs("citycode")
    frmMainData.txtRegion.Text = rs("Region")
    frmMainData.Cmbcountry.Text = rs("country")
     frmMainData.cmbCity.Text = rs("city")

Else
SqLst = "Select * From City"
SqLst = SqLst & " WHERE Region = '" & frmMainData.lstcity.Text & "'"
Set rs = db.OpenRecordset(SqLst)
If rs.RecordCount <> 0 Then
frmMainData.txtcitycode.Text = rs("citycode")
frmMainData.txtRegion.Text = rs("Region")
frmMainData.Cmbcountry.Text = rs("country")
frmMainData.cmbCity.Text = rs("city")
End If
End If
End If
End Sub

Private Sub lstpostype_Click()
If frmMainData.lstpostype.Text <> "" Then
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
 SqLst = "Select * From PosType"
        SqLst = SqLst & " WHERE posTypename = '" & frmMainData.lstpostype.Text & "'"
        'SqLst = SqLst & " and Companyname = '" & frmCampany.txtCompanyName.Text & "' ORDER BY CompanyId "
        Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
frmMainData.txtTposID.Text = rs("posTypeId")
    frmMainData.txtTposname.Text = rs("posTypename")
Else
End If
End If
End Sub

Private Sub lstprod_Click()
If frmMainData.lstprod.Text <> "" Then
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
 SqLst = "Select * From QProduct"
        SqLst = SqLst & " WHERE Productname = '" & frmMainData.lstprod.Text & "'"
        'SqLst = SqLst & " and Companyname = '" & frmCampany.txtCompanyName.Text & "' ORDER BY CompanyId "
        Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
frmMainData.txtProductID.Text = rs("ProductID")
    frmMainData.txtProductname.Text = rs("Productname")
    frmMainData.CmbCategoryname.Text = rs("Categoryname")
Else
End If
End If
End Sub

Private Sub lstype_Click()
If frmMainData.lstype.Text <> "" Then
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
 SqLst = "Select * From Type"
        SqLst = SqLst & " WHERE Typename = '" & frmMainData.lstype.Text & "'"
        'SqLst = SqLst & " and Companyname = '" & frmCampany.txtCompanyName.Text & "' ORDER BY CompanyId "
        Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
frmMainData.txtTypeID.Text = rs("TypeId")
    frmMainData.txtTypeName.Text = rs("Typename")
Else
End If
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
frmMain.hideb

End Sub

Private Sub StbMainData_Click(PreviousTab As Integer)
display.display (Me.Name)

End Sub

