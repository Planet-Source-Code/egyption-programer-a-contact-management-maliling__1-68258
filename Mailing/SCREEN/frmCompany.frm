VERSION 5.00
Begin VB.Form frmCampany 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10260
   Icon            =   "frmCompany.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   10260
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   -30
      Picture         =   "frmCompany.frx":0442
      ScaleHeight     =   465
      ScaleWidth      =   10365
      TabIndex        =   58
      Top             =   -30
      Width           =   10395
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COMPANY"
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
         Left            =   165
         TabIndex        =   59
         Top             =   105
         Width           =   1485
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   30
      TabIndex        =   49
      Top             =   2205
      Width           =   10215
      Begin VB.TextBox txtPbox 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9000
         MaxLength       =   10
         TabIndex        =   8
         Top             =   233
         Width           =   1065
      End
      Begin VB.TextBox txtPostelCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6765
         MaxLength       =   10
         TabIndex        =   7
         Top             =   233
         Width           =   870
      End
      Begin VB.ComboBox CmbRegion 
         Appearance      =   0  'Flat
         DataField       =   "City"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "frmCompany.frx":3F78
         Left            =   8010
         List            =   "frmCompany.frx":3F7A
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   750
         Width           =   1830
      End
      Begin VB.ComboBox Cmbcountry 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   150
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   750
         Width           =   2235
      End
      Begin VB.ComboBox cmbCity 
         Appearance      =   0  'Flat
         DataField       =   "City"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "frmCompany.frx":3F7C
         Left            =   3705
         List            =   "frmCompany.frx":3F7E
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   750
         Width           =   1935
      End
      Begin VB.TextBox txtaddress 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2100
         MaxLength       =   100
         TabIndex        =   6
         Top             =   233
         Width           =   3495
      End
      Begin VB.TextBox txt_addno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   615
         MaxLength       =   6
         TabIndex        =   5
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C7E6FE&
         BackStyle       =   0  'Transparent
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
         Left            =   75
         TabIndex        =   57
         Top             =   285
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C7E6FE&
         BackStyle       =   0  'Transparent
         Caption         =   "P.O. Box:"
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
         Index           =   32
         Left            =   8160
         TabIndex        =   56
         Top             =   285
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C7E6FE&
         BackStyle       =   0  'Transparent
         Caption         =   "Postal Code:"
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
         Left            =   5715
         TabIndex        =   55
         Top             =   285
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   135
         TabIndex        =   54
         Top             =   525
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   3735
         TabIndex        =   53
         Top             =   525
         Width           =   405
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
         Left            =   8055
         TabIndex        =   52
         Top             =   525
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "No.  :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   135
         TabIndex        =   51
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "St. Name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1290
         TabIndex        =   50
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.ComboBox cmbabbr 
      Height          =   315
      ItemData        =   "frmCompany.frx":3F80
      Left            =   8700
      List            =   "frmCompany.frx":3F82
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   48
      Top             =   585
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New Person"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4875
      Width           =   915
   End
   Begin VB.ListBox lst_person 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   6195
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   24
      Top             =   4875
      Width           =   3945
   End
   Begin VB.ComboBox cmbCompanyname 
      Height          =   315
      ItemData        =   "frmCompany.frx":3F84
      Left            =   3720
      List            =   "frmCompany.frx":3F86
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   570
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.CommandButton LblType 
      Caption         =   "Type name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7125
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   945
      Width           =   1575
   End
   Begin VB.CommandButton LblCategory 
      Caption         =   "Category name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   765
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   945
      Width           =   1575
   End
   Begin VB.TextBox txttel4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   765
      MaxLength       =   24
      TabIndex        =   15
      Top             =   3795
      Width           =   1845
   End
   Begin VB.TextBox txttel2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4215
      MaxLength       =   24
      TabIndex        =   13
      Top             =   3435
      Width           =   1935
   End
   Begin VB.TextBox txttel5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4215
      MaxLength       =   24
      TabIndex        =   16
      Top             =   3795
      Width           =   1935
   End
   Begin VB.ListBox LstType 
      Appearance      =   0  'Flat
      Height          =   930
      Left            =   7140
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1230
      Width           =   3030
   End
   Begin VB.ListBox Lstcategory 
      Appearance      =   0  'Flat
      Height          =   930
      Left            =   765
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   1230
      Width           =   3030
   End
   Begin VB.TextBox txtFax2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   780
      MaxLength       =   24
      TabIndex        =   18
      Top             =   4125
      Width           =   1845
   End
   Begin VB.TextBox txtEmail2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6195
      MaxLength       =   100
      TabIndex        =   21
      Top             =   4485
      Width           =   3960
   End
   Begin VB.ListBox LstProduct 
      Appearance      =   0  'Flat
      Height          =   930
      Left            =   3930
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1230
      Width           =   3030
   End
   Begin VB.TextBox txtNots 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   840
      Left            =   735
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   22
      Top             =   4875
      Width           =   4380
   End
   Begin VB.TextBox txtFax1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8310
      MaxLength       =   24
      TabIndex        =   17
      Top             =   3795
      Width           =   1845
   End
   Begin VB.TextBox txtCompanyName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3720
      MaxLength       =   100
      TabIndex        =   0
      Top             =   570
      Width           =   3780
   End
   Begin VB.TextBox txtcompanyID 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "jhjhjhkjh"
      Top             =   555
      Width           =   945
   End
   Begin VB.TextBox txttel1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   780
      MaxLength       =   24
      TabIndex        =   12
      Top             =   3450
      Width           =   1845
   End
   Begin VB.TextBox txttel3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   8310
      MaxLength       =   24
      TabIndex        =   14
      Top             =   3435
      Width           =   1845
   End
   Begin VB.TextBox txtAbbreviation 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8700
      MaxLength       =   20
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   585
      Width           =   1455
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   765
      MaxLength       =   100
      TabIndex        =   20
      Top             =   4515
      Width           =   3450
   End
   Begin VB.TextBox txtweb 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6195
      MaxLength       =   100
      TabIndex        =   19
      Top             =   4140
      Width           =   3960
   End
   Begin VB.CommandButton cmdcmovefrist 
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
      Left            =   4155
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Frist Record"
      Top             =   5805
      Width           =   495
   End
   Begin VB.CommandButton cmdcmoveprevious 
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
      Left            =   4650
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Previous Record"
      Top             =   5805
      Width           =   495
   End
   Begin VB.CommandButton cmdcmovenext 
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
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Next Record"
      Top             =   5805
      Width           =   495
   End
   Begin VB.CommandButton cmdcMovelast 
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
      Left            =   5655
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Last Record"
      Top             =   5805
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail (2):"
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
      Index           =   22
      Left            =   5205
      TabIndex        =   47
      Top             =   4500
      Width           =   855
   End
   Begin VB.Label LblProduct 
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
      Left            =   3975
      TabIndex        =   46
      Top             =   975
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel.1:"
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
      Index           =   29
      Left            =   285
      TabIndex        =   45
      Top             =   3435
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax 1:"
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
      Left            =   7620
      TabIndex        =   44
      Top             =   3840
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
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
      Height          =   225
      Index           =   15
      Left            =   2355
      TabIndex        =   43
      Top             =   585
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "Company ID:"
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
      Index           =   16
      Left            =   105
      TabIndex        =   42
      Top             =   585
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel.4:"
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
      Left            =   285
      TabIndex        =   41
      Top             =   3795
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "Abbreviation:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   7590
      TabIndex        =   40
      Top             =   600
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel.5:"
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
      Left            =   3690
      TabIndex        =   39
      Top             =   3825
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel.3:"
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
      Left            =   7605
      TabIndex        =   38
      Top             =   3465
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax 2:"
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
      Left            =   285
      TabIndex        =   37
      Top             =   4170
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
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
      Index           =   27
      Left            =   165
      TabIndex        =   36
      Top             =   4515
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel.2:"
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
      Index           =   28
      Left            =   3705
      TabIndex        =   35
      Top             =   3450
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
      Caption         =   "Web:"
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
      Index           =   30
      Left            =   5595
      TabIndex        =   34
      Top             =   4185
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C7E6FE&
      BackStyle       =   0  'Transparent
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
      Height          =   225
      Index           =   31
      Left            =   210
      TabIndex        =   33
      Top             =   4815
      Width           =   510
   End
End
Attribute VB_Name = "frmCampany"
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

Private Sub asxToolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
Unload Me
End Sub

Private Sub cmbabbr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frmCampany.searchAbbreviat
frmCampany.cmdCmovenext.Enabled = True
frmCampany.cmdCmoveprevious.Enabled = True
frmCampany.txtCompanyName.Visible = True
frmCampany.txtAbbreviation.Visible = True
frmCampany.cmbCompanyname.Visible = False
frmCampany.cmbabbr.Visible = False
frmCampany.txtcompanyID.Enabled = False
frmCampany.txtcompanyID.Locked = True
End If
End Sub

Private Sub cmbCity_Click()
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
SqLst1 = "Select distinct Region From City"
SqLst1 = SqLst1 & " WHERE city = '" & frmCampany.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
frmCampany.CmbRegion.Clear
 'Set rs = db.OpenRecordset("Category")
    Do While Not rs1.EOF
        frmCampany.CmbRegion.AddItem rs1("Region") & ""
       rs1.MoveNext
    Loop
End Sub

Private Sub cmbCity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmbCompanyname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frmCampany.search
frmCampany.cmdCmovenext.Enabled = True
frmCampany.cmdCmoveprevious.Enabled = True
frmCampany.txtCompanyName.Visible = True
frmCampany.txtAbbreviation.Visible = True
frmCampany.cmbCompanyname.Visible = False
frmCampany.cmbabbr.Visible = False
frmCampany.txtcompanyID.Enabled = False
frmCampany.txtcompanyID.Locked = True

End If
End Sub

Private Sub Cmbcountry_Click()
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
SqLst1 = "Select distinct city From City"
SqLst1 = SqLst1 & " WHERE country = '" & frmCampany.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
frmCampany.cmbCity.Clear
 'Set rs = db.OpenRecordset("Category")
    Do While Not rs1.EOF
        frmCampany.cmbCity.AddItem rs1("city")
       rs1.MoveNext
    Loop
    End Sub
Private Sub Cmbcountry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub CmbRegion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
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

Private Sub Command1_Click()
s = frmCampany.txtCompanyName.Text
'frmpersonal.Show
Add.compers (s)

End Sub

Private Sub Form_Activate()

Add.Mode
Dim Index As Integer
frmMain.Arrange Index
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
 SqLst = "Select name From Qpersonal"
SqLst = SqLst & " WHERE Companyname = '" & frmCampany.txtCompanyName.Text & "'"
frmCampany.lst_person.Clear
Set rs = db.OpenRecordset(SqLst)

Do While Not rs.EOF
frmCampany.lst_person.AddItem rs("name")
rs.MoveNext
Loop
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
display.display (Me.Name)
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
frmMain.hideb

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Add.ModeP
End Sub

Private Sub LblCategory_Click()
If frmCampany.txtcompanyID.Text = "" Then Exit Sub
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
frmCampany.LstProduct.Clear
     
        frmCampany.Lstcategory.Clear
         Set rs = db.OpenRecordset("Category")
        
     Do While Not rs.EOF
        frmCampany.Lstcategory.AddItem (rs("Categoryname"))
        rs.MoveNext
        Loop
        
 SqLst = "Select Categoryname From Qcatcomp"
SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""

Set rs = db.OpenRecordset(SqLst)

Do While Not rs.EOF
y = 0
Do While y < frmCampany.Lstcategory.ListCount
frmCampany.Lstcategory.ListIndex = y

If rs("Categoryname") = frmCampany.Lstcategory.List(frmCampany.Lstcategory.ListIndex) Then
frmCampany.Lstcategory.Selected(y) = True
GoTo e
Else
End If
y = y + 1

Loop
e:
rs.MoveNext
Loop
'=================================
 SqLst = "Select Productname From Qprocomp"
SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""

Set rs = db.OpenRecordset(SqLst)

Do While Not rs.EOF
y = 0
Do While y < frmCampany.LstProduct.ListCount
frmCampany.LstProduct.ListIndex = y

If rs("Productname") = frmCampany.LstProduct.List(frmCampany.LstProduct.ListIndex) Then
frmCampany.LstProduct.Selected(y) = True
GoTo e1
Else
End If
y = y + 1

Loop
e1:
rs.MoveNext
Loop
End Sub

Private Sub LblType_Click()
If frmCampany.txtcompanyID.Text = "" Then Exit Sub
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")

     frmCampany.LstType.Clear
    Set rs = db.OpenRecordset("Type")
     Do While Not rs.EOF
       frmCampany.LstType.AddItem rs("Typename")
        rs.MoveNext
    Loop
SqLst = "Select Typename From Qtype"
SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""

Set rs = db.OpenRecordset(SqLst)

Do While Not rs.EOF
y = 0
Do While y < frmCampany.LstType.ListCount
frmCampany.LstType.ListIndex = y

If rs("Typename") = frmCampany.LstType.List(frmCampany.LstType.ListIndex) Then
frmCampany.LstType.Selected(y) = True
GoTo e2
Else
End If
y = y + 1

Loop
e2:
rs.MoveNext
Loop
    
End Sub

Private Sub Lstcategory_ItemCheck(Item As Integer)

 Add.mode1
End Sub

Private Sub Lstcategory_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub LstProduct_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub LstType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
frmMain.hideb
End Sub

Private Sub txt_addno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 47 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
        
    End If

End Sub

Private Sub txtAbbreviation_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txtaddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txtcompanyID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frmCampany.searchcode
frmCampany.cmdCmovenext.Enabled = True
frmCampany.cmdCmoveprevious.Enabled = True
frmCampany.txtCompanyName.Visible = True
frmCampany.txtAbbreviation.Visible = True
frmCampany.cmbCompanyname.Visible = False
frmCampany.cmbabbr.Visible = False
frmCampany.txtcompanyID.Enabled = False
frmCampany.txtcompanyID.Locked = True
frmCampany.txtCompanyName.SetFocus
End If
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 47 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCompanyName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txtEmail2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txtPbox_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 47 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtPostelCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 47 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
        
    End If
End Sub


Private Sub txttel2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 47 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
        
    End If
End Sub
'

Private Sub txttel3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 47 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub
Private Sub txttel4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 47 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTel5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 47 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
        
    End If

End Sub

Private Sub txtFax1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 47 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
        
    End If
End Sub

Private Sub txtFax2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 43 Or KeyAscii = 45 Or KeyAscii = 47 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
        
    End If

End Sub

Private Sub txttel1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
If KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 43 Or _
KeyAscii = 45 Or KeyAscii = 47 Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
        
    End If

End Sub

Public Sub CompanyRec()
frmCampany.txtcompanyID.Text = rs("CompanyId")
frmCampany.txtCompanyName.Text = rs("Companyname")
frmCampany.txt_addno.Text = rs("addno")
'frmCampany.txt_blno.Text = rs("blno")
frmCampany.txtaddress.Text = rs("Address")
frmCampany.cmbCity.Text = rs("city")
frmCampany.Cmbcountry.Text = rs("country")
frmCampany.CmbRegion.Text = rs("Region")
frmCampany.txtPostelCode.Text = rs("PostalCode") & ""
frmCampany.txtPbox.Text = rs("Pobox") & ""
frmCampany.txttel1.Text = rs("tel1") & ""
frmCampany.txttel2.Text = rs("tel2") & ""
frmCampany.txttel3.Text = rs("tel3") & ""
frmCampany.txttel4.Text = rs("Tel4") & ""
frmCampany.txttel5.Text = rs("tel5") & ""
frmCampany.txtFax1.Text = rs("Fax1") & ""
frmCampany.txtFax2.Text = rs("Fax2") & ""
frmCampany.txtEmail.Text = rs("Email") & ""
frmCampany.txtEmail2.Text = rs("Email2") & ""
frmCampany.txtweb.Text = rs("web") & ""
frmCampany.txtNots.Text = rs("Notes") & ""
If IsNull(rs("Abbreviation")) Then
frmCampany.txtAbbreviation.Text = ""
Else
frmCampany.txtAbbreviation.Text = rs("Abbreviation")
End If
SqLst = "Select Categoryname From Qcatcomp"
SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
frmCampany.Lstcategory.Clear
Set rs2 = db.OpenRecordset(SqLst)
y = 0
Do While Not rs2.EOF
frmCampany.Lstcategory.AddItem rs2("Categoryname")
rs2.MoveNext
frmCampany.Lstcategory.Selected(y) = True
y = y + 1
Loop
SqLst = "Select Productname From Qprocomp"
SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
frmCampany.LstProduct.Clear
Set rs2 = db.OpenRecordset(SqLst)
y = 0
Do While Not rs2.EOF
frmCampany.LstProduct.AddItem rs2("Productname")
frmCampany.LstProduct.Selected(y) = True
y = y + 1
rs2.MoveNext
Loop
y = 0
SqLst = "Select Typename From Qtype"
SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
frmCampany.LstType.Clear
Set rs = db.OpenRecordset(SqLst)
Do While Not rs.EOF
frmCampany.LstType.AddItem rs("Typename")
frmCampany.LstType.Selected(y) = True
y = y + 1
rs.MoveNext
Loop
End Sub
Public Sub searchcode()
On Error GoTo RR
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
 SqLst = "Select * From Qcompany"
        SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
        'SqLst = SqLst & " and Companyname = '" & frmCampany.txtCompanyName.Text & "' ORDER BY CompanyId "
        Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
frmCampany.CompanyRec
Else
MsgBox ("This code is not Found")
Moving.Movefrist (Me.Name)
End If
RR:
If Err.Number = 3075 Then
MsgBox "Enter any value to complet search"
Moving.Movefrist (Me.Name)
End If
End Sub
Public Sub searchAbbreviat()
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
 SqLst = "Select * From Qcompany"
        SqLst = SqLst & " WHERE Abbreviation = '" & frmCampany.cmbabbr.Text & "'"
        'SqLst = SqLst & " and Companyname = '" & frmCampany.txtCompanyName.Text & "' ORDER BY CompanyId "
        Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
frmCampany.CompanyRec
Else
MsgBox ("This company is not Found")
Moving.Movefrist (Me.Name)
End If
End Sub
Public Sub search()
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
 SqLst = "Select * From Qcompany"
        SqLst = SqLst & " WHERE Companyname = '" & frmCampany.cmbCompanyname.Text & "'"
        'SqLst = SqLst & " and Companyname = '" & frmCampany.txtCompanyName.Text & "' ORDER BY CompanyId "
        Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
frmCampany.CompanyRec
Else
MsgBox ("This company is not Found")
Moving.Movefrist (Me.Name)
End If
End Sub

Private Sub txtweb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub
