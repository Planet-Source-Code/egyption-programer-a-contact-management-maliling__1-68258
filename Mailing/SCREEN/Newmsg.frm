VERSION 5.00
Begin VB.Form NewMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Note"
   ClientHeight    =   6510
   ClientLeft      =   1860
   ClientTop       =   2145
   ClientWidth     =   11250
   Icon            =   "Newmsg.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6510
   ScaleWidth      =   11250
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Caption         =   "Label and Letter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   2760
      TabIndex        =   29
      Top             =   3360
      Visible         =   0   'False
      Width           =   6495
      Begin VB.OptionButton Option2 
         Caption         =   "Persons Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   1305
         TabIndex        =   73
         Top             =   1290
         Width           =   1965
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Persons Letter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   3495
         TabIndex        =   72
         Top             =   1290
         Width           =   2070
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3270
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1680
         Width           =   1050
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2130
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1680
         Width           =   1050
      End
      Begin VB.OptionButton Option2 
         Caption         =   "companies Letter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3390
         TabIndex        =   31
         Top             =   525
         Width           =   2205
      End
      Begin VB.OptionButton Option2 
         Caption         =   "companies Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   30
         Top             =   525
         Value           =   -1  'True
         Width           =   2145
      End
      Begin VB.Line Line1 
         X1              =   75
         X2              =   6420
         Y1              =   1065
         Y2              =   1065
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      Caption         =   "Compose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1590
      Left            =   2220
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   7860
      Begin VB.Frame Frame7 
         Height          =   1335
         Left            =   5760
         TabIndex        =   35
         Top             =   150
         Width           =   2055
         Begin VB.OptionButton Option2 
            Caption         =   "Person email"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   225
            TabIndex        =   71
            Top             =   975
            Width           =   1770
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Another email"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   225
            TabIndex        =   37
            Top             =   585
            Width           =   1770
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Frist email"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   225
            TabIndex        =   36
            Top             =   165
            Value           =   -1  'True
            Width           =   1665
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Normal Compose"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   285
         TabIndex        =   34
         Top             =   465
         Value           =   -1  'True
         Width           =   2205
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Compose By Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   33
         Top             =   435
         Width           =   2445
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   1905
         TabIndex        =   15
         Top             =   1125
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Send"
         Height          =   330
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1125
         Width           =   1335
      End
   End
   Begin VB.Frame Frame8 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2070
      Left            =   3270
      TabIndex        =   38
      Top             =   1590
      Visible         =   0   'False
      Width           =   5160
      Begin VB.OptionButton Optlabper 
         Caption         =   "Lables for Persons do'nt have e mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   285
         TabIndex        =   70
         Top             =   1005
         Width           =   4260
      End
      Begin VB.OptionButton Optltrper 
         Caption         =   "Letter for Persons do'nt have e mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   300
         TabIndex        =   69
         Top             =   1335
         Width           =   4065
      End
      Begin VB.CommandButton Command6 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1740
         TabIndex        =   42
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2970
         TabIndex        =   41
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton optfaxem 
         Caption         =   "Letter for companies do'nt have e mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   270
         TabIndex        =   40
         Top             =   585
         Width           =   4350
      End
      Begin VB.OptionButton Optlablem 
         Caption         =   "Lables for companies do'nt have e mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   255
         TabIndex        =   39
         Top             =   270
         Value           =   -1  'True
         Width           =   4440
      End
   End
   Begin VB.ComboBox cmbpostype 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "Newmsg.frx":0442
      Left            =   4845
      List            =   "Newmsg.frx":0444
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      TabIndex        =   68
      Top             =   3090
      Width           =   2235
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   450
      Index           =   0
      Left            =   270
      TabIndex        =   61
      Top             =   2970
      Width           =   3330
      Begin VB.OptionButton Clearcompany 
         Caption         =   "Clear all"
         Height          =   240
         Left            =   1815
         TabIndex        =   63
         Top             =   150
         Width           =   1170
      End
      Begin VB.OptionButton Chkcompany 
         Caption         =   "Check all"
         Height          =   240
         Left            =   120
         TabIndex        =   62
         Top             =   135
         Width           =   1170
      End
   End
   Begin VB.ListBox Lstcompany 
      Appearance      =   0  'Flat
      Height          =   930
      Left            =   270
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   57
      Top             =   2010
      Width           =   3285
   End
   Begin VB.ListBox Lstpersonalname 
      Appearance      =   0  'Flat
      Height          =   930
      Left            =   3795
      RightToLeft     =   -1  'True
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   56
      Top             =   2010
      Width           =   3240
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1365
      Index           =   0
      Left            =   7140
      TabIndex        =   49
      Top             =   390
      Width           =   4005
      Begin VB.ComboBox cmbCity 
         DataField       =   "City"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "Newmsg.frx":0446
         Left            =   840
         List            =   "Newmsg.frx":0448
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   52
         Top             =   502
         Width           =   3105
      End
      Begin VB.ComboBox Cmbcountry 
         Height          =   315
         Left            =   840
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   51
         Top             =   135
         Width           =   3105
      End
      Begin VB.ComboBox CmbRegion 
         DataField       =   "City"
         DataSource      =   "Data1"
         Height          =   315
         ItemData        =   "Newmsg.frx":044A
         Left            =   840
         List            =   "Newmsg.frx":044C
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   50
         Top             =   870
         Width           =   3105
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
         Left            =   345
         TabIndex        =   55
         Top             =   510
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
         Left            =   165
         TabIndex        =   54
         Top             =   885
         Width           =   615
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
         Left            =   30
         TabIndex        =   53
         Top             =   135
         Width           =   720
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   1590
      Left            =   7125
      TabIndex        =   43
      Top             =   1830
      Width           =   4005
      Begin VB.OptionButton Option1 
         Caption         =   "Personal  name"
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
         Index           =   5
         Left            =   945
         TabIndex        =   64
         Top             =   1215
         Width           =   1650
      End
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
         Left            =   2055
         TabIndex        =   48
         Top             =   925
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
         Left            =   2055
         TabIndex        =   47
         Top             =   635
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
         Left            =   255
         TabIndex        =   46
         Top             =   925
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
         Left            =   255
         TabIndex        =   45
         Top             =   615
         Width           =   1080
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   255
         TabIndex        =   44
         Top             =   210
         Width           =   3390
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   255
      TabIndex        =   25
      Top             =   3405
      Width           =   10890
      Begin VB.CommandButton cmdCleara 
         Caption         =   "Clear Attach"
         Enabled         =   0   'False
         Height          =   285
         Left            =   9075
         TabIndex        =   67
         Top             =   975
         Width           =   1740
      End
      Begin VB.PictureBox picAttach 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   10365
         Picture         =   "Newmsg.frx":044E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   66
         Top             =   135
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.TextBox txtsubject 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         MaxLength       =   150
         TabIndex        =   6
         Top             =   930
         Width           =   8160
      End
      Begin VB.TextBox txtcc 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Top             =   585
         Width           =   8160
      End
      Begin VB.TextBox txtTo 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   210
         Width           =   8160
      End
      Begin VB.Label lblattach 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   9045
         TabIndex        =   65
         Top             =   585
         Width           =   1710
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subj&ect:"
         Height          =   195
         Left            =   90
         TabIndex        =   28
         Top             =   975
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Cc:"
         Height          =   195
         Left            =   255
         TabIndex        =   27
         Top             =   570
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&To:"
         Height          =   195
         Left            =   255
         TabIndex        =   26
         Top             =   255
         Width           =   300
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   255
      TabIndex        =   20
      Top             =   390
      Width           =   6840
      Begin VB.ComboBox cmbCompanyname 
         Height          =   315
         ItemData        =   "Newmsg.frx":0890
         Left            =   3480
         List            =   "Newmsg.frx":0892
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   825
         Width           =   3225
      End
      Begin VB.ComboBox CmbCategoryname 
         Height          =   315
         Left            =   75
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   330
         Width           =   3225
      End
      Begin VB.ComboBox CmbProduct 
         Height          =   315
         ItemData        =   "Newmsg.frx":0894
         Left            =   3525
         List            =   "Newmsg.frx":0896
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   315
         Width           =   3225
      End
      Begin VB.ComboBox Cmbtype 
         Height          =   315
         ItemData        =   "Newmsg.frx":0898
         Left            =   75
         List            =   "Newmsg.frx":089A
         RightToLeft     =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   885
         Width           =   3225
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
         Left            =   45
         TabIndex        =   24
         Top             =   105
         Width           =   1260
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
         Left            =   3525
         TabIndex        =   23
         Top             =   615
         Width           =   1290
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
         Left            =   3480
         TabIndex        =   22
         Top             =   90
         Width           =   1140
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
         Left            =   75
         TabIndex        =   21
         Top             =   690
         Width           =   1185
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11250
      TabIndex        =   19
      Top             =   0
      Width           =   11250
      Begin VB.CommandButton Refresh 
         Caption         =   "Refresh"
         Height          =   330
         Left            =   9555
         TabIndex        =   14
         Top             =   60
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Report "
         Height          =   330
         Left            =   8010
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   60
         Width           =   1335
      End
      Begin VB.CommandButton CompAdd 
         Caption         =   "Label and Letter"
         Height          =   330
         Left            =   6435
         TabIndex        =   12
         Top             =   60
         Width           =   1335
      End
      Begin VB.CommandButton CompOpt 
         Caption         =   "Op&tions"
         Enabled         =   0   'False
         Height          =   330
         Left            =   3315
         TabIndex        =   11
         Top             =   60
         Width           =   1335
      End
      Begin VB.CommandButton Attach 
         Caption         =   "&Attach"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1755
         TabIndex        =   10
         Top             =   60
         Width           =   1335
      End
      Begin VB.CommandButton ChkNames1 
         Caption         =   "Op&tions lab. fax"
         Height          =   330
         Left            =   4875
         TabIndex        =   9
         Top             =   60
         Width           =   1335
      End
      Begin VB.CommandButton Send 
         Caption         =   "&Send"
         Enabled         =   0   'False
         Height          =   330
         Left            =   195
         TabIndex        =   8
         Top             =   60
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   -690
         X2              =   -690
         Y1              =   0
         Y2              =   540
      End
      Begin VB.Line TopLine2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   11625
         Y1              =   15
         Y2              =   0
      End
   End
   Begin VB.TextBox txtNoteText 
      Appearance      =   0  'Flat
      Height          =   1635
      Left            =   270
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4710
      Width           =   10875
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Persons name:"
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
      Index           =   3
      Left            =   3840
      TabIndex        =   60
      Top             =   1755
      Width           =   1155
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Persons type:"
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
      Index           =   1
      Left            =   3765
      TabIndex        =   59
      Top             =   3135
      Width           =   1050
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   16
      Left            =   300
      TabIndex        =   58
      Top             =   1740
      Width           =   1290
   End
End
Attribute VB_Name = "NewMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim ws As Workspace
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim SqLst As String
Dim SqLst1 As String
Dim PathPro As String
Dim s As String
Private Sub Attach_Click()
' Handle attachments.
On Error Resume Next
   frmMain.CMDialog1.DialogTitle = "Attach"
   frmMain.CMDialog1.Filter = "All Files(*.*)|*.*|Text Files(*.txt)|*.txt"
   frmMain.CMDialog1.ShowOpen
   If Err = 0 Then
        On Error GoTo 0
        frmMain.MapiMess.AttachmentIndex = frmMain.MapiMess.AttachmentCount
        frmMain.MapiMess.AttachmentName = frmMain.CMDialog1.FileTitle
        NewMsg.lblattach.Caption = frmMain.CMDialog1.FileTitle
        NewMsg.picAttach.Visible = True
        NewMsg.cmdCleara.Enabled = True
        NewMsg.Attach.Enabled = False
        frmMain.MapiMess.AttachmentPathName = frmMain.CMDialog1.FileName
        frmMain.MapiMess.AttachmentPosition = frmMain.MapiMess.AttachmentIndex
        frmMain.MapiMess.AttachmentType = vbAttachTypeData
   End If
End Sub

Private Sub Chkcompany_Click()
If NewMsg.Chkcompany.Value = True Then
y = NewMsg.Lstcompany.ListCount
Do While Not y <= 0
y = y - 1
NewMsg.Lstcompany.Selected(y) = True

Loop
NewMsg.Lstcompany.Refresh
 'NewMsg.Chkcompany.Value = False
End If
End Sub

Private Sub ChkNames1_Click()
    ' Resolve the names.
    'Call CopyNamestoMsgBuffer(Me, True)
    'Call UpdateRecips(Me)
    If NewMsg.txtNoteText = "" Then
    NewMsg.optfaxem.Enabled = False
    NewMsg.Optltrper.Enabled = False
    End If
    NewMsg.Frame8.Visible = True
End Sub

Private Sub refSearsh()

Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
   NewMsg.CmbCategoryname.Clear
 Set rs = db.OpenRecordset("Category")
     Do While Not rs.EOF
       NewMsg.CmbCategoryname.AddItem rs("Categoryname")
        rs.MoveNext
    Loop
    NewMsg.cmbCompanyname.Clear
  NewMsg.Lstpersonalname.Clear
    Set rs = db.OpenRecordset("Company")
     NewMsg.Lstcompany.Clear
y = 0
'NewMsg.cmbCompanyname.AddItem "All"
s = rs.RecordCount
    Do While Not rs.EOF
       NewMsg.cmbCompanyname.AddItem rs("Companyname")
       NewMsg.Lstcompany.AddItem rs("Companyname")
       'NewMsg.Lstcompany.Selected(y) = True
       rs.MoveNext
 y = y + 1
       Loop
    NewMsg.cmbpostype.Clear
   '  NewMsg.CmbProduct.AddItem "All"
    Set rs = db.OpenRecordset("PosType")
    
     Do While Not rs.EOF
       NewMsg.cmbpostype.AddItem rs("posTypename")
        rs.MoveNext
    Loop
    
      NewMsg.Cmbtype.Clear
   '  NewMsg.Cmbtype.AddItem "All"
    Set rs = db.OpenRecordset("Type")
    
     Do While Not rs.EOF
       NewMsg.Cmbtype.AddItem rs("Typename")
        rs.MoveNext
    Loop
    NewMsg.Cmbcountry.Clear
SqLst1 = "Select distinct country From city"
Set rs = db.OpenRecordset(SqLst1)
Do While Not rs.EOF
NewMsg.Cmbcountry.AddItem rs("country")
rs.MoveNext
Loop
NewMsg.cmbCity.Clear
NewMsg.CmbRegion.Clear

End Sub



Private Sub Clearcompany_Click()
If Clearcompany.Value = True Then
y = NewMsg.Lstcompany.ListCount
Do While Not y <= 0
y = y - 1
NewMsg.Lstcompany.Selected(y) = False
Loop
NewMsg.Lstpersonalname.Clear
End If

End Sub


Private Sub CmbCategoryname_LostFocus()
If NewMsg.CmbCategoryname.Text = "" Then
NewMsg.CmbProduct.Clear
End If
End Sub

Private Sub cmbCity_LostFocus()
If NewMsg.cmbCity = "" Then
NewMsg.CmbRegion.Clear
Else
End If
End Sub

Private Sub Cmbcountry_LostFocus()
If NewMsg.Cmbcountry = "" Then
NewMsg.cmbCity.Clear
NewMsg.CmbRegion.Clear
Else
End If
End Sub

Private Sub cmbpostype_Click()
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
NewMsg.Lstpersonalname.Clear
X = 0
Do While X < NewMsg.Lstcompany.ListCount
NewMsg.Lstcompany.ListIndex = X
If NewMsg.Lstcompany.Selected(NewMsg.Lstcompany.ListIndex) = True Then
SqLst1 = "Select name From Qper"
SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.Lstcompany.List(NewMsg.Lstcompany.ListIndex)
SqLst1 = SqLst1 & "' and posTypename = '" & NewMsg.cmbpostype.Text & "'"

Set rs1 = db.OpenRecordset(SqLst1, dbOpenDynaset)

Do While Not rs1.EOF
NewMsg.Lstpersonalname.AddItem rs1("name")
rs1.MoveNext
Loop
y = NewMsg.Lstpersonalname.ListCount
Do While Not y <= 0
y = y - 1
NewMsg.Lstpersonalname.Selected(y) = True
Loop

X = X + 1
Else

X = X + 1


End If
Loop
NewMsg.Text1.Text = ""
End Sub

Private Sub CmbRegion_Click()
If NewMsg.cmbCity.Text = "" Then
NewMsg.CmbRegion.Clear
End If
 Add.searshCom
 NewMsg.Text1.Text = ""
End Sub

Private Sub cmdCleara_Click()
NewMsg.picAttach.Visible = False
NewMsg.lblattach.Caption = ""
NewMsg.Attach.Enabled = True
NewMsg.cmdCleara.Enabled = False
frmMain.MapiMess.Delete (2)

End Sub

Private Sub Command3_Click()
If NewMsg.Option2(3).Value = True Then
    If NewMsg.txtTo.Text = "" Then
    MsgBox ("Please...Complete Data")
    Else
     NewMsg.txtTo.Text = ";" & NewMsg.txtTo.Text
    If NewMsg.txtcc.Text <> "" Then
    NewMsg.txtcc.Text = ";" & NewMsg.txtcc.Text
    End If
    Module1.Msendmail
    End If
ElseIf NewMsg.Option2(4).Value = True Then
Module1.searshm1 (0)
'X = MsgBox("", vbYesNo)
's = 1
'If X = 6 Then

'Else
'End If
ElseIf NewMsg.Option2(5).Value = True Then
Module1.searshm (0)
ElseIf NewMsg.Option2(6).Value = True Then
Module1.sendMper (0)
End If
NewMsg.Frame6.Visible = False
End Sub

Private Sub Command1_Click()
 NewMsg.Frame5.Visible = False
End Sub

Private Sub Command2_Click()
On Error GoTo eh:
    If NewMsg.Option2(0).Value = True Then
        s = ""
        mLable (s)
    
    ElseIf NewMsg.Option2(1).Value = True Then
        frmMain.CrystalReport1.ReportFileName = (App.Path & "\fax.rpt")
        frmMain.CrystalReport1.SelectionFormula = ""
        frmMain.CrystalReport1.Destination = 0
        frmMain.CrystalReport1.WindowState = crptMaximized
        frmMain.CrystalReport1.DataFiles(0) = (PathPro)
        frmMain.CrystalReport1.Action = 1
    
    ElseIf NewMsg.Option2(7).Value = True Then
            SqLst = "DELETE * FROM perNotes "
                DBEngine.Workspaces(0).BeginTrans
                db.Execute SqLst
                DBEngine.Workspaces(0).CommitTrans
        Do While X < NewMsg.Lstpersonalname.ListCount
        NewMsg.Lstpersonalname.ListIndex = X
        If NewMsg.Lstpersonalname.Selected(NewMsg.Lstpersonalname.ListIndex) = True Then
        SqLst1 = "Select personalId From Personal"
        SqLst1 = SqLst1 & " WHERE name = '" & NewMsg.Lstpersonalname.List(NewMsg.Lstpersonalname.ListIndex) & "'"
        Set rs1 = db.OpenRecordset(SqLst1, dbOpenDynaset)
         If rs1.RecordCount <> 0 Then
           A = rs1("personalId")
           Set rs = db.OpenRecordset("perNotes")
           rs.AddNew
           rs("personalId") = A
            rs("Notes") = NewMsg.txtNoteText.Text
            rs("Supject") = NewMsg.txtsubject.Text
           rs.Update
            End If
               X = X + 1
               Else
               X = X + 1
               End If
            Loop
           
           ' End If
                frmMain.CrystalReport1.ReportFileName = (App.Path & "\perfax.rpt")
                frmMain.CrystalReport1.SelectionFormula = ""
                frmMain.CrystalReport1.Destination = 0
                frmMain.CrystalReport1.WindowState = crptMaximized
                 frmMain.CrystalReport1.DataFiles(0) = (PathPro)
                frmMain.CrystalReport1.Action = 1
        
    ElseIf NewMsg.Option2(8).Value = True Then
    SqLst = "DELETE * FROM perNotes "
                DBEngine.Workspaces(0).BeginTrans
                db.Execute SqLst
                DBEngine.Workspaces(0).CommitTrans
        Do While X < NewMsg.Lstpersonalname.ListCount
        NewMsg.Lstpersonalname.ListIndex = X
        If NewMsg.Lstpersonalname.Selected(NewMsg.Lstpersonalname.ListIndex) = True Then
        SqLst1 = "Select personalId From Personal"
        SqLst1 = SqLst1 & " WHERE name = '" & NewMsg.Lstpersonalname.List(NewMsg.Lstpersonalname.ListIndex) & "'"
        Set rs1 = db.OpenRecordset(SqLst1, dbOpenDynaset)
         If rs1.RecordCount <> 0 Then
           A = rs1("personalId")
           Set rs = db.OpenRecordset("perNotes")
           rs.AddNew
           rs("personalId") = A
            rs.Update
            End If
               X = X + 1
               Else
               X = X + 1
               End If
            Loop
           
            'End If
        frmMain.CrystalReport1.ReportFileName = (App.Path & "\perlabel.rpt")
        frmMain.CrystalReport1.SelectionFormula = ""
        frmMain.CrystalReport1.Destination = 0
        frmMain.CrystalReport1.WindowState = crptMaximized
         frmMain.CrystalReport1.DataFiles(0) = (PathPro)
        frmMain.CrystalReport1.Action = 1
    End If
     NewMsg.Frame5.Visible = False
eh:
     If Err.Number = 20526 Then MsgBox ("Please Add printer")
End Sub
Private Sub mLable(s As String)
's = ""
's = "{Qcompany.Email}=" & "'Empty'"
's = "{Qcompany.Email}<>" & "'Empty'"
's = "{Qcompany.Email1}=" & "'Empty'"
's = "{Qcompany.Email1}<>" & "'Empty'"
On Error GoTo eh

If NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\lab.rpt")
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = "{Qcompany.Email}=" & "'Empty'"
frmMain.CrystalReport1.SelectionFormula = s
GoTo el1
Exit Sub
'1===========================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\lab.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcompany.Companyname}=" & _
"'" & NewMsg.cmbCompanyname.Text & "'" & s
GoTo el1
Exit Sub
'2===========================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\lab.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcompany.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & s
GoTo el1
Exit Sub
'3===========================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\lab.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcompany.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & " And {Qcompany.City}=" & _
"'" & NewMsg.cmbCity.Text & "'" & s
GoTo el1
Exit Sub
'4===========================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\lab.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcompany.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & " And {Qcompany.City}=" & _
"'" & NewMsg.cmbCity.Text & "'" & " And {Qcompany.Region}=" & _
"'" & NewMsg.CmbRegion.Text & "'" & s
GoTo el1
Exit Sub
'5===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text <> "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qcatype.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\labcatype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatype.Categoryname}=" & _
"'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & s
GoTo el1
Exit Sub
'6===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text <> "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qprodtype.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\labprodtype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprodtype.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'" & " And {Qprodtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & s
GoTo el1
Exit Sub
'7===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qcatype.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\labcatype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatype.Categoryname}=" & _
"'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qcatype.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & s
GoTo el1
Exit Sub
'8===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qprodtype.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\labprodtype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprodtype.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'" & " And {Qprodtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qprodtype.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & s
GoTo el1
Exit Sub
'9===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qcatype.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\labcatype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatype.Categoryname}=" & _
"'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qcatype.Qcompany.City}=" & _
"'" & NewMsg.cmbCity.Text & "'" & " And {Qcatype.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & s
GoTo el1
Exit Sub
'10===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qprodtype.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\labprodtype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprodtype.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'" & " And {Qprodtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qprodtype.Qcompany.City}=" & _
"'" & NewMsg.cmbCity.Text & "'" & " And {Qprodtype.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & s
GoTo el1
Exit Sub
'11===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qcatype.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\labcatype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatype.Categoryname}=" & _
"'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qcatype.Qcompany.City}=" & _
"'" & NewMsg.cmbCity.Text & "'" & " And {Qcatype.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & " And {Qcatype.Region}=" & _
"'" & NewMsg.CmbRegion.Text & "'" & s
GoTo el1
Exit Sub
'12===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qprodtype.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\labprodtype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprodtype.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'" & " And {Qprodtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qprodtype.Qcompany.City}=" & _
"'" & NewMsg.cmbCity.Text & "'" & " And {Qprodtype.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & " And {Qprodtype.Region}=" & _
"'" & NewMsg.CmbRegion.Text & "'" & s
GoTo el1
Exit Sub
'13===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qcatcomp.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\Catlable.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Categoryname}=" & _
"'" & NewMsg.CmbCategoryname.Text & "'" & s
GoTo el1
Exit Sub

ElseIf NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qprocomp.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\prolables.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'" & s
GoTo el1
Exit Sub
'15=============================================

ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbtype.Text <> "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qtype.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\typelabels.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & s
GoTo el1
Exit Sub
'16=============================================

ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qcatcomp.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\catlabel.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Categoryname}=" & _
"'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatcomp.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & s
GoTo el1
Exit Sub
'17====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qcatcomp.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\catlabel.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Categoryname}=" & _
"'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatcomp.city}=" & _
"'" & NewMsg.cmbCity.Text & "'" & s
GoTo el1
Exit Sub
''18====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qcatcomp.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\catlabel.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Categoryname}=" & _
"'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatcomp.Region}=" & _
"'" & NewMsg.CmbRegion.Text & "'" & s
GoTo el1
Exit Sub
'19====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qprocomp.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\prolables.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'" & " And {Qprocomp.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & s
GoTo el1
Exit Sub
'20====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qprocomp.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\prolables.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'" & " And {Qprocomp.city}=" & _
"'" & NewMsg.cmbCity.Text & "'" & s
GoTo el1
Exit Sub
''21====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qprocomp.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\prolables.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'" & " And {Qprocomp.Region}=" & _
"'" & NewMsg.CmbRegion.Text & "'" & s
GoTo el1
Exit Sub
'22====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qtype.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\typelabels.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qtype.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & s
GoTo el1
Exit Sub
'23====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text = "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qtype.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\typelabels.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qtype.city}=" & _
"'" & NewMsg.cmbCity.Text & "'" & s
GoTo el1
Exit Sub
''24====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text <> "" Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qtype.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\typelabels.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qtype.Region}=" & _
"'" & NewMsg.CmbRegion.Text & "'" & s
GoTo el1
Exit Sub
'25=======================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(0).Value = True Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qcatcomp.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\catlabel.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Categoryname}like" & _
"'*" & NewMsg.Text1.Text & "*'" & s
GoTo el1
Exit Sub
'26=======================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(1).Value = True Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\lab.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcompany.Companyname}like" & _
"'*" & NewMsg.Text1.Text & "*'" & s
GoTo el1
Exit Sub
'27=======================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(2).Value = True Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qprocomp.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\prolables.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Productname}like" & _
"'*" & NewMsg.Text1.Text & "*'" & s
GoTo el1
Exit Sub
'28=======================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(3).Value = True Then
If s = " And {Qcompany.Email}=" & "'Empty'" Then s = " And {Qtype.Email}=" & "'Empty'"
frmMain.CrystalReport1.ReportFileName = (App.Path & "\typelabels.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qtype.Typename}like" & _
"'*" & NewMsg.Text1.Text & "*'" & s

el1:
    frmMain.CrystalReport1.Destination = 0
    frmMain.CrystalReport1.WindowState = crptMaximized
    frmMain.CrystalReport1.DataFiles(0) = (PathPro)
    frmMain.CrystalReport1.Action = 1

  
End If
eh:

 If Err.Number = 20526 Then MsgBox ("please Add Printer")
End Sub
Private Sub Command4_Click()
NewMsg.Frame6.Visible = False
End Sub

Private Sub Command5_Click()
On Error GoTo e
If NewMsg.Optlablem.Value = True Then
s = " And {Qcompany.Email}=" & "'Empty'"
mLable (s)
ElseIf NewMsg.optfaxem.Value = True Then
Module1.senNotesSave
frmMain.CrystalReport1.ReportFileName = (App.Path & "\fax.rpt")
    frmMain.CrystalReport1.SelectionFormula = ""
    frmMain.CrystalReport1.Destination = 0
    frmMain.CrystalReport1.WindowState = crptMaximized
     frmMain.CrystalReport1.DataFiles(0) = (PathPro)
    frmMain.CrystalReport1.Action = 1
 ElseIf NewMsg.Optltrper.Value = True Then
            SqLst = "DELETE * FROM perNotes "
                DBEngine.Workspaces(0).BeginTrans
                db.Execute SqLst
                DBEngine.Workspaces(0).CommitTrans
        Do While X < NewMsg.Lstpersonalname.ListCount
        NewMsg.Lstpersonalname.ListIndex = X
        If NewMsg.Lstpersonalname.Selected(NewMsg.Lstpersonalname.ListIndex) = True Then
        SqLst1 = "Select personalId From Personal"
        SqLst1 = SqLst1 & " WHERE name = '" & NewMsg.Lstpersonalname.List(NewMsg.Lstpersonalname.ListIndex)
        SqLst1 = SqLst1 & "' and Email = " & "'Empty'"
        Set rs1 = db.OpenRecordset(SqLst1, dbOpenDynaset)
        If rs1.RecordCount <> 0 Then
           A = rs1("personalId")
           Set rs = db.OpenRecordset("perNotes")
           rs.AddNew
           rs("personalId") = A
            rs("Notes") = NewMsg.txtNoteText.Text
            rs("Supject") = NewMsg.txtsubject.Text
           rs.Update
            End If
               X = X + 1
               Else
               X = X + 1
               End If
            Loop
           
            'End If
                frmMain.CrystalReport1.ReportFileName = (App.Path & "\perfax.rpt")
                frmMain.CrystalReport1.SelectionFormula = ""
                frmMain.CrystalReport1.Destination = 0
                frmMain.CrystalReport1.WindowState = crptMaximized
                 frmMain.CrystalReport1.DataFiles(0) = (PathPro)
                frmMain.CrystalReport1.Action = 1
        
    ElseIf NewMsg.Optlabper.Value = True Then
    SqLst = "DELETE * FROM perNotes "
                DBEngine.Workspaces(0).BeginTrans
                db.Execute SqLst
                DBEngine.Workspaces(0).CommitTrans
        Do While X < NewMsg.Lstpersonalname.ListCount
        NewMsg.Lstpersonalname.ListIndex = X
        If NewMsg.Lstpersonalname.Selected(NewMsg.Lstpersonalname.ListIndex) = True Then
       ' If NewMsg.Lstpersonalname.Selected(NewMsg.Lstpersonalname.ListIndex) = True Then
        SqLst1 = "Select personalId From Personal"
        SqLst1 = SqLst1 & " WHERE name = '" & NewMsg.Lstpersonalname.List(NewMsg.Lstpersonalname.ListIndex)
        SqLst1 = SqLst1 & "' and Email = " & "'Empty'"
        Set rs1 = db.OpenRecordset(SqLst1, dbOpenDynaset)
         If rs1.RecordCount <> 0 Then
           A = rs1("personalId")
           Set rs = db.OpenRecordset("perNotes")
           rs.AddNew
           rs("personalId") = A
            rs.Update
            End If
                X = X + 1
               Else
               X = X + 1
               End If
            Loop
           
            'End If
        frmMain.CrystalReport1.ReportFileName = (App.Path & "\perlabel.rpt")
        frmMain.CrystalReport1.SelectionFormula = ""
        frmMain.CrystalReport1.Destination = 0
        frmMain.CrystalReport1.WindowState = crptMaximized
         frmMain.CrystalReport1.DataFiles(0) = (PathPro)
        frmMain.CrystalReport1.Action = 1
End If
NewMsg.Frame8.Visible = False
e:

If Err.Number = 20526 Then MsgBox ("Please Add printer")

End Sub

Private Sub Command6_Click()
NewMsg.Frame8.Visible = False
End Sub

Private Sub Command7_Click()

End Sub

Private Sub CompAdd_Click()
  If NewMsg.txtNoteText.Text = "" Then
NewMsg.Option2(1).Enabled = False
NewMsg.Option2(7).Enabled = False
Else
NewMsg.Option2(1).Enabled = True
NewMsg.Option2(7).Enabled = True
End If
  NewMsg.Frame5.Visible = True
  
    Module1.NotesSave
    
    If NewMsg.Option2(0).Value = True Then
    ElseIf NewMsg.Option2(1).Value = True Then
    End If
    
End Sub

Private Sub CompOpt_Click()
    ' Display the Message Option form.
    OptionType = conOptionMessage
    MailOptFrm.Show 1
End Sub

Private Sub Form_Activate()
On Error Resume Next
If frmMain.Logon.Enabled = False Then
NewMsg.Send.Enabled = True
'NewMsg.ChkNames.Enabled = True
NewMsg.Attach.Enabled = True
NewMsg.CompOpt.Enabled = True
'NewMsg.CompAdd.Enabled = True
End If

refSearsh
    ' Set the MessageIndex to -1 (Compose Buffer) when this window is activated.
    frmMain.MapiMess.MsgIndex = -1
Dim Index As Integer
frmMain.Arrange Index

frmMain.Toolbar1.Buttons(5).Enabled = True
End Sub
Private Sub cmbCategoryname_Click()
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
'SqLst1 = "Select Companyname From Qcatcomp"
'SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text & "'"
'Set rs1 = db.OpenRecordset(SqLst1)
'NewMsg.cmbCompanyname.Clear
'NewMsg.Lstcompany.Clear
'y = 0
'NewMsg.cmbCompanyname.AddItem "All"
'    Do While Not rs1.EOF
'       NewMsg.cmbCompanyname.AddItem rs1("Companyname")
'       NewMsg.Lstcompany.AddItem rs1("Companyname")
'       NewMsg.Lstcompany.Selected(y) = True
'       rs1.MoveNext
' y = y + 1
'       Loop

SqLst1 = "Select Productname From QProduct"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
NewMsg.CmbProduct.Clear
'NewMsg.CmbProduct.AddItem "All"
    Do While Not rs1.EOF
       NewMsg.CmbProduct.AddItem rs1("Productname")
       rs1.MoveNext
       Loop
Add.searshCom
NewMsg.Text1.Text = ""
End Sub



Private Sub cmbCity_Click()
If NewMsg.Cmbcountry.Text = "" Then
NewMsg.cmbCity.Clear
NewMsg.CmbRegion.Clear
Else
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
SqLst1 = "Select distinct Region From QCity"
SqLst1 = SqLst1 & " WHERE city = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
NewMsg.CmbRegion.Clear
 'Set rs = db.OpenRecordset("Category")
    Do While Not rs1.EOF
        NewMsg.CmbRegion.AddItem rs1("Region") & ""
       rs1.MoveNext
    Loop
    End If
     Add.searshCom
     NewMsg.Text1.Text = ""
End Sub

Private Sub cmbCompanyname_Click()
NewMsg.Cmbtype.Text = ""
NewMsg.CmbCategoryname.Text = ""
NewMsg.CmbProduct.Clear
NewMsg.cmbCity.Clear
NewMsg.CmbRegion.Clear
NewMsg.Cmbcountry.Text = ""
NewMsg.Lstcompany.Clear
NewMsg.Lstpersonalname.Clear
NewMsg.Lstcompany.AddItem (NewMsg.cmbCompanyname.Text)
y = 0
NewMsg.Lstcompany.Selected(y) = True
NewMsg.Text1.Text = ""
End Sub

Private Sub Cmbcountry_Click()
If NewMsg.Cmbcountry = "" Then
NewMsg.cmbCity.Clear
NewMsg.CmbRegion.Clear
Else

Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
SqLst1 = "Select distinct city From QCity"
SqLst1 = SqLst1 & " WHERE country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
NewMsg.cmbCity.Clear
 'Set rs = db.OpenRecordset("Category")
    Do While Not rs1.EOF
        NewMsg.cmbCity.AddItem rs1("city")
       rs1.MoveNext
    Loop
    NewMsg.CmbRegion.Clear
    End If
     Add.searshCom
     NewMsg.Text1.Text = ""
End Sub

Private Sub CmbProduct_Click()
'NewMsg.Cmbtype.Text = ""
'Set ws = CreateWorkspace("", "admin", "")
'Set db = ws.OpenDatabase(PathPro)
'SqLst1 = "Select Companyname From Qprocomp"
'SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text & "'"
'Set rs1 = db.OpenRecordset(SqLst1)
'NewMsg.cmbCompanyname.Clear
'NewMsg.Lstcompany.Clear
'y = 0
'NewMsg.cmbCompanyname.AddItem "All"
'    Do While Not rs1.EOF
'       NewMsg.cmbCompanyname.AddItem rs1("Companyname")
''       NewMsg.Lstcompany.AddItem rs1("Companyname")
 '      NewMsg.Lstcompany.Selected(y) = True
 '      rs1.MoveNext
 'y = y + 1
 '      Loop
 Add.searshCom
 NewMsg.Text1.Text = ""
End Sub

Private Sub Cmbtype_Click()
If NewMsg.Cmbtype.Text <> "" Then
Add.searshCom
Else
End If
NewMsg.Text1.Text = ""
End Sub



Private Sub cmdPrint_Click()

On Error GoTo eh
If NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\company.rpt")
frmMain.CrystalReport1.SelectionFormula = ""
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'1===========================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\onecom.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcompany.Companyname}=" & _
"'" & NewMsg.cmbCompanyname.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'2===========================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\company.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcompany.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'3===========================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\company.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcompany.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & " And {Qcompany.City}=" & _
"'" & NewMsg.cmbCity.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'4===========================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\company.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcompany.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & " And {Qcompany.City}=" & _
"'" & NewMsg.cmbCity.Text & "'" & " And {Qcompany.Region}=" & _
"'" & NewMsg.CmbRegion.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'5===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text <> "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\Qcatype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatype.Categoryname}=" & _
"'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'6===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text <> "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\Qprodtype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprodtype.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'" & " And {Qprodtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'7===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\Qcatype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatype.Categoryname}=" & _
"'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qcatype.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'8===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\Qprodtype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprodtype.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'" & " And {Qprodtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qprodtype.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'9===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\Qcatype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatype.Categoryname}=" & _
"'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qcatype.Qcompany.City}=" & _
"'" & NewMsg.cmbCity.Text & "'" & " And {Qcatype.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'10===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\Qprodtype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprodtype.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'" & " And {Qprodtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qprodtype.Qcompany.City}=" & _
"'" & NewMsg.cmbCity.Text & "'" & " And {Qprodtype.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'11===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\Qcatype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatype.Categoryname}=" & _
"'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qcatype.Qcompany.City}=" & _
"'" & NewMsg.cmbCity.Text & "'" & " And {Qcatype.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & " And {Qcatype.Region}=" & _
"'" & NewMsg.CmbRegion.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'12===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\Qprodtype.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprodtype.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'" & " And {Qprodtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qprodtype.Qcompany.City}=" & _
"'" & NewMsg.cmbCity.Text & "'" & " And {Qprodtype.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'" & " And {Qprodtype.Region}=" & _
"'" & NewMsg.CmbRegion.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'13===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\repcompany.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Categoryname}=" & _
"'" & NewMsg.CmbCategoryname.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'14============================================
'ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text = "" And _
'    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'    frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
'    frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Categoryname}=" & _
'    "'" & NewMsg.CmbCategoryname.Text & "'"
'    frmMain.CrystalReport1.Destination = 0
'
'    frmMain.CrystalReport1.WindowState = crptMaximized
'    frmMain.CrystalReport1.DataFiles(0) = (PathPro)
'    frmMain.CrystalReport1.Action = 1
'   Exit Sub
'=============================================
'ElseIf NewMsg.CmbCategoryname.Text = "All" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
'    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'    frmMain.CrystalReport1.ReportFileName = (App.Path & "\repcompany.rpt")
'    frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Companyname}=" & _
'    "'" & NewMsg.cmbCompanyname.Text & "'"
'    frmMain.CrystalReport1.Destination = 0
'    frmMain.CrystalReport1.WindowState = crptMaximized
'    frmMain.CrystalReport1.DataFiles(0) = (PathPro)
'    frmMain.CrystalReport1.Action = 1
'      Exit Sub
'    '============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text = "" And _
'    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'    frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
'    frmMain.CrystalReport1.SelectionFormula = ""
'    frmMain.CrystalReport1.Destination = 0
'
'    frmMain.CrystalReport1.WindowState = crptMaximized
'    frmMain.CrystalReport1.DataFiles(0) = (PathPro)
'    frmMain.CrystalReport1.Action = 1
'   Exit Sub

'===========================================
ElseIf NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'15=============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text <> "" And _
'    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'    frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
'    frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Companyname}=" & _
'    "'" & NewMsg.cmbCompanyname.Text & "'"
'    frmMain.CrystalReport1.Destination = 0
'    frmMain.CrystalReport1.WindowState = crptMaximized
'    frmMain.CrystalReport1.DataFiles(0) = (PathPro)
'    frmMain.CrystalReport1.Action = 1
'      Exit Sub
'   '============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
'    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "All" Then
'    frmMain.CrystalReport1.ReportFileName = (App.Path & "\type.rpt")
'    frmMain.CrystalReport1.SelectionFormula = ""
'    frmMain.CrystalReport1.Destination = 0
'
'    frmMain.CrystalReport1.WindowState = crptMaximized
'    frmMain.CrystalReport1.DataFiles(0) = (PathPro)
''    frmMain.CrystalReport1.Action = 1
'   Exit Sub
'============================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbtype.Text <> "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\type.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'16=============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
''    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text = "" And NewMsg.CmbRegion.Text = "" _
'    And NewMsg.Cmbtype.Text = "All" Then
'    frmMain.CrystalReport1.ReportFileName = (App.Path & "\type.rpt")
'    frmMain.CrystalReport1.SelectionFormula = "{Qtype.Companyname}=" & _
'    "'" & NewMsg.cmbCompanyname.Text & "'"
'    frmMain.CrystalReport1.Destination = 0
'    frmMain.CrystalReport1.WindowState = crptMaximized
'    frmMain.CrystalReport1.DataFiles(0) = (PathPro)
'    frmMain.CrystalReport1.Action = 1
'      Exit Sub
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
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'17====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\repcompany.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Categoryname}=" & _
"'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatcomp.city}=" & _
"'" & NewMsg.cmbCity.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
''18====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\repcompany.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Categoryname}=" & _
"'" & NewMsg.CmbCategoryname.Text & "'" & " And {Qcatcomp.Region}=" & _
"'" & NewMsg.CmbRegion.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'19====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'" & " And {Qprocomp.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'20====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'" & " And {Qprocomp.city}=" & _
"'" & NewMsg.cmbCity.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
''21====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Productname}=" & _
"'" & NewMsg.CmbProduct.Text & "'" & " And {Qprocomp.Region}=" & _
"'" & NewMsg.CmbRegion.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'22====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\type.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qtype.country}=" & _
"'" & NewMsg.Cmbcountry.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'23====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text = "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\type.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qtype.city}=" & _
"'" & NewMsg.cmbCity.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
''24====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text <> "" Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\type.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qtype.Typename}=" & _
"'" & NewMsg.Cmbtype.Text & "'" & " And {Qtype.Region}=" & _
"'" & NewMsg.CmbRegion.Text & "'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'25=======================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(0).Value = True Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\repcompany.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcatcomp.Categoryname}like" & _
"'*" & NewMsg.Text1.Text & "*'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'26=======================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(1).Value = True Then

frmMain.CrystalReport1.ReportFileName = (App.Path & "\company.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qcompany.Companyname}like" & _
"'*" & NewMsg.Text1.Text & "*'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'27=======================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(2).Value = True Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\reppcompany.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qprocomp.Productname}like" & _
"'*" & NewMsg.Text1.Text & "*'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'28=======================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(3).Value = True Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\type.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qtype.Typename}like" & _
"'*" & NewMsg.Text1.Text & "*'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
Exit Sub
'=====================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(5).Value = True Then
frmMain.CrystalReport1.ReportFileName = (App.Path & "\person.rpt")
frmMain.CrystalReport1.SelectionFormula = "{Qpersonal.name}like" & _
"'*" & NewMsg.Text1.Text & "*'"
frmMain.CrystalReport1.Destination = 0
frmMain.CrystalReport1.WindowState = crptMaximized
frmMain.CrystalReport1.DataFiles(0) = (PathPro)
frmMain.CrystalReport1.Action = 1
'===================================================
End If
eh:

 If Err.Number = 20526 Then MsgBox ("Please Add printer")
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
frmMain.Toolbar1.Buttons(5).Enabled = False
End Sub

Private Sub Lstcompany_ItemCheck(Item As Integer)
If NewMsg.Chkcompany.Value = False Then
NewMsg.Clearcompany.Value = True
Else
NewMsg.Clearcompany.Value = False
End If
NewMsg.cmbpostype.Text = ""
'On Error Resume Next
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
If NewMsg.Lstcompany.Selected(NewMsg.Lstcompany.ListIndex) = True Then

If NewMsg.Lstcompany.List(NewMsg.Lstcompany.ListIndex) <> "-1" Then

SqLst1 = "Select name From Qpersonal"
SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.Lstcompany.List(NewMsg.Lstcompany.ListIndex) & "'"
Set rs1 = db.OpenRecordset(SqLst1, dbOpenDynaset)
' frmCampany.LstProduct.Clear
 'Set rs = db.OpenRecordset("Category")

    Do While Not rs1.EOF
        'NewMsg.LstPersonaltype.AddItem rs1("posTypename")
        NewMsg.Lstpersonalname.AddItem rs1("name")
        rs1.MoveNext
    Loop
    y = NewMsg.Lstpersonalname.ListCount
Do While Not y <= 0
y = y - 1
NewMsg.Lstpersonalname.Selected(y) = True
Loop

    End If
    
  Else
  removeitme
     
 
  End If
  y = 0
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

Private Sub Option2_Click(Index As Integer)
If Option2(2).Value = True Then
Option2(5).Enabled = True
Option2(4).Enabled = True
Option2(6).Enabled = True
Else
Option2(5).Enabled = False
Option2(4).Enabled = False
Option2(6).Enabled = False
End If
End Sub

Private Sub Refresh_Click()
Form_Activate
NewMsg.CmbProduct.Clear
NewMsg.CmbRegion.Clear
NewMsg.Cmbtype.Text = ""
NewMsg.Text1.Text = ""
NewMsg.cmbCity.Clear
NewMsg.Chkcompany.Value = True
frmMain.Show
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

Private Sub Send_Click()
If NewMsg.txtsubject.Text = "" Or NewMsg.txtNoteText.Text = "" Then
MsgBox ("Please...Complete Data")
Else
NewMsg.Frame6.Visible = True
' Place the Subject and Note text into the buffer.
    ' Add room in the beginning for attachment files.
   ' frmMain.Label2.Caption = "Waite..."
    
   
   End If
End Sub


Public Sub sendmail()
'If NewMsg.txtTo.Text = "" Or NewMsg.txtTo.Text = "" Or NewMsg.txtsubject.Text = "" Or NewMsg.txtNoteText.Text = "" Then Exit Sub
' Place the Subject and Note text into the buffer.
    ' Add room in the beginning for attachment files.
   
    ' NewMsg.sendM
End Sub

Public Sub searshm()


End Sub


Private Sub txtNoteText_KeyPress(KeyAscii As Integer)
NewMsg.optfaxem.Enabled = True
NewMsg.Optltrper.Enabled = True
End Sub

Private Sub txtNoteText_LostFocus()
If NewMsg.txtNoteText.Text = "" Then
NewMsg.Option2(1).Enabled = False
NewMsg.Option2(7).Enabled = False
Else
NewMsg.Option2(1).Enabled = True
NewMsg.Option2(7).Enabled = True
End If
End Sub

Public Sub removeitme()
SqLst1 = "Select name From Qpersonal"
SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.Lstcompany.List(NewMsg.Lstcompany.ListIndex) & "'"
Set rs1 = db.OpenRecordset(SqLst1, dbOpenDynaset)
' frmCampany.LstProduct.Clear
 'Set rs = db.OpenRecordset("Category")
 A = 0
e:
 Do While Not rs1.EOF
 A = -1
 Do While A < NewMsg.Lstpersonalname.ListCount
NewMsg.Lstpersonalname.ListIndex = A
    B = rs1("name")
  If B = NewMsg.Lstpersonalname.List(NewMsg.Lstpersonalname.ListIndex) Then
        NewMsg.Lstpersonalname.RemoveItem (NewMsg.Lstpersonalname.ListIndex)
       rs1.MoveNext
         GoTo e

    End If


        A = A + 1
    
  Loop
   rs1.MoveNext

  Loop
End Sub
