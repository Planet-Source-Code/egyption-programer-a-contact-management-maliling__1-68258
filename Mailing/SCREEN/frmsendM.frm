VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmsendM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send Mail"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   5385
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10275
      Begin VB.TextBox TXT_ATT 
         Height          =   345
         Left            =   1905
         TabIndex        =   14
         Top             =   2205
         Width           =   4215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6360
         TabIndex        =   13
         Top             =   2205
         Width           =   735
      End
      Begin MSComDlg.CommonDialog CMDialog1 
         Left            =   240
         Top             =   4920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox TXT_BODY 
         Height          =   1380
         Left            =   1905
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   3360
         Width           =   8040
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SEND"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   10
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox TXT_SUB 
         Height          =   345
         Left            =   1905
         TabIndex        =   9
         Top             =   2880
         Width           =   4215
      End
      Begin VB.TextBox TXT_FROM 
         Height          =   345
         Left            =   1905
         TabIndex        =   7
         Top             =   825
         Width           =   4215
      End
      Begin VB.TextBox TXT_CC 
         Height          =   345
         Left            =   1905
         TabIndex        =   5
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox TXT_TO 
         Height          =   345
         Left            =   1905
         TabIndex        =   4
         Top             =   210
         Width           =   4215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "BODY:"
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
         Left            =   1155
         TabIndex        =   12
         Top             =   3405
         Width           =   570
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SUBJECT:"
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
         Left            =   840
         TabIndex        =   8
         Top             =   2895
         Width           =   885
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "FROM:"
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
         Left            =   1140
         TabIndex        =   6
         Top             =   885
         Width           =   585
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ATTATCHMENT:"
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
         Left            =   285
         TabIndex        =   3
         Top             =   2235
         Width           =   1440
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "CC:"
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
         Left            =   1425
         TabIndex        =   2
         Top             =   1605
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TO:"
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
         Left            =   1410
         TabIndex        =   1
         Top             =   240
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmsendM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim myMail
Set myMail = CreateObject("CDONTS.NewMail")
myMail.From = TXT_FROM.Text
myMail.To = TXT_TO.Text
myMail.Subject = TXT_SUB.Text
myMail.Body = TXT_BODY.Text
myMail.AttachFile Form1.TXT_ATT.Text
myMail.Send
Set myMail = Nothing
End Sub

Private Sub Command2_Click()

With CMDialog1
.DialogTitle = "select atatchement"
.Filter = "all files(*.*)|*.*"
.FilterIndex = 1
.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
.ShowOpen
If Err = cdlCancel Then Exit Sub
TXT_ATT.Text = .FileName
End With

End Sub

