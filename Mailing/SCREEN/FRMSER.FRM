VERSION 5.00
Begin VB.Form frmser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Please Call IMex.net "
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttmpname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtCHN 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   5295
   End
   Begin VB.TextBox txtHDF 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "TEL:  02 2186321"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1170
      TabIndex        =   6
      Top             =   1080
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Please Call IMEX .NET and take Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   4230
   End
End
Attribute VB_Name = "frmser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()
End
End Sub

Private Sub cmdOK_Click()
Dim s As String
If frmser.txtCHN.Text = frmser.txttmpname.Text Then
Open "numpe" For Output As 257
'
  Print #257, frmser.txttmpname.Text
  Print #257,
'rr:
Close 257#
frmMain.Show
Unload Me
Set frmser = Nothing
Else
MsgBox ("this's Number is Mistake")
End
End If
End Sub

Public Function chk(y As String, tmpname As String)

frmser.txtHDF.Text = y
frmser.txttmpname.Text = tmpname
End Function


