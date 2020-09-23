VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form DIALER 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1935
   ClientLeft      =   3990
   ClientTop       =   3255
   ClientWidth     =   4275
   Icon            =   "Dialer.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1935
   ScaleWidth      =   4275
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   -15
      Picture         =   "Dialer.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   4275
      TabIndex        =   4
      Top             =   -15
      Width           =   4305
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Dialer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   135
         TabIndex        =   5
         Top             =   60
         Width           =   2520
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   0
      Top             =   945
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   1695
      TabIndex        =   3
      Top             =   1185
      Width           =   852
   End
   Begin VB.CommandButton QuitButton 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2655
      TabIndex        =   1
      Top             =   1185
      Width           =   852
   End
   Begin VB.CommandButton DialButton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dial"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   735
      TabIndex        =   0
      Top             =   1185
      Width           =   852
   End
   Begin VB.Label Status 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "To dial a number, click the Dial button"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   750
      TabIndex        =   2
      Top             =   645
      Width           =   2775
   End
End
Attribute VB_Name = "DIALER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z
Dim CancelFlag
Private Sub CancelButton_Click()
CancelFlag = True
CancelButton.Enabled = False
End Sub
Private Sub Dial(Number$)
Dim DialString$, FromModem$, dummy
DialString$ = "ATDT" + Number$ + ";" + vbCr
MSComm1.CommPort = 2
MSComm1.Settings = "9600,N,8,1"
'On Error Resume Next
MSComm1.PortOpen = True
If Err Then
MsgBox "COM2: not available. Change the CommPort property to another port."
Exit Sub
End If
MSComm1.InBufferCount = 0
MSComm1.Output = DialString$
Do
dummy = DoEvents()
If MSComm1.InBufferCount Then
FromModem$ = FromModem$ + MSComm1.Input
If InStr(FromModem$, "OK") Then
Beep
MsgBox "Please pick up the phone and either press Enter or click OK"
Exit Do
End If
End If
If CancelFlag Then
CancelFlag = False
Exit Do
End If
Loop
MSComm1.Output = "ATH" + vbCr
MSComm1.PortOpen = False
End Sub
Private Sub DialButton_Click()
Dim Number$, Temp$
If frmCampany.txttel1.Text <> "" Then
DialButton.Enabled = False
QuitButton.Enabled = False
CancelButton.Enabled = True
Number$ = frmCampany.txttel1.Text 'InputBox$("Enter phone number:", Number$)
If Number$ = "" Then Exit Sub
Temp$ = Status
Status = "Dialing - " + Number$
Dial Number$
DialButton.Enabled = True
QuitButton.Enabled = True
CancelButton.Enabled = False
Status = Temp$
Else
Unload Me
End If
End Sub
Private Sub Form_Load()
MSComm1.InputLen = 0
Dim Index As Integer
frmMain.Arrange Index
frmMain.hideb
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.hideb

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Add.Mode
End Sub

Private Sub QuitButton_Click()
Unload Me
End Sub

