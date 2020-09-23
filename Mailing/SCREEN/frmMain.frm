VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   Caption         =   "Mailing"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   HelpContextID   =   1
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0442
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Add"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Save"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Delete"
            Object.Tag             =   ""
            ImageIndex      =   24
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Search"
            Object.Tag             =   ""
            ImageIndex      =   25
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Exit"
            Object.Tag             =   ""
            ImageIndex      =   26
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Dial-up"
            Object.ToolTipText     =   "Dial-up"
            Object.Tag             =   ""
            ImageIndex      =   27
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   11880
      TabIndex        =   3
      Top             =   7635
      Width           =   11880
      Begin VB.Label Label2 
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   6045
         TabIndex        =   7
         Top             =   90
         Width           =   1170
      End
      Begin VB.Line MsgBoxSide 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   7260
         X2              =   7260
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line MsgBoxSide 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   60
         X2              =   60
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line MsgBoxLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   60
         X2              =   7260
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line MsgBoxLine 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   60
         X2              =   7260
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Line TimeBoxSide 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   8580
         X2              =   8580
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line TimeBoxLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   7320
         X2              =   8580
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Line TimeBoxSide 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   7320
         X2              =   7320
         Y1              =   60
         Y2              =   300
      End
      Begin VB.Line TimeBoxLine 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   7320
         X2              =   8580
         Y1              =   60
         Y2              =   60
      End
      Begin VB.Line TopLine2 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   10800
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label UnreadLbl 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   60
         Width           =   1575
      End
      Begin VB.Line TopLine2 
         BorderColor     =   &H00000000&
         Index           =   0
         X1              =   0
         X2              =   10800
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label TimeLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         Height          =   205
         Left            =   7500
         TabIndex        =   5
         Top             =   75
         Width           =   345
      End
      Begin VB.Label MsgCountLbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message Count Information"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   75
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   11850
      TabIndex        =   1
      Top             =   975
      Visible         =   0   'False
      Width           =   11880
      Begin VB.Timer Timer1 
         Interval        =   15000
         Left            =   180
         Top             =   120
      End
      Begin MSMAPI.MAPIMessages MapiMess 
         Left            =   1320
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         AddressEditFieldCount=   0
         AddressModifiable=   0   'False
         AddressResolveUI=   0   'False
         FetchSorted     =   0   'False
         FetchUnreadOnly =   -1  'True
      End
      Begin MSMAPI.MAPISession MapiSess 
         Left            =   720
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DownloadMail    =   -1  'True
         LogonUI         =   -1  'True
         NewSession      =   -1  'True
      End
      Begin MSComDlg.CommonDialog CMDialog1 
         Left            =   2010
         Top             =   195
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         FilterIndex     =   672
         FontSize        =   2.36135e-37
      End
      Begin VB.Label Label1 
         Caption         =   "These controls are invisible at run time."
         Height          =   315
         Left            =   2700
         TabIndex        =   2
         Top             =   300
         Width           =   2835
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmMain.frx":22BAB
      ScaleHeight     =   555
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   420
      Width           =   11880
      Begin VB.Image I11ENTER 
         Height          =   435
         Left            =   10875
         Picture         =   "frmMain.frx":266E1
         Top             =   120
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Image I10ENTER 
         Height          =   435
         Left            =   9885
         Picture         =   "frmMain.frx":277C0
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image I9ENTER 
         Height          =   435
         Left            =   8880
         Picture         =   "frmMain.frx":289C3
         Top             =   120
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image I8ENTER 
         Height          =   435
         Left            =   7815
         Picture         =   "frmMain.frx":29C1E
         Top             =   120
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image I7ENTER 
         Height          =   435
         Left            =   6885
         Picture         =   "frmMain.frx":2AEA2
         Top             =   120
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Image I7OFF 
         Height          =   435
         Left            =   6885
         Picture         =   "frmMain.frx":2C18A
         Top             =   120
         Width           =   945
      End
      Begin VB.Image I7ON 
         Height          =   435
         Left            =   6885
         Picture         =   "frmMain.frx":2D494
         Top             =   120
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Image I6ENTER 
         Height          =   435
         Left            =   5850
         Picture         =   "frmMain.frx":2E8DE
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image I5ENTER 
         Height          =   435
         Left            =   4845
         Picture         =   "frmMain.frx":2FB76
         Top             =   120
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Image I4OFF 
         Height          =   435
         Left            =   3600
         Picture         =   "frmMain.frx":30D43
         Top             =   120
         Width           =   1095
      End
      Begin VB.Image I4ENTER 
         Height          =   435
         Left            =   3600
         Picture         =   "frmMain.frx":321CA
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image I4ON 
         Height          =   435
         Left            =   3600
         Picture         =   "frmMain.frx":335F5
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image I3OFF 
         Height          =   435
         Left            =   2400
         Picture         =   "frmMain.frx":34BBF
         Top             =   120
         Width           =   1095
      End
      Begin VB.Image I3ON 
         Height          =   435
         Left            =   2400
         Picture         =   "frmMain.frx":35E0F
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image I3ENTER 
         Height          =   435
         Left            =   2400
         Picture         =   "frmMain.frx":370E8
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image I2ON 
         Height          =   435
         Left            =   1275
         Picture         =   "frmMain.frx":3837E
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image I2ENTER 
         Height          =   435
         Left            =   1275
         Picture         =   "frmMain.frx":397F6
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image I2OFF 
         Height          =   435
         Left            =   1275
         Picture         =   "frmMain.frx":3AB19
         Top             =   120
         Width           =   1095
      End
      Begin VB.Image I1ON 
         Height          =   435
         Left            =   180
         Picture         =   "frmMain.frx":3BE7B
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image I1ENTER 
         Height          =   435
         Left            =   180
         Picture         =   "frmMain.frx":3D36E
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Image I1off 
         Height          =   435
         Left            =   180
         Picture         =   "frmMain.frx":3E75B
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   540
      End
      Begin VB.Image I9ON 
         Height          =   435
         Left            =   8880
         Picture         =   "frmMain.frx":3FB45
         Top             =   120
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image I9OFF 
         Height          =   435
         Left            =   8880
         Picture         =   "frmMain.frx":40DF5
         Top             =   120
         Width           =   960
      End
      Begin VB.Image I8ON 
         Height          =   435
         Left            =   7815
         Picture         =   "frmMain.frx":42012
         Top             =   120
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image I8OFF 
         Height          =   435
         Left            =   7815
         Picture         =   "frmMain.frx":432EE
         Top             =   120
         Width           =   960
      End
      Begin VB.Image I6ON 
         Height          =   435
         Left            =   5850
         Picture         =   "frmMain.frx":444D6
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image I6OFF 
         Height          =   435
         Left            =   5850
         Picture         =   "frmMain.frx":457FD
         Top             =   120
         Width           =   975
      End
      Begin VB.Image I5ON 
         Height          =   435
         Left            =   4845
         Picture         =   "frmMain.frx":46A1C
         Top             =   120
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Image I5OFF 
         Height          =   435
         Left            =   4845
         Picture         =   "frmMain.frx":47D20
         Top             =   120
         Width           =   930
      End
      Begin VB.Image I10OFF 
         Height          =   435
         Left            =   9885
         Picture         =   "frmMain.frx":48F3C
         Top             =   120
         Width           =   975
      End
      Begin VB.Image I10ON 
         Height          =   435
         Left            =   9885
         Picture         =   "frmMain.frx":4A173
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image I11ON 
         Height          =   435
         Left            =   10875
         Picture         =   "frmMain.frx":4B486
         Top             =   120
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Image I11OFF 
         Height          =   435
         Left            =   10875
         Picture         =   "frmMain.frx":4C632
         Top             =   120
         Width           =   945
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5160
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3015
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483648
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   29
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4D74B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4D845
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4D957
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4DFE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4E81B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4F04D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4F1E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4F381
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4F51B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4F6B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4F84F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4F9E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4FB83
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4FC7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4FD77
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":500C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":502FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":511D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":513FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":51A8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":526FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":52AAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":52DC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":52ED9
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":53073
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":53185
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5349F
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":537B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":538CB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Menu001 
      Caption         =   "&File"
      Begin VB.Menu Menu001001 
         Caption         =   "&Main Data"
         Shortcut        =   ^M
      End
      Begin VB.Menu Menu001004 
         Caption         =   "&Companies"
         Shortcut        =   ^C
      End
      Begin VB.Menu Menu00003 
         Caption         =   "Perso&ns"
         Shortcut        =   ^N
      End
      Begin VB.Menu Menu004 
         Caption         =   "&Report"
         Shortcut        =   ^R
      End
      Begin VB.Menu k9 
         Caption         =   "-"
      End
      Begin VB.Menu PrintMessage 
         Caption         =   "&Print Message"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu PrSetup 
         Caption         =   "Prin&ter Setup..."
         Shortcut        =   ^T
      End
      Begin VB.Menu y 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu Menu003 
      Caption         =   "&Edit"
      Begin VB.Menu Menu003001 
         Caption         =   "Add"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu Menu003002 
         Caption         =   "Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu Menu003003 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu set4 
         Caption         =   "-"
      End
      Begin VB.Menu EditDelete 
         Caption         =   "Delete Messa&ge"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu MF 
         Caption         =   "Move First"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mn 
         Caption         =   "Move Next"
         Shortcut        =   {F6}
      End
      Begin VB.Menu MP 
         Caption         =   "Move Prev"
         Shortcut        =   {F7}
      End
      Begin VB.Menu ML 
         Caption         =   "Move Last"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu Mail 
      Caption         =   "&Mail"
      Begin VB.Menu Logon 
         Caption         =   "Lo&gon"
      End
      Begin VB.Menu LogOff 
         Caption         =   "Log&off"
         Enabled         =   0   'False
      End
      Begin VB.Menu uy 
         Caption         =   "-"
      End
      Begin VB.Menu rMsgList 
         Caption         =   "Update Message List"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Options 
      Caption         =   "&Options"
      Begin VB.Menu MailOpts 
         Caption         =   "&Mail..."
      End
      Begin VB.Menu FontS 
         Caption         =   "&Fonts"
         Begin VB.Menu FontScreen 
            Caption         =   "&Screen..."
         End
         Begin VB.Menu FontPrt 
            Caption         =   "&Printer..."
         End
      End
      Begin VB.Menu DispTools 
         Caption         =   "&Display Tools"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Address 
      Caption         =   "&Address"
      Begin VB.Menu ShowAB 
         Caption         =   "Show Address Book"
      End
   End
   Begin VB.Menu Window 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu wa 
         Caption         =   "&Cascade"
         Index           =   0
      End
      Begin VB.Menu wa 
         Caption         =   "Tile Horizontally"
         Index           =   1
      End
      Begin VB.Menu wa 
         Caption         =   "Tile Vertically"
         Index           =   2
      End
      Begin VB.Menu wa 
         Caption         =   "Arrange Icons"
         Index           =   3
      End
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "&About"
      Begin VB.Menu About 
         Caption         =   " &About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Index As Integer
Dim alarm As Integer
Public PathPro As String
'Public RepCaptin As String


Private Sub cmdmaindata_Click()
frmMainData.Show
End Sub




Private Sub I10ENTER_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
If frmMain.Logon.Enabled = False Then
    ' View the previous message in the list.
    If MailLst.MList.ListIndex <> 0 Then
        MailLst.MList.ItemData(MailLst.MList.ListIndex) = False
        MailLst.MList.ListIndex = MailLst.MList.ListIndex - 1
    End If
    Call ViewNextMsg
    End If
End Sub

Private Sub I11ENTER_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
If frmMain.Logon.Enabled = False Then

     ' View the next message in the list.
    If MailLst.MList.ListIndex <> MailLst.MList.ListCount - 1 Then
        MailLst.MList.ItemData(MailLst.MList.ListIndex) = False
        MailLst.MList.ListIndex = MailLst.MList.ListIndex + 1
    End If
    Call ViewNextMsg
End If
End Sub

Private Sub I1ENTER_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
frmMainData.Show
'Call hideb
End Sub


Private Sub i1OFF_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
I1ON.Visible = True
I1off.Visible = False

'**


I2ENTER.Visible = False
I2ON.Visible = False
I2OFF.Visible = True




I3ENTER.Visible = False
I3ON.Visible = False
I3OFF.Visible = True


I4ENTER.Visible = False
I4ON.Visible = False
I4OFF.Visible = True


I5ENTER.Visible = False
I5ON.Visible = False
I5OFF.Visible = True


I6ENTER.Visible = False
I6ON.Visible = False
I6OFF.Visible = True


I7ENTER.Visible = False
I7ON.Visible = False
I7OFF.Visible = True


I8ENTER.Visible = False
I8ON.Visible = False
I8OFF.Visible = True


I9ENTER.Visible = False
I9ON.Visible = False
I9OFF.Visible = True


I10ENTER.Visible = False
I10ON.Visible = False
I10OFF.Visible = True


I11ENTER.Visible = False
I11ON.Visible = False
I11OFF.Visible = True

End Sub

Private Sub i1ON_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
I1ENTER.Visible = True
I1ON.Visible = False
End Sub


Private Sub I2ENTER_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
frmCampany.Show

End Sub

Private Sub I2ON_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
I2ENTER.Visible = True
I2ON.Visible = False

End Sub

Private Sub I2OFF_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
I2ON.Visible = True
I2OFF.Visible = False

'**
I1ENTER.Visible = False
I1ON.Visible = False
I1off.Visible = True


I3ENTER.Visible = False
I3ON.Visible = False
I3OFF.Visible = True


I4ENTER.Visible = False
I4ON.Visible = False
I4OFF.Visible = True


I5ENTER.Visible = False
I5ON.Visible = False
I5OFF.Visible = True


I6ENTER.Visible = False
I6ON.Visible = False
I6OFF.Visible = True


I7ENTER.Visible = False
I7ON.Visible = False
I7OFF.Visible = True


I8ENTER.Visible = False
I8ON.Visible = False
I8OFF.Visible = True


I9ENTER.Visible = False
I9ON.Visible = False
I9OFF.Visible = True


I10ENTER.Visible = False
I10ON.Visible = False
I10OFF.Visible = True


I11ENTER.Visible = False
I11ON.Visible = False
I11OFF.Visible = True

End Sub

Private Sub I3ENTER_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
frmpersonal.Show
Call hideb
End Sub


Private Sub I4ENTER_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
NewMsg.Show
End Sub



Private Sub I5ENTER_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

Dim NewMessage As New NewMsg
    On Error Resume Next
If frmMain.Logon.Enabled = False Then

     Index = 6 ': Compose New Messa0ge
    '       = 7: Reply
    '       = 8: Reply All
    '       = 9: Forward

    ' Save the header information and current note text.
    If Index > 6 Then
        ' SVNote = GetHeader(frmMain.MapiMess) + frmMain.MapiMess.MsgNoteText
        SVNote = frmMain.MapiMess.MsgNoteText
        SVNote = GetHeader(frmMain.MapiMess) + SVNote
    End If

    frmMain.MapiMess.Action = Index

    ' Set the new message text.
    If Index > 6 Then
        frmMain.MapiMess.MsgNoteText = SVNote
        
    End If

    If SendWithMapi Then
        frmMain.MapiMess.Action = vbMessageSendDlg
       ' frmMain.MapiMess.AddressLabel = "ggg"
    Else
        Call LoadMessage(-1, NewMessage)            ' Load message into frmMain NewMSG window.
    End If
    End If
End Sub

Private Sub I6ENTER_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim NewMessage As New NewMsg
If frmMain.Logon.Enabled = False Then
    On Error Resume Next

    Index = 7 ': Compose New Messa0ge
    
    ' Save the header information and current note text.
    If Index > 6 Then
        ' SVNote = GetHeader(frmMain.MapiMess) + frmMain.MapiMess.MsgNoteText
        SVNote = frmMain.MapiMess.MsgNoteText
        SVNote = GetHeader(frmMain.MapiMess) + SVNote
    End If

    frmMain.MapiMess.Action = Index

    ' Set the new message text.
    If Index > 6 Then
        frmMain.MapiMess.MsgNoteText = SVNote
        
    End If

    If SendWithMapi Then
        frmMain.MapiMess.Action = vbMessageSendDlg
       ' frmMain.MapiMess.AddressLabel = "ggg"
    Else
        Call LoadMessage(-1, NewMessage)            ' Load message into frmMain NewMSG window.
    End If
    End If
End Sub

Private Sub I7ENTER_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim NewMessage As New NewMsg

On Error Resume Next
If frmMain.Logon.Enabled = False Then
     Index = 8 ': Compose New Messa0ge
    '       = 7: Reply
    '       = 8: Reply All
    '       = 9: Forward
    If Index > 6 Then
        ' SVNote = GetHeader(frmMain.MapiMess) + frmMain.MapiMess.MsgNoteText
        SVNote = frmMain.MapiMess.MsgNoteText
        SVNote = GetHeader(frmMain.MapiMess) + SVNote
    End If

    frmMain.MapiMess.Action = Index

    ' Set the new message text.
    If Index > 6 Then
        frmMain.MapiMess.MsgNoteText = SVNote
        
    End If

    If SendWithMapi Then
        frmMain.MapiMess.Action = vbMessageSendDlg
       ' frmMain.MapiMess.AddressLabel = "ggg"
    Else
        Call LoadMessage(-1, NewMessage)            ' Load message into frmMain NewMSG window.
    End If
End If
End Sub

Private Sub I8ENTER_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim NewMessage As New NewMsg
    On Error Resume Next
If frmMain.Logon.Enabled = False Then
    
    Index = 9 ': Compose New Messa0ge
    '       = 7: Reply
    '       = 8: Reply All
    '       = 9: Forward
    ' Save the header information and current note text.
    If Index > 6 Then
        ' SVNote = GetHeader(frmMain.MapiMess) + frmMain.MapiMess.MsgNoteText
        SVNote = frmMain.MapiMess.MsgNoteText
        SVNote = GetHeader(frmMain.MapiMess) + SVNote
    End If

    frmMain.MapiMess.Action = Index

    ' Set the new message text.
    If Index > 6 Then
        frmMain.MapiMess.MsgNoteText = SVNote
        
    End If

    If SendWithMapi Then
        frmMain.MapiMess.Action = vbMessageSendDlg
       ' frmMain.MapiMess.AddressLabel = "ggg"
    Else
        Call LoadMessage(-1, NewMessage)            ' Load message into frmMain NewMSG window.
    End If
End If
End Sub

Private Sub I9ENTER_Click()
On Error Resume Next
    ' View all selected messages that are deleted.
    If TypeOf frmMain.ActiveForm Is MsgView Then
        Call DeleteMessage
    ElseIf TypeOf frmMain.ActiveForm Is MailLst Then
        ' Delete multiple selection.
        frmMain.MapiMess.MsgIndex = MailLst.MList.ListIndex
        Call DeleteMessage
    End If
End Sub

Private Sub MDIForm_Activate()
Add.ModeP
End Sub

  Private Sub About_Click()
frmAbout.Show
End Sub


Private Sub DispTools_Click()
    DispTools.Checked = Not DispTools.Checked
    MailLst.Tools.Visible = DispTools.Checked

    
    If MailLst.Tools.Visible Then
        Factor = 1
        ToolsSize% = -MailLst.Tools.Height
    Else
        Factor = -1
        ToolsSize% = 0
    End If

    Select Case MailLst.WindowState
        Case 0    ' Change the size of the form to reflect the addition or deletion of a toolbar.
            MailLst.Height = MailLst.Height + (Factor * MailLst.Tools.Height)
        Case 2    ' If maximized, adjust the size of the list box.
            MailLst.MList.Height = ScaleHeight - 90 - MailLst.MList.Top + ToolsSize%
    End Select
End Sub

Private Sub EditDelete_Click()
' Delete the items in the list.
On Error GoTo Trap
    If TypeOf frmMain.ActiveForm Is MailLst Then
        Call I9ENTER_Click
    End If
    Exit Sub

Trap:
    ' If an error occurs, there is probably no active form.
    ' Exit the Sub procedure.
    Exit Sub
End Sub

Private Sub Exit_Click()
    ' Close the application and log off.
    If MapiSess.SessionID <> 0 Then
        Call logoff_Click
    End If
    End
End Sub

Private Sub FontPrt_Click()
    ' Set the printer fonts.
    On Error Resume Next
    CMDialog1.Flags = 2
    CMDialog1.FontName = Printer.FontName
    CMDialog1.FontSize = Printer.FontSize
    CMDialog1.FontBold = Printer.FontBold
    CMDialog1.FontItalic = Printer.FontItalic
    CMDialog1.ShowFont
    If Err = 0 Then
        Printer.FontName = CMDialog1.FontName
        Printer.FontSize = CMDialog1.FontSize
        Printer.FontBold = CMDialog1.FontBold
        Printer.FontItalic = CMDialog1.FontItalic
    End If

End Sub

Private Sub FontScreen_Click()
    ' Set the screen fonts for the active control.
    On Error Resume Next
    CMDialog1.Flags = 1
    CMDialog1.FontName = frmMain.ActiveForm.ActiveControl.FontName
    CMDialog1.FontSize = frmMain.ActiveForm.ActiveControl.FontSize
    CMDialog1.FontBold = frmMain.ActiveForm.ActiveControl.FontBold
    CMDialog1.FontItalic = frmMain.ActiveForm.ActiveControl.FontItalic
    CMDialog1.ShowFont
    If Err = 0 Then
        frmMain.ActiveForm.ActiveControl.FontName = CMDialog1.FontName
        frmMain.ActiveForm.ActiveControl.FontSize = CMDialog1.FontSize
        frmMain.ActiveForm.ActiveControl.FontBold = CMDialog1.FontBold
        frmMain.ActiveForm.ActiveControl.FontItalic = CMDialog1.FontItalic
    End If
End Sub

Private Sub logoff_Click()
    ' Log off from the mail system.
    Call LogOffUser
End Sub

Private Sub Logon_Click()
    ' Log onto the mail system.
    On Error Resume Next
    MapiSess.Action = 1
    If Err <> 0 Then
        MsgBox "Logon Failure: " + Error$
    Else
        Screen.MousePointer = 11
        MapiMess.SessionID = MapiSess.SessionID
        ' Get the message count.
        GetMessageCount
        ' Load the mail list with envelope information.
        Screen.MousePointer = 11
        Call LoadList(MapiMess)
        Screen.MousePointer = 0
        ' Adjust the buttons as needed.
        Logon.Enabled = False
        LogOff.Enabled = True
        'frmMain.SendCtl(vbMessageCompose).Enabled = True
        'frmMain.SendCtl(vbMessageReplyAll).Enabled = True
        'frmMain.SendCtl(vbMessageReply).Enabled = True
        'frmMain.SendCtl(vbMessageForward).Enabled = True
        frmMain.PrintMessage.Enabled = True
        frmMain.DispTools.Enabled = True
        frmMain.rMsgList.Enabled = True
        frmMain.EditDelete.Enabled = True
       
      End If
      
End Sub

Private Sub MailOpts_Click()
    ' Display the Mail Options form.
    OptionType = conOptionGeneral
    MailOptFrm.Show 1
End Sub

Private Sub MDIForm_Initialize()
If Right(App.Path, 1) = "\" Then
            PathPro = App.Path & "Mailing.mdb"
        Else
            PathPro = App.Path & "\Mailing.mdb"
        End If
         Set ws = CreateWorkspace("", "admin", "")
        Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
'et db = OpenDatabase(PathPro, True, False, ";pwd=eit")
End Sub

Private Sub MDIForm_Load()
    ' Ensure all the controls are sized as needed.
    TimeLbl = Time$
     SendWithMapi = True
     Call Picture1_Resize
     Call Picture2_Resize
     frmMain.MsgCountLbl = "Off Line"
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
 Call hideb
End Sub

Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
frmMain.hideb

End Sub

Private Sub Menu00003_Click()
frmpersonal.Show
End Sub

Private Sub Menu001001_Click()
frmMainData.Show
End Sub

Private Sub Menu004_Click()
NewMsg.Show
End Sub


Private Sub Picture1_Resize()
Const TimeBoxStartOffset = 1200
Const TimeBoxEndOffset = 60
Const MsgBoxStartOffset = 60
Const MsgBoxEndOffset = TimeBoxStartOffset + 90

    ' Adjust the sizes of the lines and position the time label.
    TimeLbl.Left = Picture1.Width - TimeLbl.Width - 265
    TopLine2(0).X2 = Picture1.Width
    TopLine2(1).X2 = Picture1.Width

    TimeBoxLine(0).X1 = Picture1.Width - TimeBoxStartOffset
    TimeBoxLine(0).X2 = Picture1.Width - TimeBoxEndOffset

    TimeBoxLine(1).X1 = Picture1.Width - TimeBoxStartOffset
    TimeBoxLine(1).X2 = Picture1.Width - TimeBoxEndOffset

    TimeBoxSide(0).X1 = Picture1.Width - TimeBoxStartOffset
    TimeBoxSide(0).X2 = Picture1.Width - TimeBoxStartOffset

    TimeBoxSide(1).X1 = Picture1.Width - TimeBoxEndOffset
    TimeBoxSide(1).X2 = Picture1.Width - TimeBoxEndOffset

    MsgBoxLine(0).X2 = Picture1.Width - MsgBoxEndOffset
    MsgBoxLine(1).X2 = Picture1.Width - MsgBoxEndOffset

    MsgBoxSide(1).X1 = Picture1.Width - MsgBoxEndOffset
    MsgBoxSide(1).X2 = Picture1.Width - MsgBoxEndOffset

    Picture1.Refresh
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
I1ENTER.Visible = False
I1ON.Visible = False
I1off.Visible = True


End Sub

Private Sub Picture2_Resize()
    ' Adjust the positions of the lines.
    'TopLine(0).X2 = Picture2.Width
   ' TopLine(1).X2 = Picture2.Width
    Picture2.Refresh
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
I1ENTER.Visible = False
I1ON.Visible = False
I1off.Visible = True


I2ENTER.Visible = False
I2ON.Visible = False
I2OFF.Visible = True




I3ENTER.Visible = False
I3ON.Visible = False
I3OFF.Visible = True


I4ENTER.Visible = False
I4ON.Visible = False
I4OFF.Visible = True


I5ENTER.Visible = False
I5ON.Visible = False
I5OFF.Visible = True


I6ENTER.Visible = False
I6ON.Visible = False
I6OFF.Visible = True


I7ENTER.Visible = False
I7ON.Visible = False
I7OFF.Visible = True


I8ENTER.Visible = False
I8ON.Visible = False
I8OFF.Visible = True


I9ENTER.Visible = False
I9ON.Visible = False
I9OFF.Visible = True


I10ENTER.Visible = False
I10ON.Visible = False
I10OFF.Visible = True


I11ENTER.Visible = False
I11ON.Visible = False
I11OFF.Visible = True



End Sub

Private Sub Previous_Click()
End Sub

Private Sub PrintMessage_Click()
    ' Print mail.
    Call PrintMail
End Sub

Private Sub PrSetup_Click()
' Call the printer setup procedure in the common dialog control.
On Error Resume Next
    CMDialog1.Flags = &H40  ' Printer setup dialog box only.
    CMDialog1.ShowPrinter
End Sub

Private Sub rMsgList_Click()
        Screen.MousePointer = 11
        GetMessageCount
        Call LoadList(MapiMess)
        Screen.MousePointer = 0
End Sub





Private Sub ShowAB_Click()
On Error Resume Next
    ' Show the address for the current message.
    frmMain.MapiMess.Action = vbMessageShowAdBook
    If Err Then
        If Err <> 32001 Then        ' User chose Cancel.
            MsgBox "Error: " + Error$ + " occurred trying to show the Address Book"
        End If
    Else
        If TypeOf frmMain.ActiveForm Is NewMsg Then
            Call UpdateRecips(frmMain.ActiveForm)
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    ' Update the time label.
    TimeLbl = Time$
End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
frmMain.hideb
End Sub

Private Sub wa_Click(Index As Integer)
    ' Arrange the windows as selected.
    frmMain.Arrange Index
End Sub

  
    
    
Private Sub Menu001001001_Click()
frmMainData.Show
End Sub

Private Sub Menu001001002_Click()
frmMainData.Show
End Sub

Private Sub Menu001001003_Click()
frmMainData.Show
End Sub

Private Sub Menu001001004_Click()
frmMainData.Show
End Sub

Private Sub Menu001001006_Click()
frmMainData.Show
End Sub

Private Sub Menu001002_Click()

frmEmbloyees.Show
End Sub

Private Sub Menu001003_Click()
frmTasks.Show
End Sub

Private Sub Menu001004_Click()
frmCampany.Show
End Sub

Private Sub Menu001005_Click()
frmStand.Show
End Sub

Private Sub Menu001006001_Click()
FrmALLplan.Show 1
End Sub

Private Sub Menu001007_Click()
End
End Sub

Private Sub Menu003001_Click()
Set ACFRM = frmMain.ActiveForm
Add.Add_1 (ACFRM.Name)
End Sub

Private Sub Menu003002_Click()
Set ACFRM = frmMain.ActiveForm
Save.add_edit (ACFRM.Name)
End Sub

Private Sub Menu003003_Click()
Set ACFRM = frmMain.ActiveForm
MDelete.DelRec (ACFRM.Name)
End Sub

Private Sub MF_Click()
Set ACFRM = frmMain.ActiveForm
Moving.Movefrist (ACFRM.Name)
End Sub

Private Sub ML_Click()
Set ACFRM = frmMain.ActiveForm
Moving.MoveLast (ACFRM.Name)
End Sub

Private Sub mn_Click()
Set ACFRM = frmMain.ActiveForm
Moving.NextM (ACFRM.Name)
End Sub

Private Sub MP_Click()
Set ACFRM = frmMain.ActiveForm
Moving.Previous (ACFRM.Name)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Dim ACFRM As Form
Select Case Button.Index
Case "1"
Set ACFRM = frmMain.ActiveForm
Add.Add_1 (ACFRM.Name)
Case "2"
Set ACFRM = frmMain.ActiveForm
Save.add_edit (ACFRM.Name)
Case "3"
Set ACFRM = frmMain.ActiveForm
MDelete.DelRec (ACFRM.Name)
Case "4"
Set ACFRM = frmMain.ActiveForm
Add.addsearch (ACFRM.Name)
Case "5"
Unload Me.ActiveForm
Case "6"
Set ACFRM = frmMain.ActiveForm
If ACFRM.Name = "frmCampany" Then
DIALER.Show
Else
End If
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
Case 1

Case 2

Case 3
frmReport.Show


End Select

End Sub







'*/*/*/*/*/*/*/*/*/*/*/*/*/*/*///*/*/**/*/*/*/*/********************
'---------------------------------------------------------------


Private Sub i3OFF_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
I3ON.Visible = True
I3OFF.Visible = False

'**
I1ENTER.Visible = False
I1ON.Visible = False
I1off.Visible = True


I2ENTER.Visible = False
I2ON.Visible = False
I2OFF.Visible = True


I4ENTER.Visible = False
I4ON.Visible = False
I4OFF.Visible = True


I5ENTER.Visible = False
I5ON.Visible = False
I5OFF.Visible = True


I6ENTER.Visible = False
I6ON.Visible = False
I6OFF.Visible = True


I7ENTER.Visible = False
I7ON.Visible = False
I7OFF.Visible = True


I8ENTER.Visible = False
I8ON.Visible = False
I8OFF.Visible = True


I9ENTER.Visible = False
I9ON.Visible = False
I9OFF.Visible = True


I10ENTER.Visible = False
I10ON.Visible = False
I10OFF.Visible = True


I11ENTER.Visible = False
I11ON.Visible = False
I11OFF.Visible = True

End Sub

Private Sub i3ON_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
I3ENTER.Visible = True
I3ON.Visible = False
End Sub
Private Sub i4OFF_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
I4ON.Visible = True
I4OFF.Visible = False

'**
I1ENTER.Visible = False
I1ON.Visible = False
I1off.Visible = True


I2ENTER.Visible = False
I2ON.Visible = False
I2OFF.Visible = True




I3ENTER.Visible = False
I3ON.Visible = False
I3OFF.Visible = True


I5ENTER.Visible = False
I5ON.Visible = False
I5OFF.Visible = True


I6ENTER.Visible = False
I6ON.Visible = False
I6OFF.Visible = True


I7ENTER.Visible = False
I7ON.Visible = False
I7OFF.Visible = True


I8ENTER.Visible = False
I8ON.Visible = False
I8OFF.Visible = True


I9ENTER.Visible = False
I9ON.Visible = False
I9OFF.Visible = True


I10ENTER.Visible = False
I10ON.Visible = False
I10OFF.Visible = True


I11ENTER.Visible = False
I11ON.Visible = False
I11OFF.Visible = True

End Sub

Private Sub i4ON_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
I4ENTER.Visible = True
I4ON.Visible = False
End Sub
Private Sub i5OFF_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
I5ON.Visible = True
I5OFF.Visible = False

'**

I1ENTER.Visible = False
I1ON.Visible = False
I1off.Visible = True


I2ENTER.Visible = False
I2ON.Visible = False
I2OFF.Visible = True




I3ENTER.Visible = False
I3ON.Visible = False
I3OFF.Visible = True


I4ENTER.Visible = False
I4ON.Visible = False
I4OFF.Visible = True


I6ENTER.Visible = False
I6ON.Visible = False
I6OFF.Visible = True


I7ENTER.Visible = False
I7ON.Visible = False
I7OFF.Visible = True


I8ENTER.Visible = False
I8ON.Visible = False
I8OFF.Visible = True


I9ENTER.Visible = False
I9ON.Visible = False
I9OFF.Visible = True


I10ENTER.Visible = False
I10ON.Visible = False
I10OFF.Visible = True


I11ENTER.Visible = False
I11ON.Visible = False
I11OFF.Visible = True

End Sub

Private Sub i5ON_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
I5ENTER.Visible = True
I5ON.Visible = False
End Sub
Private Sub i6OFF_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
I6ON.Visible = True
I6OFF.Visible = False

'**
I1ENTER.Visible = False
I1ON.Visible = False
I1off.Visible = True


I2ENTER.Visible = False
I2ON.Visible = False
I2OFF.Visible = True




I3ENTER.Visible = False
I3ON.Visible = False
I3OFF.Visible = True


I4ENTER.Visible = False
I4ON.Visible = False
I4OFF.Visible = True


I5ENTER.Visible = False
I5ON.Visible = False
I5OFF.Visible = True



I7ENTER.Visible = False
I7ON.Visible = False
I7OFF.Visible = True


I8ENTER.Visible = False
I8ON.Visible = False
I8OFF.Visible = True


I9ENTER.Visible = False
I9ON.Visible = False
I9OFF.Visible = True


I10ENTER.Visible = False
I10ON.Visible = False
I10OFF.Visible = True


I11ENTER.Visible = False
I11ON.Visible = False
I11OFF.Visible = True

End Sub

Private Sub i6ON_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
I6ENTER.Visible = True
I6ON.Visible = False
End Sub
Private Sub i7OFF_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
I7ON.Visible = True
I7OFF.Visible = False

'**
I1ENTER.Visible = False
I1ON.Visible = False
I1off.Visible = True


I2ENTER.Visible = False
I2ON.Visible = False
I2OFF.Visible = True




I3ENTER.Visible = False
I3ON.Visible = False
I3OFF.Visible = True


I4ENTER.Visible = False
I4ON.Visible = False
I4OFF.Visible = True


I5ENTER.Visible = False
I5ON.Visible = False
I5OFF.Visible = True


I6ENTER.Visible = False
I6ON.Visible = False
I6OFF.Visible = True


I8ENTER.Visible = False
I8ON.Visible = False
I8OFF.Visible = True


I9ENTER.Visible = False
I9ON.Visible = False
I9OFF.Visible = True


I10ENTER.Visible = False
I10ON.Visible = False
I10OFF.Visible = True


I11ENTER.Visible = False
I11ON.Visible = False
I11OFF.Visible = True

End Sub

Private Sub i7ON_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
I7ENTER.Visible = True
I7ON.Visible = False
End Sub
Private Sub i8OFF_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
I8ON.Visible = True
I8OFF.Visible = False

'**
I1ENTER.Visible = False
I1ON.Visible = False
I1off.Visible = True


I2ENTER.Visible = False
I2ON.Visible = False
I2OFF.Visible = True




I3ENTER.Visible = False
I3ON.Visible = False
I3OFF.Visible = True


I4ENTER.Visible = False
I4ON.Visible = False
I4OFF.Visible = True


I5ENTER.Visible = False
I5ON.Visible = False
I5OFF.Visible = True


I6ENTER.Visible = False
I6ON.Visible = False
I6OFF.Visible = True


I7ENTER.Visible = False
I7ON.Visible = False
I7OFF.Visible = True


I9ENTER.Visible = False
I9ON.Visible = False
I9OFF.Visible = True


I10ENTER.Visible = False
I10ON.Visible = False
I10OFF.Visible = True


I11ENTER.Visible = False
I11ON.Visible = False
I11OFF.Visible = True

End Sub

Private Sub i8ON_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
I8ENTER.Visible = True
I8ON.Visible = False
End Sub
Private Sub i9OFF_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
I9ON.Visible = True
I9OFF.Visible = False

'**
I1ENTER.Visible = False
I1ON.Visible = False
I1off.Visible = True


I2ENTER.Visible = False
I2ON.Visible = False
I2OFF.Visible = True




I3ENTER.Visible = False
I3ON.Visible = False
I3OFF.Visible = True


I4ENTER.Visible = False
I4ON.Visible = False
I4OFF.Visible = True


I5ENTER.Visible = False
I5ON.Visible = False
I5OFF.Visible = True


I6ENTER.Visible = False
I6ON.Visible = False
I6OFF.Visible = True


I7ENTER.Visible = False
I7ON.Visible = False
I7OFF.Visible = True


I8ENTER.Visible = False
I8ON.Visible = False
I8OFF.Visible = True



I10ENTER.Visible = False
I10ON.Visible = False
I10OFF.Visible = True


I11ENTER.Visible = False
I11ON.Visible = False
I11OFF.Visible = True

End Sub

Private Sub i9ON_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
I9ENTER.Visible = True
I9ON.Visible = False
End Sub
Private Sub i10OFF_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
I10ON.Visible = True
I10OFF.Visible = False

'**
I1ENTER.Visible = False
I1ON.Visible = False
I1off.Visible = True


I2ENTER.Visible = False
I2ON.Visible = False
I2OFF.Visible = True




I3ENTER.Visible = False
I3ON.Visible = False
I3OFF.Visible = True


I4ENTER.Visible = False
I4ON.Visible = False
I4OFF.Visible = True


I5ENTER.Visible = False
I5ON.Visible = False
I5OFF.Visible = True


I6ENTER.Visible = False
I6ON.Visible = False
I6OFF.Visible = True


I7ENTER.Visible = False
I7ON.Visible = False
I7OFF.Visible = True


I8ENTER.Visible = False
I8ON.Visible = False
I8OFF.Visible = True


I9ENTER.Visible = False
I9ON.Visible = False
I9OFF.Visible = True


I11ENTER.Visible = False
I11ON.Visible = False
I11OFF.Visible = True

End Sub

Private Sub i10ON_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
I10ENTER.Visible = True
I10ON.Visible = False
End Sub
Private Sub i11OFF_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
I11ON.Visible = True
I11OFF.Visible = False

'**
I1ENTER.Visible = False
I1ON.Visible = False
I1off.Visible = True


I2ENTER.Visible = False
I2ON.Visible = False
I2OFF.Visible = True




I3ENTER.Visible = False
I3ON.Visible = False
I3OFF.Visible = True


I4ENTER.Visible = False
I4ON.Visible = False
I4OFF.Visible = True


I5ENTER.Visible = False
I5ON.Visible = False
I5OFF.Visible = True


I6ENTER.Visible = False
I6ON.Visible = False
I6OFF.Visible = True


I7ENTER.Visible = False
I7ON.Visible = False
I7OFF.Visible = True


I8ENTER.Visible = False
I8ON.Visible = False
I8OFF.Visible = True


I9ENTER.Visible = False
I9ON.Visible = False
I9OFF.Visible = True


I10ENTER.Visible = False
I10ON.Visible = False
I10OFF.Visible = True


End Sub

Private Sub i11ON_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
I11ENTER.Visible = True
I11ON.Visible = False
End Sub

Sub hideb()
I1ENTER.Visible = False
I1ON.Visible = False
I1off.Visible = True


I2ENTER.Visible = False
I2ON.Visible = False
I2OFF.Visible = True




I3ENTER.Visible = False
I3ON.Visible = False
I3OFF.Visible = True


I4ENTER.Visible = False
I4ON.Visible = False
I4OFF.Visible = True


I5ENTER.Visible = False
I5ON.Visible = False
I5OFF.Visible = True


I6ENTER.Visible = False
I6ON.Visible = False
I6OFF.Visible = True


I7ENTER.Visible = False
I7ON.Visible = False
I7OFF.Visible = True


I8ENTER.Visible = False
I8ON.Visible = False
I8OFF.Visible = True


I9ENTER.Visible = False
I9ON.Visible = False
I9OFF.Visible = True


I10ENTER.Visible = False
I10ON.Visible = False
I10OFF.Visible = True


I11ENTER.Visible = False
I11ON.Visible = False
I11OFF.Visible = True

End Sub
 Sub SendCt1_Click()
Dim NewMessage As New NewMsg
    On Error Resume Next

    ' Index = 6: Compose New Messa0ge
    '       = 7: Reply
    '       = 8: Reply All
    '       = 9: Forward

    ' Save the header information and current note text.
    If Index > 6 Then
        ' SVNote = GetHeader(frmMain.MapiMess) + frmMain.MapiMess.MsgNoteText
        SVNote = frmMain.MapiMess.MsgNoteText
        SVNote = GetHeader(frmMain.MapiMess) + SVNote
    End If

    frmMain.MapiMess.Action = Index

    ' Set the new message text.
    If Index > 6 Then
        frmMain.MapiMess.MsgNoteText = SVNote
        
    End If

    If SendWithMapi Then
        frmMain.MapiMess.Action = vbMessageSendDlg
       ' frmMain.MapiMess.AddressLabel = "ggg"
    Else
        Call LoadMessage(-1, NewMessage)            ' Load message into frmMain NewMSG window.
    End If
End Sub

