VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4470
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00808080&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":08CA
   ScaleHeight     =   4470
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   3060
      Top             =   2730
   End
   Begin VB.PictureBox p1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5295
      Picture         =   "frmSplash.frx":8466C
      ScaleHeight     =   375
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   3345
      Width           =   15
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ßÇÝÉ ÇáÍÞæÞ ãÍÝæÙÉ áÔÑßÉ ÅãíßÓ ÏæÊ äÊ áÊßäæáæÌíÇ ÇáãÚáæãÇÊ"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3840
      TabIndex        =   1
      Top             =   480
      Width           =   4425
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetVolumeInformation _
 Lib "kernel32" Alias "GetVolumeInformationA" _
 (ByVal lpRootPathName As String, _
 ByVal lpVolumeNameBuffer As String, _
 ByVal nVolumeNameSize As Long, _
 lpVolumeSerialNumber As Long, _
 lpMaximumComponentLength As Long, _
 lpFileSystemFlags As Long, _
 ByVal lpFileSystemNameBuffer As String, _
 ByVal nFileSystemNameSize As Long) As Long
 Dim X As String
 Dim y As String
 Dim A As String
 Dim B As String
 Dim z As String
  Dim f As String
 Dim i As Integer
 
 
 
Option Explicit

Private Sub Form_Click()
Unload Me
frmMain.Show
End Sub

Private Sub Timer2_Timer()
    If p1.Width < 2985 Then
         p1.Width = p1.Width + 165

    Else
        Timer2.Enabled = False
        Unload frmSplash
        frmMain.Show
    End If
End Sub

Private Sub Frame1_Click()
End Sub

Public Sub CHN()
On Error GoTo h
Dim obj_FSO As Object, obj_Drive As Object
 Set obj_FSO = CreateObject("Scripting.FileSystemObject")
 Set obj_Drive = obj_FSO.GetDrive("c:\")
 Dim Serial&
 GetVolumeInformation "c:\", "", 255, Serial, _
  0, 0, "", 255
 'MsgBox Trim(Str(Serial))
X = Trim(Str(Serial))
y = ((Val(X) + 123456789 + 264779970))
Dim tmpname As String, tmpnumber As String
Open "numpe" For Input As 257
Line Input #257, tmpname
f = Len(y)

A = 0
For i = 1 To f
z = Right(y, i)
A = Val(A) + z
Next

f = Len(y)
For i = 1 To f
z = Left(y, i)
A = Val(A) + z
Next
If tmpname = A Then
frmMain.Show
Unload Me
Set frmSplash = Nothing
Else
frmser.Show
tmpname = A
frmser.chk (y), (tmpname)
 Unload Me
End If
'ÝÍÕ ÇáÑÞã
'áæ ãæÌæÏ

Close 257#

h:
If Err.Number = 53 Then

'Open "numpe" For Output As 257
'Close 257#
frmser.Show
f = Len(y)

A = 0
For i = 1 To f
z = Right(y, i)
A = Val(A) + z
Next

f = Len(y)
For i = 1 To f
z = Left(y, i)
A = Val(A) + z
Next

frmser.Show
tmpname = A
frmser.chk (y), (tmpname)
 Unload Me
 ElseIf Err.Number = 55 Then
  Close 257#
frmser.Show
frmser.chk (y), (tmpname)
 End If
End Sub


