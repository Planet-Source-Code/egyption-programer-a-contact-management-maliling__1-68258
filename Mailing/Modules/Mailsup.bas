Attribute VB_Name = "Module1"
Public Const conMailLongDate = 0
Public Const conMailListView = 1
Dim db As Database
Dim ws As Workspace
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim SqLst As String
Dim SqLst1 As String
Public Const conOptionGeneral = 1       ' Constant for Option Dialog Type - General Options
Public Const conOptionMessage = 2       ' Constant for Option Dialog Type - Message Options

Public Const conUnreadMessage = "*"     ' Constant for string to indicate unread message

Public Const vbRecipTypeTo = 1
Public Const vbRecipTypeCc = 2

Public Const vbMessageFetch = 1
Public Const vbMessageSendDlg = 2
Public Const vbMessageSend = 3
Public Const vbMessageSaveMsg = 4
Public Const vbMessageCopy = 5
Public Const vbMessageCompose = 6
Public Const vbMessageReply = 7
Public Const vbMessageReplyAll = 8
Public Const vbMessageForward = 9
Public Const vbMessageDelete = 10
Public Const vbMessageShowAdBook = 11
Public Const vbMessageShowDetails = 12
Public Const vbMessageResolveName = 13
Public Const vbRecipientDelete = 14
Public Const vbAttachmentDelete = 15

Public Const vbAttachTypeData = 0
Public Const vbAttachTypeEOLE = 1
Public Const vbAttachTypeSOLE = 2

Type ListDisplay
    Name As String * 20
    Subject As String * 40
    Date As String * 20
End Type

Public currentRCIndex As Integer
Public UnRead As Integer
Public SendWithMapi As Integer
Public ReturnRequest As Integer
Public OptionType As Integer

' Windows API functions
#If Win32 Then
    Declare Function GetProfileString Lib "kernel32" (ByVal lpAppName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
#Else
    Declare Function GetProfileString% Lib "Kernel" (ByVal lpSection$, ByVal lpEntry$, ByVal lpDefault$, ByVal Buffer$, ByVal cbBuffer%)
#End If

Sub Attachments(Msg As Form)
On Error Resume Next
    ' Clear the current attachment list.
    Msg.aList.Clear

    ' If there are attachments, load them into the list box.
    If frmMain.MapiMess.AttachmentCount Then
        Msg.NumAtt = frmMain.MapiMess.AttachmentCount & " Files"
        For i% = 0 To frmMain.MapiMess.AttachmentCount - 1
            frmMain.MapiMess.AttachmentIndex = i%
            A$ = frmMain.MapiMess.AttachmentName
            Select Case frmMain.MapiMess.AttachmentType
                Case vbAttachTypeData
                    A$ = A$ + " (Data File)"
                Case vbAttachTypeEOLE
                    A$ = A$ + " (Embedded OLE Object)"
                Case vbAttachTypeSOLE
                    A$ = A$ + " (Static OLE Object)"
                Case Else
                    A$ = A$ + " (Unknown attachment type)"
            End Select
            Msg.aList.AddItem A$
        Next i%
        
        If Not Msg.AttachWin.Visible Then
            Msg.AttachWin.Visible = True
            Call SizeMessageWindow(Msg)
            ' If Msg.WindowState = 0 Then
            '    Msg.Height = Msg.Height + Msg.AttachWin.Height
            ' End If
        End If
    
    Else
        If Msg.AttachWin.Visible Then
            Msg.AttachWin.Visible = False
            Call SizeMessageWindow(Msg)
            ' If Msg.WindowState = 0 Then
            '    Msg.Height = Msg.Height - Msg.AttachWin.Height
            ' End If
        End If
    End If
    Msg.Refresh
End Sub

Sub CopyNamestoMsgBuffer(Msg As Form, fResolveNames As Integer)
    On Error Resume Next
    Call KillRecips(frmMain.MapiMess)
    Call SetRCList(NewMsg.txtTo, frmMain.MapiMess, vbRecipTypeTo, fResolveNames)
    Call SetRCList(NewMsg.txtcc, frmMain.MapiMess, vbRecipTypeCc, fResolveNames)
End Sub

Function DateFromMapiDate$(ByVal s$, wFormat%)
' This procedure formats a MAPI date in one of
' two formats for viewing the message.
    y$ = Left$(s$, 4)
    m$ = Mid$(s$, 6, 2)
    d$ = Mid$(s$, 9, 2)
    T$ = Mid$(s$, 12)
    Ds# = DateValue(m$ + "/" + d$ + "/" + y$) + TimeValue(T$)
    Select Case wFormat
        Case conMailLongDate
            f$ = "dddd, mmmm d, yyyy, h:mmAM/PM"
        Case conMailListView
            f$ = "mm/dd/yy hh:mm"
    End Select
    DateFromMapiDate = Format$(Ds#, f$)
End Function

Sub DeleteMessage()
On Error Resume Next
    ' If the currently active form is a message, set MListIndex to
    ' the correct value.
    If TypeOf Screen.ActiveForm Is MsgView Then
        MailLst.MList.ListIndex = Val(Screen.ActiveForm.Tag)
        ViewingMsg = True
    End If

   ' Delete the mail message.
    If MailLst.MList.ListIndex <> -1 Then
        frmMain.MapiMess.MsgIndex = MailLst.MList.ListIndex
        frmMain.MapiMess.Action = vbMessageDelete
        X% = MailLst.MList.ListIndex
        MailLst.MList.RemoveItem X%
        If X% < MailLst.MList.ListCount - 1 Then
            MailLst.MList.ListIndex = X%
        Else
            MailLst.MList.ListIndex = MailLst.MList.ListCount - 1
        End If
        frmMain.MsgCountLbl = Format$(frmMain.MapiMess.MsgCount) + " Messages"

        ' Adjust the index values for currently viewed messages.
        If ViewingMsg Then
            Screen.ActiveForm.Tag = Str$(-1)
        End If

        For i = 0 To Forms.Count - 1
            If TypeOf Forms(i) Is MsgView Then
                If Val(Forms(i).Tag) > X% Then
                    Forms(i).Tag = Val(Forms(i).Tag) - 1
                End If
            End If
        Next i
        
        ' If the user is viewing a message, load the next message into the MsgView form
        ' if the message isn't currently displayed.
        If ViewingMsg Then
            ' First check to see if the message is currently being viewed.
            WindowNum% = FindMsgWindow((MailLst.MList.ListIndex))
            If WindowNum% > 0 Then
                If Forms(WindowNum%).Caption <> Screen.ActiveForm.Caption Then
                    Unload Screen.ActiveForm
                     ' Find the correct window again and display it.  The index isn't valid after the unload.
                     Forms(FindMsgWindow((MailLst.MList.ListIndex))).Show
                Else
                     Forms(WindowNum%).Show
                End If
            Else
                Call LoadMessage(MailLst.MList.ListIndex, Screen.ActiveForm)
            End If
        Else
            ' Check to see if there was a window viewing the message, and unload the window.
            WindowNum% = FindMsgWindow(X%)
            If WindowNum% > 0 Then
                Unload Forms(X%)
            End If
        End If
     End If
End Sub

Sub DisplayAttachedFile(ByVal FileName As String)
On Error Resume Next
        ' Determine the filename extension.
        ext$ = FileName
        junk$ = Token$(ext$, ".")
        ' Get the application from the WIN.INI file.
        Buffer$ = String$(256, " ")
        errCode% = GetProfileString("Extensions", ext$, "NOTFOUND", Buffer$, Len(Left(Buffer$, Chr(0)) - 1))
        If errCode% Then
            Buffer$ = Mid$(Buffer$, 1, InStr(Buffer$, Chr(0)) - 1)
            If Buffer$ <> "NOTFOUND" Then
                ' Strip off the ^.EXT information from the string.
                EXEName$ = Token$(Buffer$, " ")
                errCode% = Shell(EXEName$ + " " + FileName, 1)
                If Err Then
                    MsgBox "Error occurred during the shell: " + Error$
                End If
            Else
                MsgBox "Application that uses: <" + ext$ + "> not found in WIN.INI"
            End If
        End If
End Sub

Function FindMsgWindow(Index As Integer) As Integer
On Error Resume Next
' This function searches through the active windows
' and locates those with the MsgView type and then
' checks to see if the tag contains the index the user
' is searching for.
        For i = 0 To Forms.Count - 1
            If TypeOf Forms(i) Is MsgView Then
                If Val(Forms(i).Tag) = Index Then
                    FindMsgWindow = i
                    Exit Function
                End If
            End If
        Next i
        FindMsgWindow = -1
End Function

Function GetHeader(Msg As Control) As String
On Error Resume Next
Dim CR As String
CR = Chr$(13) + Chr$(10)
      Header$ = String$(25, "-") + CR
      Header$ = Header$ + "Form: " + Msg.MsgOrigDisplayName + CR
      Header$ = Header$ + "To: " + GetRCList(Msg, vbRecipTypeTo) + CR
      Header$ = Header$ + "Cc: " + GetRCListto(Msg, vbRecipTypeCc) + CR
      Header$ = Header$ + "Subject: " + Msg.MsgSubject + CR
      Header$ = Header$ + "Date: " + DateFromMapiDate$(Msg.MsgDateReceived, conMailLongDate) + CR + CR
      GetHeader = Header$
End Function

Sub GetMessageCount()
On Error Resume Next
    '  Reads all mail messages and displays the count.
    Screen.MousePointer = 11
    frmMain.MapiMess.FetchUnreadOnly = 0
    frmMain.MapiMess.Action = vbMessageFetch
    frmMain.MsgCountLbl = Format$(frmMain.MapiMess.MsgCount) + " Messages"
    Screen.MousePointer = 0
End Sub

Function GetRCList(Msg As Control, RCType As Integer) As String
' Given a list of recipients, this function returns
' a list of recipients of the specified type in the
' following format:
'
'       Person 1;Person 2;Person 3
On Error Resume Next
    For i = 0 To Msg.RecipCount - 1
        Msg.RecipIndex = i
        If RCType = Msg.RecipType Then
                A$ = A$ + ";" + Msg.RecipDisplayName
        End If
    Next i
    If A$ <> "" Then
       A$ = Mid$(A$, 2)  ' Strip off the leading ";".
    End If
    GetRCList = A$
End Function
Function GetRCListto(Msg As Control, RCType As Integer) As String
' Given a list of recipients, this function returns
' a list of recipients of the specified type in the
' following format:
'
'       Person 1;Person 2;Person 3
On Error Resume Next

    For i = 0 To Msg.RecipCount - 1
        Msg.RecipIndex = i
        If RCType = Msg.RecipType Then
                A$ = A$ + ";" + Msg.RecipDisplayName
        End If
    Next i
    If A$ <> "" Then
       A$ = Mid$(A$, 2)  ' Strip off the leading ";".
    End If
    GetRCListto = A$
End Function
Sub KillRecips(MsgControl As Control)
On Error Resume Next
    ' Delete each recipient.  Loop until no recipients exist.
    While MsgControl.RecipCount
        MsgControl.Action = vbRecipientDelete
    Wend
End Sub

Sub LoadList(mailctl As Control)
On Error Resume Next
' This procedure loads the mail message headers
' into the MailLst.MList.  Unread messages have
' conUnreadMessage placed at the beginning of the string.
    MailLst.MList.Clear
    UnRead = 0
    StartIndex = 0
    For i = 0 To mailctl.MsgCount - 1
        mailctl.MsgIndex = i
        If Not mailctl.MsgRead Then
            A$ = conUnreadMessage + " "
            If UnRead = 0 Then
                StartIndex = i  ' Start position in the mail list.
            End If
            UnRead = UnRead + 1
        Else
            A$ = "  "
        End If
        A$ = A$ + Mid$(Format$(mailctl.MsgOrigDisplayName, "!" + String$(10, "@")), 1, 10)
        If mailctl.MsgSubject <> "" Then
            B$ = Mid$(Format$(mailctl.MsgSubject, "!" + String$(35, "@")), 1, 35)
        Else
            B$ = String$(30, " ")
        End If
        c$ = Mid$(Format$(DateFromMapiDate(mailctl.MsgDateReceived, conMailListView), "!" + String$(15, "@")), 1, 15)
        MailLst.MList.AddItem A$ + Chr$(9) + B$ + Chr$(9) + c$
        MailLst.MList.Refresh
    Next i

    MailLst.MList.ListIndex = StartIndex
    
    ' Enable the correct buttons.
    'frmMain.Next.Enabled = True
    'frmMain.Previous.Enabled = True
    'frmMain![Delete].Enabled = True

    ' Adjust the value of the labels displaying message counts.
    If UnRead Then
        frmMain.UnreadLbl = " - " + Format$(UnRead) + " Unread"
        MailLst.Icon = MailLst.NewMail.Picture
    Else
        frmMain.UnreadLbl = ""
        MailLst.Icon = MailLst.nonew.Picture
    End If
End Sub
    

Sub LoadMessage(ByVal Index As Integer, Msg As Form)
' This procedure loads the specified mail message into
' a form to either view or edit a message.

On Error Resume Next
    If TypeOf Msg Is MsgView Then
        A$ = MailLst.MList.List(Index)
        ' Message is unread; reset the text.
        If Mid$(A$, 1, 1) = conUnreadMessage Then
            Mid$(A$, 1, 1) = " "
            MailLst.MList.List(Index) = A$
            UnRead = UnRead - 1
            If UnRead Then
                frmMain.UnreadLbl = Format$(UnRead) + " Unread"
            Else
                frmMain.UnreadLbl = ""
                ' Change the icon on the list window.
                MailLst.Icon = MailLst.nonew.Picture
            End If
        End If
    End If

    ' These fields only apply to viewing.
    If TypeOf Msg Is MsgView Then
        frmMain.MapiMess.MsgIndex = Index
        Msg.txtDate = DateFromMapiDate$(frmMain.MapiMess.MsgDateReceived, conMailLongDate)
        Msg.txtFrom = frmMain.MapiMess.MsgOrigDisplayName
        MailLst.MList.ItemData(Index) = True
    End If
    ' These fields apply to both form types.
    Call Attachments(Msg)
    Msg.txtNoteText = frmMain.MapiMess.MsgNoteText
    Msg.txtsubject = frmMain.MapiMess.MsgSubject
    Msg.Caption = frmMain.MapiMess.MsgSubject
    Msg.Tag = Index
    Call UpdateRecips(Msg)
    Msg.Refresh
    Msg.Show
End Sub

Sub LogOffUser()
    On Error Resume Next
    frmMain.MapiSess.Action = 2
    If Err <> 0 Then
        MsgBox "Logoff Failure: " + ErrorR
    Else
        frmMain.MapiMess.SessionID = 0
        ' Adjust the menu items.
        frmMain.LogOff.Enabled = 0
        frmMain.Logon.Enabled = -1
        ' Unload all forms except the MDI form.
        Do Until Forms.Count = 1
            i = Forms.Count - 1
            If TypeOf Forms(i) Is MDIForm Then
                ' Do nothing.
            Else
                Unload Forms(i)
            End If
        Loop
        ' Disable the toolbar buttons.
      '  frmMain.Next.Enabled = False
      '  frmMain.Previous.Enabled = False
      '  frmMain![Delete].Enabled = False
      '  frmMain.SendCtl(vbMessageCompose).Enabled = False
      '  frmMain.SendCtl(vbMessageReplyAll).Enabled = False
      '  frmMain.SendCtl(vbMessageReply).Enabled = False
      '  frmMain.SendCtl(vbMessageForward).Enabled = False
        frmMain.rMsgList.Enabled = False
        frmMain.PrintMessage.Enabled = False
        frmMain.DispTools.Enabled = False
        frmMain.EditDelete.Enabled = False
                          
        ' Reset the caption for the status bar labels.
        frmMain.MsgCountLbl = "Off Line"
        frmMain.UnreadLbl = ""
    End If

End Sub

Sub PrintLongText(ByVal LongText As String)
On Error Resume Next
' This procedure prints a text stream to a printer and
' ensures that words are not split between lines and
' that they wrap as needed.
    Do Until LongText = ""
        Word$ = Token$(LongText, " ")
        If Printer.TextWidth(Word$) + Printer.CurrentX > Printer.Width - Printer.TextWidth("ZZZZZZZZ") Then
            Printer.Print
        End If
        Printer.Print " " + Word$;
    Loop
End Sub

Sub PrintMail()
    ' In List view, all selected messages are printed.
    ' In Message view, the selected message is printed.
On Error Resume Next
    If TypeOf Screen.ActiveForm Is MsgView Then
        Call PrintMessage(frmMain.MapiMess, False)
        Printer.EndDoc
    ElseIf TypeOf Screen.ActiveForm Is MailLst Then
        For i = 0 To MailLst.MList.ListCount - 1
            If MailLst.MList.Selected(i) Then
                frmMain.MapiMess.MsgIndex = i
                Call PrintMessage(frmMain.MapiMess, False)
            End If
        Next i
        Printer.EndDoc
    End If
End Sub

Sub PrintMessage(Msg As Control, fNewPage As Integer)
On Error Resume Next
'   This procedure prints a mail message.
    Screen.MousePointer = 11
    ' Start a new page if needed.
    If fNewPage Then
        Printer.NewPage
    End If
    Printer.FontName = "Arial"
    Printer.FontBold = True
    Printer.DrawWidth = 10
    Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
    Printer.Print
    Printer.FontSize = 9.75
    Printer.Print "From:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print Msg.MsgOrigDisplayName
    Printer.Print "To:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print GetRCList(Msg, vbRecipTypeTo)
    Printer.Print "Cc:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print GetRCList(Msg, vbRecipTypeCc)
    Printer.Print "Subject:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print Msg.MsgSubject
    Printer.Print "Date:";
    Printer.CurrentX = Printer.TextWidth(String$(30, " "))
    Printer.Print DateFromMapiDate$(Msg.MsgDateReceived, conMailLongDate)
    Printer.Print
    Printer.DrawWidth = 5
    Printer.Line (0, Printer.CurrentY)-(Printer.Width, Printer.CurrentY)
    Printer.FontSize = 9.75
    Printer.FontBold = False
    Call PrintLongText(Msg.MsgNoteText)
    Printer.Print
    Screen.MousePointer = 0
End Sub

Sub SaveMessage(Msg As Form)
On Error Resume Next
    ' Save the current subject and note text.
    ' Copy the message to the compose buffer.
    ' Reset the subject and message text.
    ' Save the message.
    svSub = Msg.txtsubject
    SVNote = Msg.txtNoteText
    frmMain.MapiMess.Action = vbMessageCopy
    frmMain.MapiMess.MsgSubject = svSub
    frmMain.MapiMess.MsgNoteText = SVNote
    frmMain.MapiMess.Action = vbMessageSaveMsg
End Sub

Sub SetRCList(ByVal NameList As String, Msg As Control, RCType As Integer, fResolveNames As Integer)
'1 Given a list of recipients:
'
'       Person 1;Person 2;Person 3
'
' this procedure places the names into the Msg.Recip
' structures.
    On Error Resume Next
    If NameList = "" Then
        Exit Sub
    End If

    i = Msg.RecipCount
    Do
        Msg.RecipIndex = i
        Msg.RecipDisplayName = Trim$(Token(NameList, ";"))
        If fResolveNames Then
            Msg.Action = vbMessageResolveName
        End If
        Msg.RecipType = RCType
        i = i + 1
    Loop Until (NameList = "")
End Sub

Sub SizeMessageWindow(MsgWindow As Form)
On Error Resume Next
    If MsgWindow.WindowState <> 1 Then
        ' Determine the minimum window size based
        ' on the visiblity of AttachWin (Attachment window).
        If MsgWindow.AttachWin.Visible Then    ' Attachment window.
            MinSize = 3700
        Else
            MinSize = 3700 - MsgWindow.AttachWin.Height
        End If

        ' Maintain the minimum form size.
        If MsgWindow.Height < MinSize And (MsgWindow.WindowState = 0) Then
            MsgWindow.Height = MinSize
            Exit Sub

        End If
        ' Adjust the size of the text box.
        If MsgWindow.ScaleHeight > MsgWindow.txtNoteText.Top Then
            If MsgWindow.AttachWin.Visible Then
                X% = MsgWindow.AttachWin.Height
            Else
                X% = 0
            End If
            MsgWindow.txtNoteText.Height = MsgWindow.ScaleHeight - MsgWindow.txtNoteText.Top - X%
            MsgWindow.txtNoteText.Width = MsgWindow.ScaleWidth
        End If
    End If

End Sub

Function Token$(tmp$, search$)
    On Error Resume Next
    X = InStr(1, tmp$, search$)
    If X Then
       Token$ = Mid$(tmp$, 1, X - 1)
       tmp$ = Mid$(tmp$, X + 1)
    Else
       Token$ = tmp$
       tmp$ = ""
    End If
End Function

Sub UpdateRecips(Msg As Form)
On Error Resume Next
' This procedure updates the correct edit fields and the
' recipient information.
    Msg.txtTo.Text = GetRCList(frmMain.MapiMess, vbRecipTypeTo)
    Msg.txtcc.Text = GetRCList(frmMain.MapiMess, vbRecipTypeCc)
End Sub

Sub ViewNextMsg()
On Error Resume Next
    ' Check to see if the message is currently loaded.
    ' If it is loaded, show that form.
    ' If it is not loaded, load the message.
    WindowNum% = FindMsgWindow((MailLst.MList.ListIndex))
    If WindowNum% > 0 Then
        Forms(WindowNum%).Show
    Else
        If TypeOf Screen.ActiveForm Is MsgView Then
            Call LoadMessage(MailLst.MList.ListIndex, Screen.ActiveForm)
        Else
            Dim Msg As New MsgView
            Call LoadMessage(MailLst.MList.ListIndex, Msg)
        End If
    End If
End Sub


Public Function sendM(s As Integer)
On Error Resume Next
'frmsendM.CmbCompanyName.Clear
B:
If s <> 0 Then Module1.Msendmail
s = 0
A$ = ""
    Do While Not rs1.EOF
   ' m = rs1.AbsolutePosition
    If s > 50 Then GoTo e:
    If IsNull(rs1("Email")) Or rs1("Email") = "Empty" Then
    Else
    A$ = A$ + ";" & rs1("Email")
    s = s + 1
    
    End If
   ' If A$ <> "" Then
   '    A$ = Mid$(A$, 2)  ' Strip off the leading ";".
   ' End If
    NewMsg.txtTo.Text = A$
       rs1.MoveNext
    Loop
   Module1.Msendmail
    Exit Function
e:
 
    s = 0
    A$ = ""
Do While Not rs1.EOF
m = rs1.AbsolutePosition
    If s > 50 Then GoTo B:
  If IsNull(rs1("Email")) Or rs1("Email") = "Empty" Then
    Else
    A$ = A$ + ";" & rs1("Email")
    s = s + 1
    End If
   '     If A$ <> "" Then
   '    A$ = Mid$(A$, 2)  ' Strip off the leading ";".
   ' End If
    NewMsg.txtcc.Text = A$
       rs1.MoveNext
    Loop
  Module1.Msendmail
   Exit Function
End Function
Public Function sendMper(s As Integer)
On Error Resume Next
'frmsendM.CmbCompanyName.Clear

X = 0
B:
If s <> 0 Then Module1.Msendmail
s = 0
A$ = ""

    Do While X < NewMsg.Lstpersonalname.ListCount
NewMsg.Lstpersonalname.ListIndex = X
If NewMsg.Lstpersonalname.Selected(NewMsg.Lstpersonalname.ListIndex) = True Then
SqLst1 = "Select Email From Personal"
SqLst1 = SqLst1 & " WHERE name = '" & NewMsg.Lstpersonalname.List(NewMsg.Lstpersonalname.ListIndex) & "'"
Set rs1 = db.OpenRecordset(SqLst1, dbOpenDynaset)
    If rs1.RecordCount <> 0 Then
    If s > 50 Then GoTo e:
    If IsNull(rs1("Email")) Or rs1("Email") = "Empty" Then
    Else
    A$ = A$ + ";" & rs1("Email")
    s = s + 1
    
    End If
   ' If A$ <> "" Then
   '    A$ = Mid$(A$, 2)  ' Strip off the leading ";".
   ' End If
    NewMsg.txtTo.Text = A$
    End If
       X = X + 1
       
       Else
       X = X + 1
       End If
    Loop
   Module1.Msendmail
    Exit Function
e:
 
    s = 0
    A$ = ""
Do While X < NewMsg.Lstpersonalname.ListCount
If NewMsg.Lstpersonalname.Selected(NewMsg.Lstpersonalname.ListIndex) = True Then
SqLst1 = "Select Email From Personal"
SqLst1 = SqLst1 & " WHERE name = '" & NewMsg.Lstpersonalname.List(NewMsg.Lstpersonalname.ListIndex) & "'"
Set rs1 = db.OpenRecordset(SqLst1, dbOpenDynaset)
 If rs1.RecordCount <> 0 Then
    If s > 50 Then GoTo B:
  If IsNull(rs1("Email")) Or rs1("Email") = "Empty" Then
    Else
    A$ = A$ + ";" & rs1("Email")
    s = s + 1
    End If
   '     If A$ <> "" Then
   '    A$ = Mid$(A$, 2)  ' Strip off the leading ";".
   ' End If
    NewMsg.txtcc.Text = A$
    End If
      X = X + 1
       Else
       X = X + 1
       End If
    Loop
  Module1.Msendmail
   Exit Function
End Function

Public Function sendM1(s As Integer)
On Error Resume Next
'frmsendM.CmbCompanyName.Clear
B:
If s <> 0 Then Module1.Msendmail
s = 0
A$ = ""
    Do While Not rs1.EOF
    m = rs1.AbsolutePosition
    If s > 50 Then GoTo e:
    If IsNull(rs1("Email2")) Or rs1("Email2") = "Empty" Then
    Else
    A$ = A$ + ";" & rs1("Email2")
    s = s + 1
    
    End If
    'If A$ <> "" Then
    '   A$ = Mid$(A$, 2)  ' Strip off the leading ";".
    'End If
    NewMsg.txtTo.Text = A$
       rs1.MoveNext
    Loop
   Module1.Msendmail
    Exit Function
e:
 
    s = 0
    A$ = ""
Do While Not rs1.EOF
m = rs1.AbsolutePosition
    If s > 50 Then GoTo B:
  If IsNull(rs1("Email2")) Or rs1("Email2") = "Empty" Then
    Else
    A$ = A$ + ";" & rs1("Email2")
    s = s + 1
    End If
      '  If A$ <> "" Then
      ' A$ = Mid$(A$, 2)  ' Strip off the leading ";".
    'End If
    NewMsg.txtcc.Text = A$
       rs1.MoveNext
    Loop
  Module1.Msendmail
   Exit Function
End Function
Public Function Msendmail()
On Error Resume Next
If NewMsg.txtTo.Text <> "" Then
       NewMsg.txtTo.Text = Mid$(NewMsg.txtTo.Text, 2)  ' Strip off the leading ";".
    End If
If NewMsg.txtcc.Text <> "" Then
NewMsg.txtcc.Text = Mid$(NewMsg.txtcc.Text, 2)  ' Strip off the leading ";".
End If
If frmMain.MapiMess.AttachmentCount > 0 Then
        txtNoteText = String$(frmMain.MapiMess.AttachmentCount, "*") + txtNoteText
    End If
    frmMain.MapiMess.MsgSubject = NewMsg.txtsubject.Text
    frmMain.MapiMess.MsgNoteText = NewMsg.txtNoteText.Text
       frmMain.MapiMess.MsgReceiptRequested = ReturnRequest
    Call CopyNamestoMsgBuffer(NewMsg, True)
                  
    On Error Resume Next
    frmMain.MapiMess.Action = vbMessageSend
    If Err Then
        MsgBox "An error occurred during a send: " + Str$(Err)
      ' MsgBox ("Your Mail not send")
       frmMain.Label2.Caption = "Waite..."
       NewMsg.txtTo.Text = ""
        NewMsg.txtcc.Text = ""
    Else
        'Unload Me
        frmMain.Label2.Caption = "Your Mail is sending"
        NewMsg.txtTo.Text = ""
        NewMsg.txtcc.Text = ""
        
    End If
     frmMain.Label2.Caption = ""
End Function

Public Function searshm(s As Integer)
If Right(App.Path, 1) = "\" Then
            PathPro = App.Path & "Mailing.mdb"
        Else
            PathPro = App.Path & "\Mailing.mdb"
        End If

On Error Resume Next
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
If NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select Email From Qcompany"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'1==================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
    SqLst1 = "Select Email From Qcompany"
SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'2======================================
 ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select Email From Qcompany"
SqLst1 = SqLst1 & " WHERE country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'3===========================================

ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
    SqLst1 = "Select Email From Qcompany"
SqLst1 = SqLst1 & " WHERE City = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
 '4===========================================
 ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
    SqLst1 = "Select Email From Qcompany"
SqLst1 = SqLst1 & " WHERE City = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'5===================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text <> "" Then
    SqLst1 = "Select Email From Qcatype"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'6=================
 ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text <> "" Then
     SqLst1 = "Select Email From Qprodtype"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'7=================

ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
     SqLst1 = "Select Email From Qcatype"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'8=================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
    SqLst1 = "Select Email From Qprodtype"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'9=================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
SqLst1 = "Select Email From Qcatype"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'and country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and City = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'10=================
 ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
    SqLst1 = "Select Qcompany.Email From Qprodtype,Qcompany"
SqLst1 = SqLst1 & " WHERE Qprodtype.Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Qprodtype.Typename = '" & NewMsg.Cmbtype.Text & "'and Qcompany.country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and Qcompany.City = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'11===============
  ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
 SqLst1 = "Select Email From Qcatype"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'and country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and City = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'12===============
 ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
     SqLst1 = "Select Qcompany.Email From Qprodtype,Qcompany"
SqLst1 = SqLst1 & " WHERE Qprodtype.Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Qprodtype.Typename = '" & NewMsg.Cmbtype.Text & "'and Qcompany.country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and Qcompany.city = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and Qcompany.Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'13===============
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select Email From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' ORDER BY CategoryID "
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function

'14============================================
'ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text = "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select Email From Qprocomp"
'SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
'SqLst1 = SqLst1 & "'"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.sendM (0)
'Exit Function
'=============================================
'ElseIf NewMsg.CmbCategoryname.Text = "All" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select Email From Qcatcomp"
'SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text
'SqLst1 = SqLst1 & "' ORDER BY CategoryID "
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.sendM (0)
'Exit Function
'============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text = "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select Email From Qprocomp"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.sendM (0)
'Exit Function
'===========================================
ElseIf NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select Email From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'15=============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text <> "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select Email From Qprocomp"
'SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text & "'"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.sendM (0)
'Exit Function
'============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "All" Then
'SqLst1 = "Select Email From Qtype"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.sendM (0)
'Exit Function
'============================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbtype.Text <> "" Then
SqLst1 = "Select Email From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'16=============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbtype.Text = "All" Then
'SqLst1 = "Select Email From Qtype"
'SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text & "'"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.sendM (0)
'Exit Function
'====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select Email From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'17====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
SqLst1 = "Select Email From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' and city = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
''18====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
SqLst1 = "Select Email From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'19====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select Email From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'20====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
SqLst1 = "Select Email From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "' and city = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
''21====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
SqLst1 = "Select Email From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "' and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
'22====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
SqLst1 = "Select Email From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
    '23====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text = "" Then
SqLst1 = "Select Email From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "' and city = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
''25====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text <> "" Then
SqLst1 = "Select Email From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "' and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
  ''26====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(0).Value = True Then
SqLst1 = "Select Email From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
  ''27====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(1).Value = True Then
SqLst1 = "Select Email From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Companyname like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
  ''28====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(2).Value = True Then
SqLst1 = "Select Email From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
Exit Function
  ''29====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(3).Value = True Then
SqLst1 = "Select Email From Qtype"
SqLst1 = SqLst1 & " WHERE Typename like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM (0)
End If
'30
End Function
Public Function searshm1(s As Integer)

On Error Resume Next
If Right(App.Path, 1) = "\" Then
            PathPro = App.Path & "Mailing.mdb"
        Else
            PathPro = App.Path & "\Mailing.mdb"
        End If

Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
If NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select Email2 From Qcompany"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
'===========================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select Email2 From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' ORDER BY CategoryID "
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
'============================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select Email2 From Qprocomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
'=============================================
ElseIf NewMsg.CmbCategoryname.Text = "All" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select Email2 From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text
SqLst1 = SqLst1 & "' ORDER BY CategoryID "
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
'============================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select Email2 From Qprocomp"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
'===========================================
ElseIf NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select Email2 From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
'=============================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text <> "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select Email2 From Qprocomp"
SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
'============================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "All" Then
SqLst1 = "Select Email2 From Qtype"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
'============================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbtype.Text <> "" Then
SqLst1 = "Select Email2 From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
'=============================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbtype.Text = "All" Then
SqLst1 = "Select Email2 From Qtype"
SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
'====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select Email2 From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
'====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
SqLst1 = "Select Email2 From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' and city = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
''====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
SqLst1 = "Select Email2 From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
'====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select Email2 From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
'====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
SqLst1 = "Select Email2 From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "' and city = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
''====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
SqLst1 = "Select Email2 From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "' and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
'====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
SqLst1 = "Select Email2 From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
    '====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text = "" Then
SqLst1 = "Select Email2 From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "' and city = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
''====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text <> "" Then
SqLst1 = "Select Email2 From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "' and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
  ''====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(0).Value = True Then
SqLst1 = "Select Email2 From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
  ''====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(1).Value = True Then
SqLst1 = "Select Email2 From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Companyname like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
  ''====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(2).Value = True Then
SqLst1 = "Select Email2 From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
Exit Function
  ''====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(3).Value = True Then
SqLst1 = "Select Email2 From Qtype"
SqLst1 = SqLst1 & " WHERE Typename like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.sendM1 (0)
End If

End Function
Public Function NotesSave()
On Error Resume Next
If Right(App.Path, 1) = "\" Then
            PathPro = App.Path & "Mailing.mdb"
        Else
            PathPro = App.Path & "\Mailing.mdb"
        End If
'Set ws = CreateWorkspace("", "admin", "")
'Set db = ws.OpenDatabase(PathPro)
'If NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select CompanyId From Qcompany"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.SaveM
'Exit Function
'===========================================
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
If NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select CompanyId From Qcompany"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'1==================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
    SqLst1 = "Select CompanyId From Qcompany"
SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'2======================================
 ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select CompanyId From Qcompany"
SqLst1 = SqLst1 & " WHERE country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'3===========================================

ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
    SqLst1 = "Select CompanyId From Qcompany"
SqLst1 = SqLst1 & " WHERE City = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
 '4===========================================
 ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
    SqLst1 = "Select CompanyId From Qcompany"
SqLst1 = SqLst1 & " WHERE City = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'5===================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text <> "" Then
    SqLst1 = "Select CompanyId From Qcatype"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'6=================
 ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text <> "" Then
     SqLst1 = "Select CompanyId From Qprodtype"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'7=================

ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
     SqLst1 = "Select CompanyId From Qcatype"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'8=================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
    SqLst1 = "Select CompanyId From Qprodtype"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'9=================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
SqLst1 = "Select CompanyId From Qcatype"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'and country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and City = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'10=================
 ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
    SqLst1 = "Select Qcompany.CompanyId From Qprodtype,Qcompany"
SqLst1 = SqLst1 & " WHERE Qprodtype.Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Qprodtype.Typename = '" & NewMsg.Cmbtype.Text & "'and Qcompany.country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and Qcompany.City = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'11===============
  ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
 SqLst1 = "Select CompanyId From Qcatype"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'and country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and City = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'12===============
 ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
     SqLst1 = "Select Qcompany.CompanyId From Qprodtype,Qcompany"
SqLst1 = SqLst1 & " WHERE Qprodtype.Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Qprodtype.Typename = '" & NewMsg.Cmbtype.Text & "'and Qcompany.country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and Qcompany.city = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and Qcompany.Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'13===============
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select CompanyId From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' ORDER BY CategoryID "
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function

'14============================================
'ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text = "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select CompanyId From Qprocomp"
'SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
'SqLst1 = SqLst1 & "'"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.SaveM
'Exit Function
'=============================================
'ElseIf NewMsg.CmbCategoryname.Text = "All" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select CompanyId From Qcatcomp"
'SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text
'SqLst1 = SqLst1 & "' ORDER BY CategoryID "
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.SaveM
'Exit Function
'============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text = "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select CompanyId From Qprocomp"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.SaveM
'Exit Function
'===========================================
ElseIf NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select CompanyId From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'15=============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text <> "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select CompanyId From Qprocomp"
'SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text & "'"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.SaveM
'Exit Function
'============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "All" Then
'SqLst1 = "Select CompanyId From Qtype"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.SaveM
'Exit Function
'============================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbtype.Text <> "" Then
SqLst1 = "Select CompanyId From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'16=============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbtype.Text = "All" Then
'SqLst1 = "Select CompanyId From Qtype"
'SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text & "'"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.SaveM
'Exit Function
'====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select CompanyId From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'17====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
SqLst1 = "Select CompanyId From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' and city = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
''18====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
SqLst1 = "Select CompanyId From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'19====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select CompanyId From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'20====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
SqLst1 = "Select CompanyId From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "' and city = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
''21====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
SqLst1 = "Select CompanyId From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "' and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'22====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
SqLst1 = "Select CompanyId From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
    '23====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text = "" Then
SqLst1 = "Select CompanyId From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "' and city = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
''25====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text <> "" Then
SqLst1 = "Select CompanyId From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "' and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
  ''26====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(0).Value = True Then
SqLst1 = "Select CompanyId From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
  ''27====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(1).Value = True Then
SqLst1 = "Select CompanyId From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Companyname like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
  ''28====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(2).Value = True Then
SqLst1 = "Select CompanyId From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
  ''29====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(3).Value = True Then
SqLst1 = "Select CompanyId From Qtype"
SqLst1 = SqLst1 & " WHERE Typename like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
End If

End Function
Public Function senNotesSave()
On Error Resume Next
If Right(App.Path, 1) = "\" Then
            PathPro = App.Path & "Mailing.mdb"
        Else
            PathPro = App.Path & "\Mailing.mdb"
        End If
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
If NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select CompanyId From Qcompany"
SqLst1 = SqLst1 & " WHERE email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'1==================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
    SqLst1 = "Select CompanyId From Qcompany"
SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'2======================================
 ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select CompanyId From Qcompany"
SqLst1 = SqLst1 & " WHERE country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'3===========================================

ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
    SqLst1 = "Select CompanyId From Qcompany"
SqLst1 = SqLst1 & " WHERE City = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
 '4===========================================
 ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
    SqLst1 = "Select CompanyId From Qcompany"
SqLst1 = SqLst1 & " WHERE City = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'and Region = '" & NewMsg.CmbRegion.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'5===================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text <> "" Then
    SqLst1 = "Select CompanyId From Qcatype"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'6=================
 ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text <> "" Then
     SqLst1 = "Select CompanyId From Qprodtype"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'7=================

ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
     SqLst1 = "Select CompanyId From Qcatype"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'and country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'8=================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
    SqLst1 = "Select CompanyId From Qprodtype"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'and country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'9=================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
SqLst1 = "Select CompanyId From Qcatype"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'and country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and City = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'10=================
 ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
    SqLst1 = "Select Qcompany.CompanyId From Qprodtype,Qcompany"
SqLst1 = SqLst1 & " WHERE Qprodtype.Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Qprodtype.Typename = '" & NewMsg.Cmbtype.Text & "'and Qcompany.country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and Qcompany.City = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'11===============
  ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
 SqLst1 = "Select CompanyId From Qcatype"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'and country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and City = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and Region = '" & NewMsg.CmbRegion.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'12===============
 ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
     SqLst1 = "Select Qcompany.CompanyId From Qprodtype,Qcompany"
SqLst1 = SqLst1 & " WHERE Qprodtype.Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Qprodtype.Typename = '" & NewMsg.Cmbtype.Text & "'and Qcompany.country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and Qcompany.city = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and Qcompany.Region = '" & NewMsg.CmbRegion.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'13===============
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select CompanyId From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function

'14============================================
'ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text = "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select CompanyId From Qprocomp"
'SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
'SqLst1 = SqLst1 & "'"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.SaveM
'Exit Function
'=============================================
'ElseIf NewMsg.CmbCategoryname.Text = "All" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select CompanyId From Qcatcomp"
'SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text
'SqLst1 = SqLst1 & "' ORDER BY CategoryID "
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.SaveM
'Exit Function
'============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text = "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select CompanyId From Qprocomp"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.SaveM
'Exit Function
'===========================================
ElseIf NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select CompanyId From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'15=============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text <> "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select CompanyId From Qprocomp"
'SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text & "'"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.SaveM
'Exit Function
'============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "All" Then
'SqLst1 = "Select CompanyId From Qtype"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.SaveM
'Exit Function
'============================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbtype.Text <> "" Then
SqLst1 = "Select CompanyId From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'16=============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbtype.Text = "All" Then
'SqLst1 = "Select CompanyId From Qtype"
'SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text & "'"
'Set rs1 = db.OpenRecordset(SqLst1)
'Module1.SaveM
'Exit Function
'====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select CompanyId From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'17====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
SqLst1 = "Select CompanyId From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' and city = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
''18====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
SqLst1 = "Select CompanyId From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' and Region = '" & NewMsg.CmbRegion.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'19====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select CompanyId From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'20====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
SqLst1 = "Select CompanyId From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "' and city = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
''21====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
SqLst1 = "Select CompanyId From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "' and Region = '" & NewMsg.CmbRegion.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
'22====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
SqLst1 = "Select CompanyId From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
    '23====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text = "" Then
SqLst1 = "Select CompanyId From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "' and city = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
''25====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text <> "" Then
SqLst1 = "Select CompanyId From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "' and Region = '" & NewMsg.CmbRegion.Text
SqLst1 = SqLst1 & "'and email = " & "'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
  ''26====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(0).Value = True Then
SqLst1 = "Select CompanyId From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname like" & _
"'*" & NewMsg.Text1.Text & "*' And email = " & " 'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
  ''27====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(1).Value = True Then
SqLst1 = "Select CompanyId From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Companyname like" & _
"'*" & NewMsg.Text1.Text & "*' And email = " & " 'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
  ''28====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(2).Value = True Then
SqLst1 = "Select CompanyId From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname like" & _
"'*" & NewMsg.Text1.Text & "*' And email = " & " 'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
Exit Function
  ''29====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(3).Value = True Then
SqLst1 = "Select CompanyId From Qtype"
SqLst1 = SqLst1 & " WHERE Typename like" & _
"'*" & NewMsg.Text1.Text & "*' And email = " & " 'Empty'"
Set rs1 = db.OpenRecordset(SqLst1)
Module1.SaveM
End If

End Function


Public Function SaveM()
On Error Resume Next
 If NewMsg.txtNoteText.Text = "" Then
 NewMsg.Option2(1).Enabled = False
  NewMsg.Option2(7).Enabled = False
 'MsgBox ("Please...Complete Data")
 Else
 NewMsg.Option2(1).Enabled = True
   SqLst = "DELETE * FROM Notes "
        DBEngine.Workspaces(0).BeginTrans
        db.Execute SqLst
        DBEngine.Workspaces(0).CommitTrans
    Do While Not rs1.EOF
   A = rs1("CompanyId")
   Set rs = db.OpenRecordset("Notes")
   rs.AddNew
   rs("CompanyId") = A
    rs("Notes") = NewMsg.txtNoteText.Text
    rs("Supject") = NewMsg.txtsubject.Text
   rs.Update
    rs1.MoveNext
    Loop
    End If
End Function



