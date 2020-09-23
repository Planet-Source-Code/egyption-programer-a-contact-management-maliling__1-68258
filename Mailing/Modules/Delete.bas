Attribute VB_Name = "MDelete"
Dim db As Database
Dim ws As Workspace
Dim rs, rs1 As Recordset
Dim SqLst As String
Dim StrMessage As String
Dim X, y, d As String
Public Function DelRec(ScrName As String)
If Right(App.Path, 1) = "\" Then
            PathPro = App.Path & "Mailing.mdb"
        Else
            PathPro = App.Path & "\Mailing.mdb"
        End If

On Error Resume Next
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
Select Case ScrName
Case "frmMainData"
    X = frmMainData.StbMainData.Tab
    Select Case X
     Case 0
        SqLst = "Select CategoryID From Qcatcomp"
            SqLst = SqLst & " WHERE CategoryID = " & frmMainData.txtCategorycode.Text & ""
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
                MsgBox ("There are Companies joined  by this Category")
           Else
      SqLst = "Select CategoryID From QProduct"
            SqLst = SqLst & " WHERE CategoryID = " & frmMainData.txtCategorycode.Text & ""
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
                MsgBox ("There are Products joined  by this Category")
       Else
        SqLst = "DELETE * FROM Category "
        
        SqLst = SqLst & " WHERE CategoryID = " & frmMainData.txtCategorycode.Text & ""
        
        StrMessage = "are you sure?"
        DBEngine.Workspaces(0).BeginTrans
        
        If MsgBox(StrMessage, vbYesNo) = vbYes Then
        db.Execute SqLst
        
            DBEngine.Workspaces(0).CommitTrans
             display.display (frmMainData.Name)
        Else
            DBEngine.Workspaces(0).Rollback
        End If
        End If
        End If
    Case 1
        SqLst = "Select ProductID From procomp"
            SqLst = SqLst & " WHERE ProductID = " & frmMainData.txtProductID.Text & ""
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
                MsgBox ("There are Companies joined  by this Product")
            Else
        SqLst = "DELETE * FROM Product "
        SqLst = SqLst & " WHERE ProductID = " & frmMainData.txtProductID.Text & ""
        StrMessage = "are you sure?"
        DBEngine.Workspaces(0).BeginTrans
        
        If MsgBox(StrMessage, vbYesNo) = vbYes Then
        db.Execute SqLst
            DBEngine.Workspaces(0).CommitTrans
            display.display (frmMainData.Name)
        Else
            DBEngine.Workspaces(0).Rollback
        End If
        End If
        
        
    Case 2
       SqLst = "Select citycode From Company"
            SqLst = SqLst & " WHERE citycode = " & frmMainData.txtcitycode.Text & ""
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
                MsgBox ("There are Companies joined  by this Region")
            Else
        SqLst = "DELETE * FROM city "
        
        SqLst = SqLst & " WHERE citycode = " & frmMainData.txtcitycode.Text & ""
        
        StrMessage = "are you sure?"
        DBEngine.Workspaces(0).BeginTrans
        
        If MsgBox(StrMessage, vbYesNo) = vbYes Then
        db.Execute SqLst
            DBEngine.Workspaces(0).CommitTrans
            display.display (frmMainData.Name)
        Else
            DBEngine.Workspaces(0).Rollback
        End If
        End If
    Case 3
       SqLst = "Select TypeId From comtyp"
            SqLst = SqLst & " WHERE TypeId = " & frmMainData.txtTypeID.Text & ""
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
                MsgBox ("There are Companies joined  by this Type")
            Else
        SqLst = "DELETE * FROM Type "
        SqLst = SqLst & " WHERE TypeId = " & frmMainData.txtTypeID.Text & ""
        StrMessage = "Are you sure?"
        DBEngine.Workspaces(0).BeginTrans
        
        If MsgBox(StrMessage, vbYesNo) = vbYes Then
        db.Execute SqLst
            DBEngine.Workspaces(0).CommitTrans
            display.display (frmMainData.Name)
        Else
            DBEngine.Workspaces(0).Rollback
        End If
   End If
   Case 4
       SqLst = "Select TypeId From pertype"
            SqLst = SqLst & " WHERE posTypeId = " & frmMainData.txtTposID.Text & ""
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
                MsgBox ("There are Companies joined  by this Type")
            Else
        SqLst = "DELETE * FROM PosType "
        SqLst = SqLst & " WHERE posTypeId = " & frmMainData.txtTposID.Text & ""
        StrMessage = "Are you sure?"
        DBEngine.Workspaces(0).BeginTrans
        
        If MsgBox(StrMessage, vbYesNo) = vbYes Then
        db.Execute SqLst
            DBEngine.Workspaces(0).CommitTrans
            display.display (frmMainData.Name)
        Else
            DBEngine.Workspaces(0).Rollback
        End If
   End If
      End Select
Case "frmCampany"
 If frmCampany.txtcompanyID.Text <> "" Then
    
    SqLst = "DELETE * FROM Company "
        SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
        StrMessage = "are you sure?"
        DBEngine.Workspaces(0).BeginTrans
        
        If MsgBox(StrMessage, vbYesNo) = vbYes Then
        db.Execute SqLst
        DBEngine.Workspaces(0).CommitTrans
            display.display (frmCampany.Name)
        Else
            DBEngine.Workspaces(0).Rollback
        End If
         End If
Case "frmpersonal"
 If frmpersonal.txtPersonalID.Text <> "" Then
    
    SqLst = "DELETE * FROM Personal "
        SqLst = SqLst & " WHERE personalId = " & frmpersonal.txtPersonalID.Text & ""
        StrMessage = "are you sure?"
        DBEngine.Workspaces(0).BeginTrans
        
        If MsgBox(StrMessage, vbYesNo) = vbYes Then
        db.Execute SqLst
        DBEngine.Workspaces(0).CommitTrans
            display.display (frmpersonal.Name)
        Else
            DBEngine.Workspaces(0).Rollback
        End If
         End If


End Select
End Function


