Attribute VB_Name = "Save"
Dim db As Database
Dim ws As Workspace
Dim rs As Recordset
Dim rs1 As Recordset
Dim rs2 As Recordset
Dim SqLst As String
Dim SqLst1 As String
Dim d, stdate, endate As Date
Dim A As Integer
Dim y As Integer
Dim X

 Public Function add_edit(ScrName As String)

'On Error GoTo eh
If Right(App.Path, 1) = "\" Then
            PathPro = App.Path & "Mailing.mdb"
        Else
            PathPro = App.Path & "\Mailing.mdb"
        End If
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
Select Case ScrName
Case "frmMainData"
    X = frmMainData.StbMainData.Tab
    Select Case X
    Case 0
    If frmMainData.txtCategorycode.Text = "" Or frmMainData.txtCategoryname.Text = "" Then
    MsgBox ("Please...Complete Data")
    GoTo en1
    End If
            SqLst = "Select * From Category"
            SqLst = SqLst & " WHERE Categoryname = '" & frmMainData.txtCategoryname.Text
            SqLst = SqLst & "' ORDER BY CategoryID "
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
                MsgBox ("Mistake...Duplicate Data")
            Else
            SqLst = "Select * From Category"
            SqLst = SqLst & " WHERE CategoryID = " & frmMainData.txtCategorycode.Text
            SqLst = SqLst & " ORDER BY CategoryID "
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
                rs.Edit
                rs("Categoryname") = frmMainData.txtCategoryname.Text
                rs.Update
            Else
                rs.AddNew
                    rs("Categoryname") = frmMainData.txtCategoryname.Text
                    rs("CategoryID") = frmMainData.txtCategorycode.Text
                rs.Update
            End If
        End If
        frmMainData.lstcat.Clear
    Set rs = db.OpenRecordset("Category")
    Do While Not rs.EOF
        frmMainData.lstcat.AddItem rs("Categoryname")
    rs.MoveNext
    Loop
     Case 1
   If frmMainData.txtProductID.Text = "" Or frmMainData.txtProductname.Text = "" Or _
   frmMainData.CmbCategoryname.Text = "" Then
   MsgBox ("Please...Complete Data")
   GoTo en1
    End If
         
           SqLst1 = "Select CategoryID From Category"
                    SqLst1 = SqLst1 & " WHERE Categoryname = '" & frmMainData.CmbCategoryname.Text & "'"
                    Set rs1 = db.OpenRecordset(SqLst1)
                    A = rs1("CategoryID")
            SqLst = "Select * From Product"
            SqLst = SqLst & " WHERE Productname = '" & frmMainData.txtProductname.Text
            SqLst = SqLst & "' ORDER BY ProductID "
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
                MsgBox ("Mistake...Duplicate Data")
            Else
            SqLst = "Select * From Product"
            SqLst = SqLst & " WHERE ProductID = " & frmMainData.txtProductID.Text
            SqLst = SqLst & " ORDER BY ProductID "
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
                rs.Edit
                 rs("CategoryID") = A
                rs("Productname") = frmMainData.txtProductname.Text
                rs.Update
            Else
                rs.AddNew
                    rs("ProductID") = frmMainData.txtProductID.Text
                     rs("CategoryID") = A
                rs("Productname") = frmMainData.txtProductname.Text
                rs.Update
            End If
        End If
         SqLst = "Select distinct Productname From QProduct"
            SqLst = SqLst & " WHERE Categoryname ='" & frmMainData.CmbCategoryname.Text & "'"
            Set rs = db.OpenRecordset(SqLst)
        frmMainData.lstprod.Clear
        Do While Not rs.EOF
            frmMainData.lstprod.AddItem rs("Productname")
            rs.MoveNext
        Loop
         Case 2
    If frmMainData.txtcitycode.Text = "" Or frmMainData.cmbCity.Text = "" _
    Or frmMainData.txtRegion.Text = "" Or frmMainData.Cmbcountry.Text = "" Then
    MsgBox ("Please...Complete Data")
    GoTo en1
    End If
            SqLst = "Select * From City"
            SqLst = SqLst & " WHERE Region = '" & frmMainData.txtRegion.Text
            SqLst = SqLst & "'and country = '" & frmMainData.Cmbcountry.Text & "'and city = '" & frmMainData.cmbCity.Text & "'"
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
                MsgBox ("Mistake...Duplicate Data")
            Else
            SqLst = "Select * From City"
            SqLst = SqLst & " WHERE citycode = " & frmMainData.txtcitycode.Text
            SqLst = SqLst & " ORDER BY citycode "
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
                rs.Edit
                rs("city") = frmMainData.cmbCity.Text
                rs("country") = frmMainData.Cmbcountry.Text
                rs("Region") = frmMainData.txtRegion.Text
                rs.Update
            Else
            SqLst = "Select * From City"
            SqLst = SqLst & " WHERE city = '" & frmMainData.cmbCity.Text
            SqLst = SqLst & "' and Region = '" & frmMainData.txtRegion.Text
           SqLst = SqLst & "' and country = '" & frmMainData.Cmbcountry.Text & "' ORDER BY citycode "
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
            MsgBox ("Mistake...Duplicate Data")
            Else
                rs.AddNew
                  rs("citycode") = frmMainData.txtcitycode.Text
                rs("city") = frmMainData.cmbCity.Text
                rs("country") = frmMainData.Cmbcountry.Text
                rs("Region") = frmMainData.txtRegion.Text
                rs.Update
            End If
        End If
        End If
        SqLst1 = "Select distinct city From City"
        SqLst1 = SqLst1 & " WHERE country = '" & frmMainData.Cmbcountry.Text & "'"
        Set rs1 = db.OpenRecordset(SqLst1)
            
        frmMainData.lstcity.Clear
        frmMainData.cmbCity.Clear
        Do While Not rs1.EOF
        frmMainData.lstcity.AddItem rs1("city")
        frmMainData.cmbCity.AddItem rs1("city")
        rs1.MoveNext
        Loop
        Moving.MoveLast (frmMainData.Name)
        SqLst1 = "Select distinct Region From City"
        SqLst1 = SqLst1 & " WHERE city = '" & frmMainData.cmbCity.Text & "'"
        Set rs1 = db.OpenRecordset(SqLst1)
        frmMainData.lstcity.Clear
        Do While Not rs1.EOF
        frmMainData.lstcity.AddItem rs1("Region")
        rs1.MoveNext
        Loop
        frmMainData.lstcity.ListIndex = 0
         Case 3
    If frmMainData.txtTypeID.Text = "" Or frmMainData.txtTypeName.Text = "" Then
    MsgBox ("Please...Complete Data")
    GoTo en1
    End If
            SqLst = "Select * From Type"
            SqLst = SqLst & " WHERE Typename = '" & frmMainData.txtTypeName.Text
            SqLst = SqLst & "' ORDER BY TypeId "
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
            MsgBox ("Mistake...Duplicate Data")
            Else
            SqLst = "Select * From Type"
            SqLst = SqLst & " WHERE TypeId = " & frmMainData.txtTypeID.Text
            SqLst = SqLst & " ORDER BY TypeId "
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
                rs.Edit
                rs("Typename") = frmMainData.txtTypeName.Text
                rs.Update
            Else
                rs.AddNew
                    rs("Typename") = frmMainData.txtTypeName.Text
                    rs("TypeId") = frmMainData.txtTypeID.Text
                rs.Update
            End If
        End If
        frmMainData.lstype.Clear
    Set rs = db.OpenRecordset("Type")
    Do While Not rs.EOF
        frmMainData.lstype.AddItem rs("Typename")
    rs.MoveNext
    Loop
        frmMainData.cmdCmovenext.Enabled = True
    frmMainData.cmdCmoveprevious.Enabled = True
    frmMainData.cmdCMovelast.Enabled = True
    frmMainData.cmdCmovefrist.Enabled = True
     Case 4
    If frmMainData.txtTposID.Text = "" Or frmMainData.txtTposname.Text = "" Then
    MsgBox ("Please...Complete Data")
    GoTo en1
    End If
            SqLst = "Select * From PosType"
            SqLst = SqLst & " WHERE posTypename = '" & frmMainData.txtTposname.Text
            SqLst = SqLst & "' ORDER BY posTypeId "
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
            MsgBox ("Mistake...Duplicate Data")
            Else
            SqLst = "Select * From PosType"
            SqLst = SqLst & " WHERE posTypeId = " & frmMainData.txtTposID.Text
            SqLst = SqLst & " ORDER BY posTypeId "
            Set rs = db.OpenRecordset(SqLst)
            If rs.RecordCount <> 0 Then
                rs.Edit
                rs("posTypename") = frmMainData.txtTposname.Text
                rs.Update
            Else
                rs.AddNew
                   rs("posTypeId") = frmMainData.txtTposID.Text
                    rs("posTypename") = frmMainData.txtTposname.Text
                rs.Update
            End If
        End If
        frmMainData.lstpostype.Clear
    Set rs = db.OpenRecordset("PosType")
    Do While Not rs.EOF
        frmMainData.lstpostype.AddItem rs("posTypename")
    rs.MoveNext
    Loop
    frmMainData.cmdCmovenext.Enabled = True
    frmMainData.cmdCmoveprevious.Enabled = True
    frmMainData.cmdCMovelast.Enabled = True
    frmMainData.cmdCmovefrist.Enabled = True
End Select
Case "frmpersonal"
    
    If frmpersonal.txtPersonalID.Text = "" Or frmpersonal.cmb_company.Text = "" Or _
frmpersonal.cmb_titel.Text = "" Or frmpersonal.txt_name.Text = "" Or frmpersonal.txt_pos.Text = "" Then
MsgBox ("Please...Complete Data")
    GoTo en1
    End If
 
     SqLst = "Select * From personal"
        SqLst = SqLst & " WHERE personalId = " & frmpersonal.txtPersonalID.Text & ""
        'SqLst = SqLst & " and Companyname = '" & frmCampany.txtCompanyName.Text & "' ORDER BY CompanyId "
        Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
        'SqLst1 = "Select * From Personal"
'SqLst1 = SqLst1 & " WHERE name = '" & frmpersonal.txt_name.Text
'SqLst1 = SqLst1 & "' ORDER BY personalId "
'Set rs1 = db.OpenRecordset(SqLst1)
'If rs1.RecordCount <> 0 Then

'MsgBox ("Mistake...Duplicate Data")
'GoTo en1
'End If
rs.Edit
rs("name") = frmpersonal.txt_name.Text

 rs("Title") = frmpersonal.cmb_titel.Text
rs("Position") = frmpersonal.txt_pos.Text
'frmpersonal.cmb_company.Text = rs("Companyname")
        
y = 0
SqLst = "DELETE * FROM pertype "
        SqLst = SqLst & " WHERE personalId = " & frmpersonal.txtPersonalID.Text & ""
        DBEngine.Workspaces(0).BeginTrans
        db.Execute SqLst
        DBEngine.Workspaces(0).CommitTrans
        
Do While y < frmpersonal.lst_typ.ListCount
frmpersonal.lst_typ.ListIndex = y
If frmpersonal.lst_typ.Selected(frmpersonal.lst_typ.ListIndex) = True Then

    SqLst1 = "Select posTypeId From PosType"
    SqLst1 = SqLst1 & " WHERE posTypename = '" & frmpersonal.lst_typ.List(frmpersonal.lst_typ.ListIndex) & "'"
    Set rs1 = db.OpenRecordset(SqLst1)
    A = rs1("posTypeId")
    
        
    
    SqLst = "Select * From pertype"
    SqLst = SqLst & " WHERE personalId = " & frmpersonal.txtPersonalID.Text & ""
    SqLst = SqLst & " and posTypeId = " & A & " ORDER BY personalId "
    Set rs2 = db.OpenRecordset(SqLst)
    If rs2.RecordCount = 0 Then
        rs2.AddNew
        rs2("posTypeId") = A
        rs2("personalId") = frmpersonal.txtPersonalID.Text
        rs2.Update
    End If
    Else
   
    
    End If
     y = y + 1
Loop
y = 0

            SqLst1 = "Select CompanyId From Company"
            SqLst1 = SqLst1 & " WHERE Companyname = '" & frmpersonal.cmb_company.Text & "'"
           
            Set rs1 = db.OpenRecordset(SqLst1)
            A = rs1("CompanyId")
            rs("CompanyId") = A
            If frmpersonal.txt_tel.Text = "" Then
               ' rs("Tel1") = "0"
            Else
                rs("Tel1") = frmpersonal.txt_tel.Text
            End If
            If frmpersonal.chkpos.Value = 0 Then
             rs("showPos") = "No"
            Else
            'frmpersonal.chkpos.Value = 1
             rs("showPos") = "Yes"
            End If
            
            If frmpersonal.txt_fax.Text = "" Then
              '  rs("fax1") = "0"
            Else
                rs("fax1") = frmpersonal.txt_fax.Text
            End If
            If frmpersonal.txt_notes.Text = "" Then
                rs("Notes") = "Empty"
            Else
               rs("Notes") = frmpersonal.txt_notes.Text
            End If
            If frmpersonal.txt_mail.Text = "" Then
                rs("Email") = "Empty"
            Else
                rs("Email") = frmpersonal.txt_mail.Text
            End If
                           
            If frmpersonal.txt_mobile.Text = "" Then
              '  rs("Tel2") = "0"
            Else
                 rs("Tel2") = frmpersonal.txt_mobile.Text
            End If
rs.Update
Else
SqLst = "Select * From Personal"
SqLst = SqLst & " WHERE name = '" & frmpersonal.txt_name.Text
SqLst = SqLst & "' ORDER BY personalId "
Set rs = db.OpenRecordset(SqLst)
If rs.RecordCount <> 0 Then

MsgBox ("Mistake...Duplicate Data")
'frmCampany.cmbCompanyname.Text = frmCampany.txtCompanyName.Text
'frmCampany.search
Else
rs.AddNew
rs("personalId") = frmpersonal.txtPersonalID.Text
rs("name") = frmpersonal.txt_name.Text

 rs("Title") = frmpersonal.cmb_titel.Text
rs("Position") = frmpersonal.txt_pos.Text
'frmpersonal.cmb_company.Text = rs("Companyname")
        

            SqLst1 = "Select CompanyId From Company"
            SqLst1 = SqLst1 & " WHERE Companyname = '" & frmpersonal.cmb_company.Text & "'"
           
            Set rs1 = db.OpenRecordset(SqLst1)
            A = rs1("CompanyId")
            rs("CompanyId") = A
            If frmpersonal.txt_tel.Text = "" Then
              '  rs("Tel1") = "0"
            Else
                rs("Tel1") = frmpersonal.txt_tel.Text
            End If
            If frmpersonal.txt_fax.Text = "" Then
              '  rs("fax1") = "0"
            Else
                rs("fax1") = frmpersonal.txt_fax.Text
            End If
            If frmpersonal.chkpos.Value = 0 Then
             rs("showPos") = "No"
            Else
            'frmpersonal.chkpos.Value = 1
             rs("showPos") = "Yes"
            End If
            If frmpersonal.txt_notes.Text = "" Then
                rs("Notes") = "Empty"
            Else
               rs("Notes") = frmpersonal.txt_notes.Text
            End If
            If frmpersonal.txt_mail.Text = "" Then
                rs("Email") = "Empty"
            Else
                rs("Email") = frmpersonal.txt_mail.Text
            End If
                           
            If frmpersonal.txt_mobile.Text = "" Then
              '  rs("Tel2") = "0"
            Else
                 rs("Tel2") = frmpersonal.txt_mobile.Text
            End If
rs.Update
y = 0
SqLst = "DELETE * FROM pertype "
        SqLst = SqLst & " WHERE personalId = " & frmpersonal.txtPersonalID.Text & ""
        DBEngine.Workspaces(0).BeginTrans
        db.Execute SqLst
        DBEngine.Workspaces(0).CommitTrans
        
Do While y < frmpersonal.lst_typ.ListCount
frmpersonal.lst_typ.ListIndex = y
If frmpersonal.lst_typ.Selected(frmpersonal.lst_typ.ListIndex) = True Then

    SqLst1 = "Select posTypeId From PosType"
    SqLst1 = SqLst1 & " WHERE posTypename = '" & frmpersonal.lst_typ.List(frmpersonal.lst_typ.ListIndex) & "'"
    Set rs1 = db.OpenRecordset(SqLst1)
    A = rs1("posTypeId")
    
        
    
    SqLst = "Select * From pertype"
    SqLst = SqLst & " WHERE personalId = " & frmpersonal.txtPersonalID.Text & ""
    SqLst = SqLst & " and posTypeId = " & A & " ORDER BY personalId "
    Set rs2 = db.OpenRecordset(SqLst)
    If rs2.RecordCount = 0 Then
        rs2.AddNew
        rs2("posTypeId") = A
        rs2("personalId") = frmpersonal.txtPersonalID.Text
        rs2.Update
    End If
    Else
   
    
    End If
     y = y + 1
Loop
'y = 0

End If
End If
PersonalDis
frmpersonal.cmdCmovenext.Enabled = True
frmpersonal.cmdCmoveprevious.Enabled = True
frmpersonal.cmdCmovefrist.Enabled = True
frmpersonal.cmdCMovelast.Enabled = True
'******************************************************

    Case "frmCampany"
   If frmCampany.txtcompanyID.Text = "" Or frmCampany.txtCompanyName.Text = "" Or _
frmCampany.cmbCity.Text = "" Or frmCampany.Cmbcountry.Text = "" Or frmCampany.txttel1.Text = "" Then
MsgBox ("Please...Complete Data")
    GoTo en1
    End If
 
     SqLst = "Select * From company"
        SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
        'SqLst = SqLst & " and Companyname = '" & frmCampany.txtCompanyName.Text & "' ORDER BY CompanyId "
        Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
rs.Edit
rs("Companyname") = frmCampany.txtCompanyName.Text
rs("addno") = frmCampany.txt_addno.Text
rs("Address") = frmCampany.txtaddress.Text
y = 0
SqLst = "DELETE * FROM catcomp "
        SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
        DBEngine.Workspaces(0).BeginTrans
        db.Execute SqLst
        DBEngine.Workspaces(0).CommitTrans
        
Do While y < frmCampany.Lstcategory.ListCount
frmCampany.Lstcategory.ListIndex = y
If frmCampany.Lstcategory.Selected(frmCampany.Lstcategory.ListIndex) = True Then

    SqLst1 = "Select CategoryID From Category"
    SqLst1 = SqLst1 & " WHERE Categoryname = '" & frmCampany.Lstcategory.List(frmCampany.Lstcategory.ListIndex) & "'"
    Set rs1 = db.OpenRecordset(SqLst1)
    A = rs1("CategoryID")
    
        
    
    SqLst = "Select * From catcomp"
    SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
    SqLst = SqLst & " and CategoryID = " & A & " ORDER BY CompanyId "
    Set rs2 = db.OpenRecordset(SqLst)
    If rs2.RecordCount = 0 Then
        rs2.AddNew
        rs2("CategoryID") = A
        rs2("CompanyId") = frmCampany.txtcompanyID.Text
        rs2.Update
    End If
    Else
   
    
    End If
     y = y + 1
Loop
y = 0
SqLst = "DELETE * FROM comtyp "
        SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
        DBEngine.Workspaces(0).BeginTrans
        db.Execute SqLst
        DBEngine.Workspaces(0).CommitTrans
        
Do While y < frmCampany.LstType.ListCount
frmCampany.LstType.ListIndex = y
If frmCampany.LstType.Selected(frmCampany.LstType.ListIndex) = True Then

    SqLst1 = "Select TypeId From Type"
    SqLst1 = SqLst1 & " WHERE Typename = '" & frmCampany.LstType.List(frmCampany.LstType.ListIndex) & "'"
    Set rs1 = db.OpenRecordset(SqLst1)
    A = rs1("TypeId")
    
        
    
    SqLst = "Select * From comtyp"
    SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
    SqLst = SqLst & " and TypeId = " & A & " ORDER BY CompanyId "
    Set rs2 = db.OpenRecordset(SqLst)
    If rs2.RecordCount = 0 Then
        rs2.AddNew
        rs2("TypeId") = A
        rs2("CompanyId") = frmCampany.txtcompanyID.Text
        rs2.Update
    End If
    Else
    End If
     y = y + 1
Loop
y = 0
SqLst = "DELETE * FROM procomp "
        SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
        DBEngine.Workspaces(0).BeginTrans
        db.Execute SqLst
        DBEngine.Workspaces(0).CommitTrans
        
Do While y < frmCampany.LstProduct.ListCount
frmCampany.LstProduct.ListIndex = y
If frmCampany.LstProduct.Selected(frmCampany.LstProduct.ListIndex) = True Then

    SqLst1 = "Select ProductID From Product"
    SqLst1 = SqLst1 & " WHERE Productname = '" & frmCampany.LstProduct.List(frmCampany.LstProduct.ListIndex) & "'"
    Set rs1 = db.OpenRecordset(SqLst1)
    A = rs1("ProductID")
    
        
    
    SqLst = "Select * From procomp"
    SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
    SqLst = SqLst & " and ProductID = " & A & " ORDER BY CompanyId "
    Set rs2 = db.OpenRecordset(SqLst)
    If rs2.RecordCount = 0 Then
        rs2.AddNew
        rs2("ProductID") = A
        rs2("CompanyId") = frmCampany.txtcompanyID.Text
        rs2.Update
    End If
    Else
   
    
    End If
     y = y + 1
Loop
            SqLst1 = "Select * From city"
            SqLst1 = SqLst1 & " WHERE Region = '" & frmCampany.CmbRegion.Text
            SqLst1 = SqLst1 & "' and city = '" & frmCampany.cmbCity.Text
            SqLst1 = SqLst1 & "' and country = '" & frmCampany.Cmbcountry.Text & "'"
            Set rs1 = db.OpenRecordset(SqLst1)
            A = rs1("citycode")
            rs("citycode") = A
             rs("PostalCode") = frmCampany.txtPostelCode.Text
            'End If
           ' If frmCampany.txtPbox.Text = "" Then
         ' frmCampany.txtPbox.Text = ""
         '   Else
                rs("PoBox") = frmCampany.txtPbox.Text
         '   End If
         '   If frmCampany.txttel1.Text = "" Then
              '  rs("tel1") = 0
         '   Else
                rs("tel1") = frmCampany.txttel1.Text
         '   End If
         '   If frmCampany.txttel2.Text = "" Then
              '  frmCampany.txttel2.Text = 0
         '   Else
                rs("tel2") = frmCampany.txttel2.Text
         '   End If
         '   If frmCampany.txttel3.Text = "" Then
              '  rs("tel3") = 0
         '   Else
                rs("tel3") = frmCampany.txttel3.Text
         '   End If
         '   If frmCampany.txttel4.Text = "" Then
              '  rs("tel4") = 0
         '   Else
               rs("tel4") = frmCampany.txttel4.Text
         '   End If
           ' If frmCampany.txttel5.Text = "" Then
              '  rs("tel5") = 0
           ' Else
                rs("tel5") = frmCampany.txttel5.Text
           ' End If
           ' If frmCampany.txtFax1.Text = "" Then
              '  rs("Fax1") = 0
           ' Else
                rs("Fax1") = frmCampany.txtFax1.Text
           ' End If
           ' If frmCampany.txtFax2.Text = "" Then
              '  rs("Fax2") = 0
           ' Else
                rs("Fax2") = frmCampany.txtFax2.Text
           ' End If
           
           
            If frmCampany.txtEmail.Text = "" Then
                rs("EMail") = "Empty"
            Else
             rs("EMail") = frmCampany.txtEmail.Text
            End If
             If frmCampany.txtEmail2.Text = "" Then
                rs("Email2") = "Empty"
            Else
                rs("Email2") = frmCampany.txtEmail2.Text
            End If
            If frmCampany.txtweb.Text = "" Then
                rs("web") = "Empty"
            Else
                rs("web") = frmCampany.txtweb.Text
            End If
            
            
  If frmCampany.txt_addno.Text = "" Then
              '  rs("addno") = "Empty"
            Else
                rs("addno") = frmCampany.txt_addno.Text
            End If
            
            
            
            
           
    
              If frmCampany.txtNots.Text = "" Then
                rs("Notes") = "Empty"
            Else
                rs("Notes") = frmCampany.txtNots.Text
            End If
             
            If frmCampany.txtAbbreviation.Text = "" Then
                rs("Abbreviation") = "Empty"
            Else
                rs("Abbreviation") = frmCampany.txtAbbreviation.Text
            End If
rs.Update
Else
SqLst = "Select * From Company"
SqLst = SqLst & " WHERE Companyname = '" & frmCampany.txtCompanyName.Text
SqLst = SqLst & "' ORDER BY CompanyId "
Set rs = db.OpenRecordset(SqLst)
If rs.RecordCount <> 0 Then

MsgBox ("Mistake...Duplicate Data")
frmCampany.cmbCompanyname.Text = frmCampany.txtCompanyName.Text
frmCampany.search
Else

rs.AddNew
rs("CompanyId") = frmCampany.txtcompanyID.Text

rs("Companyname") = frmCampany.txtCompanyName.Text

rs("Addno") = frmCampany.txt_addno.Text

rs("Address") = frmCampany.txtaddress.Text

            SqLst1 = "Select * From city"
            SqLst1 = SqLst1 & " WHERE Region = '" & frmCampany.CmbRegion.Text
            SqLst1 = SqLst1 & "' and city = '" & frmCampany.cmbCity.Text
            SqLst1 = SqLst1 & "' and country = '" & frmCampany.Cmbcountry.Text & "'"
            Set rs1 = db.OpenRecordset(SqLst1)
            A = rs1("citycode")
            rs("citycode") = A
            'If frmCampany.txtPostelCode.Text = "" Then
            '    rs("PostalCode") = 0
            'Else
                rs("PostalCode") = frmCampany.txtPostelCode.Text
            'End If
           ' If frmCampany.txtPbox.Text = "" Then
         ' frmCampany.txtPbox.Text = ""
         '   Else
                rs("PoBox") = frmCampany.txtPbox.Text
         '   End If
         '   If frmCampany.txttel1.Text = "" Then
              '  rs("tel1") = 0
         '   Else
                rs("tel1") = frmCampany.txttel1.Text
         '   End If
         '   If frmCampany.txttel2.Text = "" Then
              '  frmCampany.txttel2.Text = 0
         '   Else
                rs("tel2") = frmCampany.txttel2.Text
         '   End If
         '   If frmCampany.txttel3.Text = "" Then
              '  rs("tel3") = 0
         '   Else
                rs("tel3") = frmCampany.txttel3.Text
         '   End If
         '   If frmCampany.txttel4.Text = "" Then
              '  rs("tel4") = 0
         '   Else
               rs("tel4") = frmCampany.txttel4.Text
         '   End If
           ' If frmCampany.txttel5.Text = "" Then
              '  rs("tel5") = 0
           ' Else
                rs("tel5") = frmCampany.txttel5.Text
           ' End If
           ' If frmCampany.txtFax1.Text = "" Then
              '  rs("Fax1") = 0
           ' Else
                rs("Fax1") = frmCampany.txtFax1.Text
           ' End If
           ' If frmCampany.txtFax2.Text = "" Then
              '  rs("Fax2") = 0
           ' Else
                rs("Fax2") = frmCampany.txtFax2.Text
           ' End If
           
           
            If frmCampany.txtEmail.Text = "" Then
                rs("EMail") = "Empty"
            Else
             rs("EMail") = frmCampany.txtEmail.Text
            End If
             If frmCampany.txtEmail2.Text = "" Then
                rs("Email2") = "Empty"
            Else
                rs("Email2") = frmCampany.txtEmail2.Text
            End If
            If frmCampany.txtweb.Text = "" Then
                rs("web") = "Empty"
            Else
                rs("web") = frmCampany.txtweb.Text
            End If
          
              If frmCampany.txtNots.Text = "" Then
                rs("Notes") = "Empty"
            Else
                rs("Notes") = frmCampany.txtNots.Text
            End If
             
            If frmCampany.txtAbbreviation.Text = "" Then
                rs("Abbreviation") = "Empty"
            Else
                rs("Abbreviation") = frmCampany.txtAbbreviation.Text
            End If
rs.Update
y = 0
SqLst = "DELETE * FROM catcomp "
        SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
        DBEngine.Workspaces(0).BeginTrans
        db.Execute SqLst
        DBEngine.Workspaces(0).CommitTrans
        
Do While y < frmCampany.Lstcategory.ListCount
frmCampany.Lstcategory.ListIndex = y
If frmCampany.Lstcategory.Selected(frmCampany.Lstcategory.ListIndex) = True Then

    SqLst1 = "Select CategoryID From Category"
    SqLst1 = SqLst1 & " WHERE Categoryname = '" & frmCampany.Lstcategory.List(frmCampany.Lstcategory.ListIndex) & "'"
    Set rs1 = db.OpenRecordset(SqLst1)
    A = rs1("CategoryID")
       SqLst = "Select * From catcomp"
    SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
    SqLst = SqLst & " and CategoryID = " & A & " ORDER BY CompanyId "
    Set rs2 = db.OpenRecordset(SqLst)
    If rs2.RecordCount = 0 Then
        rs2.AddNew
        rs2("CategoryID") = A
        rs2("CompanyId") = frmCampany.txtcompanyID.Text
        rs2.Update
    End If
    Else
   
    End If
     y = y + 1
Loop
y = 0
SqLst = "DELETE * FROM comtyp "
        SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
        DBEngine.Workspaces(0).BeginTrans
        db.Execute SqLst
        DBEngine.Workspaces(0).CommitTrans
        
Do While y < frmCampany.LstType.ListCount
frmCampany.LstType.ListIndex = y
If frmCampany.LstType.Selected(frmCampany.LstType.ListIndex) = True Then

    SqLst1 = "Select TypeId From Type"
    SqLst1 = SqLst1 & " WHERE Typename = '" & frmCampany.LstType.List(frmCampany.LstType.ListIndex) & "'"
    Set rs1 = db.OpenRecordset(SqLst1)
    A = rs1("TypeId")
    
        
    
    SqLst = "Select * From comtyp"
    SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
    SqLst = SqLst & " and TypeId = " & A & " ORDER BY CompanyId "
    Set rs2 = db.OpenRecordset(SqLst)
    If rs2.RecordCount = 0 Then
        rs2.AddNew
        rs2("TypeId") = A
        rs2("CompanyId") = frmCampany.txtcompanyID.Text
        rs2.Update
    End If
    Else
    End If
     y = y + 1
Loop
y = 0
SqLst = "DELETE * FROM procomp "
        SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
        DBEngine.Workspaces(0).BeginTrans
        db.Execute SqLst
        DBEngine.Workspaces(0).CommitTrans
        
Do While y < frmCampany.LstProduct.ListCount
frmCampany.LstProduct.ListIndex = y
If frmCampany.LstProduct.Selected(frmCampany.LstProduct.ListIndex) = True Then

    SqLst1 = "Select ProductID From Product"
    SqLst1 = SqLst1 & " WHERE Productname = '" & frmCampany.LstProduct.List(frmCampany.LstProduct.ListIndex) & "'"
    Set rs1 = db.OpenRecordset(SqLst1)
    A = rs1("ProductID")
    
        
    
    SqLst = "Select * From procomp"
    SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
    SqLst = SqLst & " and ProductID = " & A & " ORDER BY CompanyId "
    Set rs2 = db.OpenRecordset(SqLst)
    If rs2.RecordCount = 0 Then
        rs2.AddNew
        rs2("ProductID") = A
        rs2("CompanyId") = frmCampany.txtcompanyID.Text
        rs2.Update
    End If
    Else
   
    
    End If
     y = y + 1
Loop
End If
End If

frmCampany.cmdCmovenext.Enabled = True
frmCampany.cmdCmoveprevious.Enabled = True
frmCampany.cmdCmovefrist.Enabled = True
frmCampany.cmdCMovelast.Enabled = True
frmCampany.Command1.Enabled = True

End Select
eh:

 If Err.Number = 3022 Then
 MsgBox ("Mistake...Duplicate Name")
' ElseIf Err.Number <> 3022 Then
' MsgBox ("Mistake...")
End If
en1:
savemenu
End Function
Public Sub PersonalDis()
frmpersonal.txt_name.SetFocus

        SqLst = "Select posTypename From Qper"
SqLst = SqLst & " WHERE personalId = " & frmpersonal.txtPersonalID.Text & ""
frmpersonal.lst_typ.Clear
Set rs = db.OpenRecordset(SqLst)
y = 0
Do While Not rs.EOF
frmpersonal.lst_typ.AddItem rs("posTypename")
frmpersonal.lst_typ.Selected(y) = True
y = y + 1
rs.MoveNext


Loop
SqLst = "Select name From Qpersonal"
SqLst = SqLst & " WHERE Companyname = '" & frmpersonal.cmb_company.Text & "'"
frmpersonal.lst_person.Clear
Set rs = db.OpenRecordset(SqLst)

Do While Not rs.EOF
frmpersonal.lst_person.AddItem rs("name")
rs.MoveNext
Loop
       
End Sub
Public Sub savemenu()
frmMain.MP.Enabled = True
frmMain.mn.Enabled = True
frmMain.ML.Enabled = True
frmMain.MF.Enabled = True
End Sub

