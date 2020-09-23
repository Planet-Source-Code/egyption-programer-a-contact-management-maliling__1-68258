Attribute VB_Name = "Moving"
Dim db As Database
Dim ws As Workspace
Dim rs As Recordset
Dim SqLst As String
Dim X As String

Public Function Movefrist(ScrName As String)
On Error Resume Next
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
     Set rs = db.OpenRecordset("Category")
            If rs.RecordCount <> 0 Then
                rs.MoveFirst
                CategoryRec
                GoTo cF:
            Else
                MsgBox "Empty Table"
            End If
        
    Case 1
    Set rs = db.OpenRecordset("QProduct")
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
           ProductRec
           GoTo cF:
        Else
            MsgBox "Empty Table"
        End If
        
    Case 2
          Set rs = db.OpenRecordset("City")
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
            CityRec
            GoTo cF:
        Else
            MsgBox "Empty Table"
        End If
    Case 3
         Set rs = db.OpenRecordset("Type")
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
            TypeRec
            GoTo cF:
        Else
            MsgBox "Empty Table"
        End If
    Case 4
         Set rs = db.OpenRecordset("PosType")
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
            TypeposRec
            GoTo cF:
        Else
            MsgBox "Empty Table"
        End If
    
cF:
        frmMainData.cmdCMovelast.Enabled = True
        frmMainData.cmdCmovenext.Enabled = True
        frmMainData.cmdCmovefrist.Enabled = False
        frmMainData.cmdCmoveprevious.Enabled = False
    End Select
Case "frmpersonal"
'frmpersonal.txtCompanyName.Visible = True
'frmpersonal.cmbCompanyname.Visible = False
         SqLst = "Select * From Qpersonal"
             SqLst = SqLst & " ORDER BY personalId "
            Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
        rs.MoveFirst
            PersonalDis
        Else
        MsgBox "Empty Table"
        End If
        frmpersonal.cmdCMovelast.Enabled = True
        frmpersonal.cmdCmovenext.Enabled = True
        frmpersonal.cmdCmovefrist.Enabled = False
        frmpersonal.cmdCmoveprevious.Enabled = False

Case "frmCampany"

frmCampany.txtCompanyName.Visible = True
frmCampany.cmbCompanyname.Visible = False
frmCampany.txtAbbreviation.Visible = True
frmCampany.cmbabbr.Visible = False
            SqLst = "Select * From Qcompany"
             SqLst = SqLst & " ORDER BY CompanyId "
            Set rs = db.OpenRecordset(SqLst)
       ' Set rs = db.OpenRecordset("Qcompany")
        If rs.RecordCount <> 0 Then
        rs.MoveFirst
            CompanyRec
            
        Else
        MsgBox "Empty Table"
        End If
        frmCampany.cmdCMovelast.Enabled = True
        frmCampany.cmdCmovenext.Enabled = True
        frmCampany.cmdCmovefrist.Enabled = False
        frmCampany.cmdCmoveprevious.Enabled = False
        frmCampany.Command1.Enabled = True

End Select
Moving.fristmenu
End Function
Public Sub fristmenu()
frmMain.MP.Enabled = False
frmMain.mn.Enabled = True
frmMain.ML.Enabled = True
frmMain.MF.Enabled = False
End Sub
Public Sub lastmenu()
frmMain.MP.Enabled = True
frmMain.mn.Enabled = False
frmMain.ML.Enabled = False
frmMain.MF.Enabled = True
End Sub
Public Function MoveLast(ScrName As String)
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
     Set rs = db.OpenRecordset("Category")
            If rs.RecordCount <> 0 Then
                rs.MoveLast
                CategoryRec
                GoTo cL:
            Else
                MsgBox "Empty Table"
            End If
        
    Case 1
    Set rs = db.OpenRecordset("QProduct")
        If rs.RecordCount <> 0 Then
            rs.MoveLast
           ProductRec
           GoTo cL:
        Else
            MsgBox "Empty Table"
        End If
        
    Case 2
          Set rs = db.OpenRecordset("City")
        If rs.RecordCount <> 0 Then
            rs.MoveLast
            CityRec
            GoTo cL:
        Else
            MsgBox "Empty Table"
        End If
    Case 3
         Set rs = db.OpenRecordset("Type")
        If rs.RecordCount <> 0 Then
            rs.MoveLast
            TypeRec
            GoTo cL:
        Else
            MsgBox "Empty Table"
        End If
     Case 4
         Set rs = db.OpenRecordset("PosType")
        If rs.RecordCount <> 0 Then
            rs.MoveLast
            TypeposRec
            GoTo cL:
        Else
            MsgBox "Empty Table"
        End If

cL:
        frmMainData.cmdCMovelast.Enabled = False
        frmMainData.cmdCmovenext.Enabled = False
        frmMainData.cmdCmovefrist.Enabled = True
        frmMainData.cmdCmoveprevious.Enabled = True
    End Select

Case "frmCampany"
frmCampany.txtCompanyName.Visible = True
frmCampany.cmbCompanyname.Visible = False
frmCampany.txtAbbreviation.Visible = True
frmCampany.cmbabbr.Visible = False
         SqLst = "Select * From Qcompany"
             SqLst = SqLst & " ORDER BY CompanyId "
            Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
          rs.MoveLast
            CompanyRec
            
        Else
        MsgBox "Empty Table"
        End If
        frmCampany.cmdCMovelast.Enabled = False
        frmCampany.cmdCmovenext.Enabled = False
        frmCampany.cmdCmovefrist.Enabled = True
        frmCampany.cmdCmoveprevious.Enabled = True
        frmCampany.Command1.Enabled = True
Case "frmpersonal"
'frmpersonal.txtCompanyName.Visible = True
'frmpersonal.cmbCompanyname.Visible = False
         SqLst = "Select * From Qpersonal"
             SqLst = SqLst & " ORDER BY personalId "
            Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
          rs.MoveLast
            PersonalDis
            
        Else
        MsgBox "Empty Table"
        End If
        frmpersonal.cmdCMovelast.Enabled = False
        frmpersonal.cmdCmovenext.Enabled = False
        frmpersonal.cmdCmovefrist.Enabled = True
        frmpersonal.cmdCmoveprevious.Enabled = True

End Select
Moving.lastmenu
End Function

Public Function Previous(ScrName As String)
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
    frmMainData.cmdCMovelast.Enabled = True
    frmMainData.cmdCmovenext.Enabled = True
    frmMainData.cmdCmovefrist.Enabled = True
    frmMainData.cmdCmoveprevious.Enabled = True
    Save.savemenu
    Select Case X
    Case 0
                SqLst = "Select * From Category"
                Set rs = db.OpenRecordset(SqLst)
               rs.FindFirst "CategoryID = " & frmMainData.txtCategorycode & ""
                If rs.NoMatch = False Then
                   rs.MovePrevious
                If rs.BOF = True Then GoTo C6
                CategoryRec
 
                Else
C6:
                    frmMainData.cmdCMovelast.Enabled = True
                    frmMainData.cmdCmovenext.Enabled = True
                    frmMainData.cmdCmovefrist.Enabled = False
                    frmMainData.cmdCmoveprevious.Enabled = False
                    Moving.fristmenu
                End If
       Case 1
                SqLst = "Select * From QProduct"
                Set rs = db.OpenRecordset(SqLst)
               rs.FindFirst "ProductID = " & frmMainData.txtProductID & ""
                If rs.NoMatch = False Then
                   rs.MovePrevious
                If rs.BOF = True Then GoTo C7
                ProductRec
 
                Else
C7:
                    frmMainData.cmdCMovelast.Enabled = True
                    frmMainData.cmdCmovenext.Enabled = True
                    frmMainData.cmdCmovefrist.Enabled = False
                    frmMainData.cmdCmoveprevious.Enabled = False
                    Moving.fristmenu
                End If
            Case 2
             SqLst = "Select * From City"
                Set rs = db.OpenRecordset(SqLst)
               rs.FindFirst "citycode = " & frmMainData.txtcitycode & ""
                If rs.NoMatch = False Then
                   rs.MovePrevious
                If rs.BOF = True Then GoTo C8
                CityRec
 
                Else
C8:
                    frmMainData.cmdCMovelast.Enabled = True
                    frmMainData.cmdCmovenext.Enabled = True
                    frmMainData.cmdCmovefrist.Enabled = False
                    frmMainData.cmdCmoveprevious.Enabled = False
                    Moving.fristmenu
                End If
                
            Case 3
                SqLst = "Select * From Type"
                Set rs = db.OpenRecordset(SqLst)
               rs.FindFirst "TypeId = " & frmMainData.txtTypeID & ""
                If rs.NoMatch = False Then
                   rs.MovePrevious
                If rs.BOF = True Then GoTo C9
                TypeRec
 
                Else
C9:
                    frmMainData.cmdCMovelast.Enabled = True
                    frmMainData.cmdCmovenext.Enabled = True
                    frmMainData.cmdCmovefrist.Enabled = False
                    frmMainData.cmdCmoveprevious.Enabled = False
                    Moving.fristmenu
                    End If
          Case 4
                SqLst = "Select * From PosType"
                Set rs = db.OpenRecordset(SqLst)
                rs.FindFirst "posTypeId = " & frmMainData.txtTposID & ""
                If rs.NoMatch = False Then
                   rs.MovePrevious
                If rs.BOF = True Then GoTo C10
                TypeposRec
 
                Else
C10:
                    frmMainData.cmdCMovelast.Enabled = True
                    frmMainData.cmdCmovenext.Enabled = True
                    frmMainData.cmdCmovefrist.Enabled = False
                    frmMainData.cmdCmoveprevious.Enabled = False
                    Moving.fristmenu
                    End If
        End Select
Case "frmCampany"
    frmCampany.cmdCMovelast.Enabled = True
    frmCampany.cmdCmovenext.Enabled = True
    frmCampany.cmdCmovefrist.Enabled = True
    frmCampany.cmdCmoveprevious.Enabled = True
    Save.savemenu
    frmCampany.Command1.Enabled = True

   SqLst = "Select * From Qcompany"
    SqLst = SqLst & " ORDER BY CompanyId "
   Set rs = db.OpenRecordset(SqLst)
   rs.FindFirst "CompanyId = " & frmCampany.txtcompanyID & ""

   If rs.NoMatch = False Then
       rs.MovePrevious
   If rs.BOF = True Then GoTo C11
       CompanyRec
   Else

C11:
       frmCampany.cmdCMovelast.Enabled = True
       frmCampany.cmdCmovenext.Enabled = True
       frmCampany.cmdCmovefrist.Enabled = False
       frmCampany.cmdCmoveprevious.Enabled = False
       Moving.fristmenu
       frmCampany.Command1.Enabled = True
   End If
Case "frmpersonal"
    frmpersonal.cmdCMovelast.Enabled = True
    frmpersonal.cmdCmovenext.Enabled = True
    frmpersonal.cmdCmovefrist.Enabled = True
    frmpersonal.cmdCmoveprevious.Enabled = True
    Save.savemenu

   SqLst = "Select * From Qpersonal"
    SqLst = SqLst & " ORDER BY personalId "
   Set rs = db.OpenRecordset(SqLst)
   rs.FindFirst "personalId = " & frmpersonal.txtPersonalID & ""

   If rs.NoMatch = False Then
       rs.MovePrevious
   If rs.BOF = True Then GoTo C14
       PersonalDis
   Else

C14:
       frmpersonal.cmdCMovelast.Enabled = True
       frmpersonal.cmdCmovenext.Enabled = True
       frmpersonal.cmdCmovefrist.Enabled = False
       frmpersonal.cmdCmoveprevious.Enabled = False
       Moving.fristmenu
       
   End If

End Select
'rs.Close


End Function
Public Function NextM(ScrName As String)
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
    frmMainData.cmdCMovelast.Enabled = True
    frmMainData.cmdCmovenext.Enabled = True
    frmMainData.cmdCmovefrist.Enabled = True
    frmMainData.cmdCmoveprevious.Enabled = True
    Save.savemenu
    Select Case X
    Case 0
                SqLst = "Select * From Category"
                Set rs = db.OpenRecordset(SqLst)
               rs.FindFirst "CategoryID = " & frmMainData.txtCategorycode & ""
                If rs.NoMatch = False Then
                   rs.MoveNext
                If rs.EOF = True Then GoTo C1
                CategoryRec
 
                Else
C1:
                    frmMainData.cmdCMovelast.Enabled = False
                    frmMainData.cmdCmovenext.Enabled = False
                    frmMainData.cmdCmovefrist.Enabled = True
                    frmMainData.cmdCmoveprevious.Enabled = True
                    Moving.lastmenu
                End If
       Case 1
          SqLst = "Select * From QProduct"
                Set rs = db.OpenRecordset(SqLst)
               rs.FindFirst "ProductID = " & frmMainData.txtProductID & ""
                If rs.NoMatch = False Then
                   rs.MoveNext
                If rs.EOF = True Then GoTo C2
                ProductRec
 
                Else
C2:
                    frmMainData.cmdCMovelast.Enabled = False
                    frmMainData.cmdCmovenext.Enabled = False
                    frmMainData.cmdCmovefrist.Enabled = True
                    frmMainData.cmdCmoveprevious.Enabled = True
                    Moving.lastmenu
                End If
            Case 2
              SqLst = "Select * From City"
                Set rs = db.OpenRecordset(SqLst)
               rs.FindFirst "citycode = " & frmMainData.txtcitycode & ""
                If rs.NoMatch = False Then
                   rs.MoveNext
                If rs.EOF = True Then GoTo C3
                CityRec
 
                Else
C3:
                    frmMainData.cmdCMovelast.Enabled = False
                    frmMainData.cmdCmovenext.Enabled = False
                    frmMainData.cmdCmovefrist.Enabled = True
                    frmMainData.cmdCmoveprevious.Enabled = True
                    Moving.lastmenu
                End If
                
            Case 3
                SqLst = "Select * From Type"
                Set rs = db.OpenRecordset(SqLst)
               rs.FindFirst "TypeId = " & frmMainData.txtTypeID & ""
                If rs.NoMatch = False Then
                   rs.MoveNext
                If rs.EOF = True Then GoTo C4
                TypeRec
 
                Else
C4:
                    frmMainData.cmdCMovelast.Enabled = False
                    frmMainData.cmdCmovenext.Enabled = False
                    frmMainData.cmdCmovefrist.Enabled = True
                    frmMainData.cmdCmoveprevious.Enabled = True
                    Moving.lastmenu
                    End If
            Case 4
                SqLst = "Select * From PosType"
                Set rs = db.OpenRecordset(SqLst)
               rs.FindFirst "posTypeId = " & frmMainData.txtTposID & ""
                If rs.NoMatch = False Then
                   rs.MoveNext
                If rs.EOF = True Then GoTo C12
                TypeposRec
 
                Else
C12:
                    frmMainData.cmdCMovelast.Enabled = False
                    frmMainData.cmdCmovenext.Enabled = False
                    frmMainData.cmdCmovefrist.Enabled = True
                    frmMainData.cmdCmoveprevious.Enabled = True
                    Moving.lastmenu
                    End If
               
        End Select
Case "frmpersonal"
    frmpersonal.cmdCMovelast.Enabled = True
    frmpersonal.cmdCmovenext.Enabled = True
    frmpersonal.cmdCmovefrist.Enabled = True
    frmpersonal.cmdCmoveprevious.Enabled = True
    Save.savemenu

  SqLst = "Select * From Qpersonal"
    SqLst = SqLst & " ORDER BY personalId "
   Set rs = db.OpenRecordset(SqLst)
   rs.FindFirst "personalId = " & frmpersonal.txtPersonalID & ""

   If rs.NoMatch = False Then
       rs.MoveNext
   If rs.EOF = True Then GoTo C13
       PersonalDis
   Else

C13:
       frmpersonal.cmdCMovelast.Enabled = False
       frmpersonal.cmdCmovenext.Enabled = False
       frmpersonal.cmdCmovefrist.Enabled = True
       frmpersonal.cmdCmoveprevious.Enabled = True
       Moving.lastmenu
   End If
Case "frmCampany"
    frmCampany.cmdCMovelast.Enabled = True
    frmCampany.cmdCmovenext.Enabled = True
    frmCampany.cmdCmovefrist.Enabled = True
    frmCampany.cmdCmoveprevious.Enabled = True
    Save.savemenu
    frmCampany.Command1.Enabled = True
  SqLst = "Select * From Qcompany"
    SqLst = SqLst & " ORDER BY CompanyId "
   Set rs = db.OpenRecordset(SqLst)
   rs.FindFirst "CompanyId = " & frmCampany.txtcompanyID & ""

   If rs.NoMatch = False Then
       rs.MoveNext
   If rs.EOF = True Then GoTo C5
       CompanyRec
   Else

C5:
       frmCampany.cmdCMovelast.Enabled = False
       frmCampany.cmdCmovenext.Enabled = False
       frmCampany.cmdCmovefrist.Enabled = True
       frmCampany.cmdCmoveprevious.Enabled = True
       Moving.lastmenu
       frmCampany.Command1.Enabled = True
   End If
      
End Select
'rs.Close
End Function
Public Sub TypeRec()
    frmMainData.txtTypeID.Text = rs("TypeId")
    frmMainData.txtTypeName.Text = rs("Typename")
End Sub
Public Sub CategoryRec()
    frmMainData.txtCategorycode.Text = rs("CategoryID")
    frmMainData.txtCategoryname.Text = rs("Categoryname")
End Sub
Public Sub CityRec()
    frmMainData.txtcitycode.Text = rs("citycode")
    frmMainData.txtRegion.Text = rs("Region") & ""
    frmMainData.Cmbcountry.Text = rs("country")
     frmMainData.cmbCity.Text = rs("city")
End Sub
Public Sub CompanyRec()
    frmCampany.txtcompanyID.Text = rs("CompanyId")
     frmCampany.txtCompanyName.Text = rs("Companyname")
     frmCampany.txt_addno.Text = rs("addno") & ""
    frmCampany.txtaddress.Text = rs("Address") & ""
    frmCampany.cmbCity.Text = rs("city") & ""
    frmCampany.Cmbcountry.Text = rs("country") & ""
    frmCampany.CmbRegion.Text = rs("Region") & ""
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
Set rs = db.OpenRecordset(SqLst)
y = 0
Do While Not rs.EOF
frmCampany.Lstcategory.AddItem rs("Categoryname")
rs.MoveNext
frmCampany.Lstcategory.Selected(y) = True
y = y + 1
Loop
SqLst = "Select Productname From Qprocomp"
SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
frmCampany.LstProduct.Clear
Set rs = db.OpenRecordset(SqLst)
y = 0
Do While Not rs.EOF
frmCampany.LstProduct.AddItem rs("Productname")
frmCampany.LstProduct.Selected(y) = True
y = y + 1
rs.MoveNext
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
SqLst = "Select name From Qpersonal"
SqLst = SqLst & " WHERE Companyname = '" & frmCampany.txtCompanyName.Text & "'"
frmCampany.lst_person.Clear
Set rs = db.OpenRecordset(SqLst)

Do While Not rs.EOF
frmCampany.lst_person.AddItem rs("name")
rs.MoveNext
Loop
frmCampany.Command1.Enabled = True
End Sub
Public Sub ProductRec()
    frmMainData.txtProductID.Text = rs("ProductID")
    frmMainData.txtProductname.Text = rs("Productname")
    frmMainData.CmbCategoryname.Text = rs("Categoryname")
    End Sub
Public Sub TypeposRec()
    frmMainData.txtTposID.Text = rs("posTypeId")
    frmMainData.txtTposname.Text = rs("posTypename")
End Sub
Public Sub PersonalDis()
'frmpersonal.txt_name.SetFocus
frmpersonal.txtPersonalID.Text = rs("personalId")
frmpersonal.cmb_company.Text = rs("Companyname")
        frmpersonal.txt_tel.Text = rs("Tel1") & ""
        frmpersonal.txt_fax.Text = rs("fax1") & ""
        frmpersonal.cmb_titel.Text = rs("Title") & ""
        frmpersonal.txt_pos.Text = rs("Position") & ""
        frmpersonal.txt_notes.Text = rs("Notes") & ""
        frmpersonal.txt_name.Text = rs("name") & ""
        frmpersonal.txt_mail.Text = rs("Email") & ""
        frmpersonal.txt_mobile.Text = rs("Tel2") & ""
         If IsNull(rs("showPos")) Or rs("showPos") = "No" Then
        frmpersonal.chkpos.Value = 0
        Else
         frmpersonal.chkpos.Value = 1
         End If
        SqLst = "Select posTypename From Qper"
SqLst = SqLst & " WHERE personalId = " & frmpersonal.txtPersonalID.Text & ""
frmpersonal.lst_typ.Clear
Set rs = db.OpenRecordset(SqLst)
y = 0
Do While Not rs.EOF

frmpersonal.lst_typ.AddItem rs("posTypename")
rs.MoveNext
frmpersonal.lst_typ.Selected(y) = True
y = y + 1
Loop
SqLst = "Select name From Qpersonal"
SqLst = SqLst & " WHERE Companyname = '" & frmpersonal.cmb_company.Text & "'"
frmpersonal.lst_person.Clear
Set rs = db.OpenRecordset(SqLst)
y = 0
Do While Not rs.EOF
frmpersonal.lst_person.AddItem rs("name")
'frmpersonal.lst_person.Selected(y) = True
y = y + 1
rs.MoveNext
Loop
       
End Sub

