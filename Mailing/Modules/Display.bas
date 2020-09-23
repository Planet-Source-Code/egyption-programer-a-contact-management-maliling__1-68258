Attribute VB_Name = "Display"
Dim db As Database
Dim ws As Workspace
Dim rs As Recordset
Dim SqLst As String
Dim SqLst1 As String
Dim X, d As String

Public Function display(ScrName As String)
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
    
        Set rs = db.OpenRecordset("Category")
            If rs.RecordCount <> 0 Then
                rs.MoveFirst
                CategoryRec
    frmMainData.lstcat.Clear
    Do While Not rs.EOF
        frmMainData.lstcat.AddItem rs("Categoryname")
    rs.MoveNext
    Loop
            Else
                MsgBox "Empty Table"
           '     Clear.CategoryRec
            End If
        
    Case 1
            frmMainData.CmbCategoryname.Clear
            Set rs = db.OpenRecordset("Category")
        Do While Not rs.EOF
            frmMainData.CmbCategoryname.AddItem rs("Categoryname")
            rs.MoveNext
        Loop
            Set rs = db.OpenRecordset("QProduct")
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
            ProductRec
            SqLst = "Select distinct Productname From QProduct"
            SqLst = SqLst & " WHERE Categoryname ='" & frmMainData.CmbCategoryname.Text & "'"
            Set rs = db.OpenRecordset(SqLst)
        If rs.RecordCount <> 0 Then
            frmMainData.lstprod.Clear
        Do While Not rs.EOF
            frmMainData.lstprod.AddItem rs("Productname")
            rs.MoveNext
        Loop
        Else
            MsgBox "Empty Table"
            ' Clear.DepartmentRec
        End If
        End If
           Case 2
            frmMainData.Cmbcountry.Clear
            SqLst1 = "Select distinct country From city"
            Set rs = db.OpenRecordset(SqLst1)
            Do While Not rs.EOF
            frmMainData.Cmbcountry.AddItem rs("country")
            rs.MoveNext
            Loop
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
       Set rs = db.OpenRecordset("city")
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
        CityRec
        
        Else
            MsgBox "Empty Table"
           ' Clear.CityRec
        End If
     
    Case 3
         Set rs = db.OpenRecordset("Type")
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
           TypeRec
            frmMainData.lstype.Clear
        Do While Not rs.EOF
        frmMainData.lstype.AddItem rs("Typename")
        rs.MoveNext
    Loop
        Else
            MsgBox "Empty Table"
           ' Clear.DepartmentRec
        End If
    Case 4
         Set rs = db.OpenRecordset("PosType")
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
           TypeposRec
            frmMainData.lstpostype.Clear
        Do While Not rs.EOF
        frmMainData.lstpostype.AddItem rs("posTypename")
        rs.MoveNext
    Loop
        Else
            MsgBox "Empty Table"
           ' Clear.DepartmentRec
        End If
    End Select
    Case "frmpersonal"

 frmpersonal.cmb_company.Clear
SqLst1 = "Select distinct Companyname From Company"
Set rs = db.OpenRecordset(SqLst1)
Do While Not rs.EOF
frmpersonal.cmb_company.AddItem rs("Companyname")
rs.MoveNext
Loop
    
    'rs.Close
    Set rs = db.OpenRecordset("Qpersonal")
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
            PersonalDis
'SqLst = "Select posTypename From Qper"
'SqLst = SqLst & " WHERE personalId = " & frmpersonal.txtPersonalID.Text & ""
'frmpersonal.lst_typ.Clear
'Set rs = db.OpenRecordset(SqLst)
'y = 0
'Do While Not rs.EOF
'frmpersonal.lst_typ.AddItem rs("posTypename")
'frmpersonal.lst_typ.Selected(y) = True
'y = y + 1
'rs.MoveNext


'Loop
'SqLst = "Select name From Qpersonal"
'SqLst = SqLst & " WHERE Companyname = " & frmpersonal.cmb_company.Text & ""
'frmpersonal.lst_person.Clear
'Set rs = db.OpenRecordset(SqLst)
'y = 0
'Do While Not rs.EOF
'frmpersonal.lst_person.AddItem rs("name")
'frmpersonal.lst_person.Selected(y) = True
'y = y + 1
'rs.MoveNext
'Loop

    Else
        MsgBox "Empty Table"
       ' Clear.CompanyRec
    End If
'////////////////////////////////////////////////////////////////////
'********************************************************************

Case "frmCampany"
frmCampany.Cmbcountry.Clear
SqLst1 = "Select distinct country From city"
Set rs = db.OpenRecordset(SqLst1)
Do While Not rs.EOF
frmCampany.Cmbcountry.AddItem rs("country")
rs.MoveNext
Loop
Set rs = db.OpenRecordset("Qcompany")
If rs.RecordCount <> 0 Then
rs.MoveFirst
CompanyRec
SqLst = "Select Categoryname From Qcatcomp"
SqLst = SqLst & " WHERE CompanyId = " & frmCampany.txtcompanyID.Text & ""
frmCampany.Lstcategory.Clear
Set rs = db.OpenRecordset(SqLst)
y = 0
Do While Not rs.EOF
frmCampany.Lstcategory.AddItem rs("Categoryname")
frmCampany.Lstcategory.Selected(y) = True
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
Else
MsgBox "Empty Table"
' Clear.CompanyRec
End If
'Case "frmNotes"
'Set rs = db.OpenRecordset("Category")
'frmNotes.CmbCategoryname.Clear
'Do While Not rs.EOF
'frmNotes.CmbCategoryname.AddItem rs("Categoryname")
'
'rs.MoveNext
'Loop
End Select
End Function
Public Sub TypeposRec()
    frmMainData.txtTposID.Text = rs("posTypeId")
    frmMainData.txtTposname.Text = rs("posTypename")
End Sub
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
    frmMainData.txtRegion.Text = rs("Region")
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
            SqLst = "Select name From Qpersonal"
            SqLst = SqLst & " WHERE Companyname = '" & frmCampany.cmbCompanyname.Text & "'"
            frmCampany.lst_person.Clear
            Set rs = db.OpenRecordset(SqLst)
            
            Do While Not rs.EOF
            frmCampany.lst_person.AddItem rs("name")
            rs.MoveNext
            Loop
            End Sub
Public Sub ProductRec()
    frmMainData.txtProductID.Text = rs("ProductID")
    frmMainData.txtProductname.Text = rs("Productname")
    frmMainData.CmbCategoryname.Text = rs("Categoryname")
    End Sub

Public Sub TrueDes()
frmMain.Toolbar1.Buttons(1).Enabled = True
frmMain.Toolbar1.Buttons(2).Enabled = True
frmMain.Toolbar1.Buttons(3).Enabled = True
frmMain.Toolbar1.Buttons(4).Enabled = True
frmMain.Toolbar1.Buttons(5).Enabled = True
frmMain.Toolbar1.Buttons(6).Enabled = True
End Sub
Public Sub falseDes()
frmMain.Toolbar1.Buttons(1).Enabled = False
frmMain.Toolbar1.Buttons(2).Enabled = False
frmMain.Toolbar1.Buttons(3).Enabled = False
frmMain.Toolbar1.Buttons(4).Enabled = False
frmMain.Toolbar1.Buttons(5).Enabled = False
frmMain.Toolbar1.Buttons(6).Enabled = False
End Sub

Public Sub PersonalDis()
frmpersonal.txtPersonalID.Text = rs("personalId")
frmpersonal.cmb_company.Text = rs("Companyname")
frmpersonal.txt_name.SetFocus
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
y = 0
SqLst = "Select posTypename From Qper"
SqLst = SqLst & " WHERE personalId = " & frmpersonal.txtPersonalID.Text & ""
frmpersonal.lst_typ.Clear
Set rs = db.OpenRecordset(SqLst)
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


