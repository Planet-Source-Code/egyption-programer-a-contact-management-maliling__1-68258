Attribute VB_Name = "Add"
Dim db As Database
Dim ws As Workspace
Dim rs As Recordset
Dim SqLst As String
Dim X, d As String
Public Function searshCom()
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
SqLst1 = "Select distinct Companyname From Qcompany"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'1==================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
    SqLst1 = "Select distinct Companyname From Qcompany"
SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'2======================================
 ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select distinct Companyname From Qcompany"
SqLst1 = SqLst1 & " WHERE country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'3===========================================

ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
    SqLst1 = "Select distinct Companyname From Qcompany"
SqLst1 = SqLst1 & " WHERE City = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
 '4===========================================
 ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
    SqLst1 = "Select distinct Companyname From Qcompany"
SqLst1 = SqLst1 & " WHERE City = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'5===================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text <> "" Then
    SqLst1 = "Select distinct Companyname From Qcatype"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'6=================
 ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text <> "" Then
     SqLst1 = "Select distinct Companyname From Qprodtype"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'7=================

ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
     SqLst1 = "Select distinct Companyname From Qcatype"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'8=================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
    SqLst1 = "Select distinct Companyname From Qprodtype"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Typename = '" & NewMsg.Cmbtype.Text & "'and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'9=================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
SqLst1 = "Select distinct Qcompany.Companyname From Qcatype,Qcompany"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Qcatype.Typename = '" & NewMsg.Cmbtype.Text & "'and Qcompany.country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and Qcompany.City = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function

'10=================
 ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
    SqLst1 = "Select distinct Qcompany.Companyname From Qprodtype,Qcompany"
SqLst1 = SqLst1 & " WHERE Qprodtype.Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Qprodtype.Typename = '" & NewMsg.Cmbtype.Text & "'and Qcompany.country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and Qcompany.City = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'11===============
  ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
 SqLst1 = "Select distinct Qcompany.Companyname From Qcatype,Qcompany"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'and Qcatype.Typename = '" & NewMsg.Cmbtype.Text & "'and Qcompany.country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and Qcompany.City = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and Qcompany.Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'12===============
 ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
    NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text <> "" _
    And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
     SqLst1 = "Select distinct Qcompany.Companyname From Qprodtype,Qcompany"
SqLst1 = SqLst1 & " WHERE Qprodtype.Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "'and Qprodtype.Typename = '" & NewMsg.Cmbtype.Text & "'and Qcompany.country = '" & NewMsg.Cmbcountry.Text
SqLst1 = SqLst1 & "'and Qcompany.city = '" & NewMsg.cmbCity.Text
SqLst1 = SqLst1 & "'and Qcompany.Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'13===============
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select distinct Companyname From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function

'14============================================
'ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text = "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select distinct Companyname From Qprocomp"
'SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
'SqLst1 = SqLst1 & "'"
'Set rs1 = db.OpenRecordset(SqLst1)
'GoTo endf
'Exit Function
'=============================================
'ElseIf NewMsg.CmbCategoryname.Text = "All" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select distinct Companyname From Qcatcomp"
'SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text
'SqLst1 = SqLst1 & "' ORDER BY CategoryID "
'Set rs1 = db.OpenRecordset(SqLst1)
'GoTo endf
'Exit Function
'============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text = "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select distinct Companyname From Qprocomp"
'Set rs1 = db.OpenRecordset(SqLst1)
'GoTo endf
'Exit Function
'===========================================
ElseIf NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select distinct Companyname From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'15=============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "All" And NewMsg.cmbCompanyname.Text <> "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "" Then
'SqLst1 = "Select distinct Companyname From Qprocomp"
'SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text & "'"
'Set rs1 = db.OpenRecordset(SqLst1)
'GoTo endf
'Exit Function
'============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbcountry.Text = "" And NewMsg.Cmbtype.Text = "All" Then
'SqLst1 = "Select distinct Companyname From Qtype"
'Set rs1 = db.OpenRecordset(SqLst1)
'GoTo endf
'Exit Function
'============================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbtype.Text <> "" Then
SqLst1 = "Select distinct Companyname From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'16=============================================
'ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text <> "" And _
'NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text = "" And NewMsg.CmbRegion.Text = "" _
'And NewMsg.Cmbtype.Text = "All" Then
'SqLst1 = "Select distinct Companyname From Qtype"
'SqLst1 = SqLst1 & " WHERE Companyname = '" & NewMsg.cmbCompanyname.Text & "'"
'Set rs1 = db.OpenRecordset(SqLst1)
'GoTo endf
'Exit Function
'====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select distinct Companyname From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'17====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
SqLst1 = "Select distinct Companyname From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' and city = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
''18====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
SqLst1 = "Select distinct Companyname From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & NewMsg.CmbCategoryname.Text
SqLst1 = SqLst1 & "' and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'19====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text = "" Then
SqLst1 = "Select distinct Companyname From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'20====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text = "" Then
SqLst1 = "Select distinct Companyname From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "' and city = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
''21====================================================
ElseIf NewMsg.CmbCategoryname.Text <> "" And NewMsg.CmbProduct.Text <> "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text = "" And NewMsg.CmbRegion.Text <> "" Then
SqLst1 = "Select distinct Companyname From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname = '" & NewMsg.CmbProduct.Text
SqLst1 = SqLst1 & "' and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
'22====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text = "" And NewMsg.Text1.Text = "" And NewMsg.CmbRegion.Text = "" _
And NewMsg.Cmbcountry.Text <> "" And NewMsg.Cmbtype.Text <> "" Then
SqLst1 = "Select distinct Companyname From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "' and country = '" & NewMsg.Cmbcountry.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
    '23====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" _
And NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text = "" Then
SqLst1 = "Select distinct Companyname From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "' and city = '" & NewMsg.cmbCity.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
''25====================================================
ElseIf NewMsg.CmbCategoryname.Text = "" And NewMsg.CmbProduct.Text = "" And NewMsg.cmbCompanyname.Text = "" And _
NewMsg.cmbCity.Text <> "" And NewMsg.Text1.Text = "" And NewMsg.Cmbcountry.Text <> "" And _
NewMsg.Cmbtype.Text <> "" And NewMsg.CmbRegion.Text <> "" Then
SqLst1 = "Select distinct Companyname From Qtype"
SqLst1 = SqLst1 & " WHERE Typename = '" & NewMsg.Cmbtype.Text
SqLst1 = SqLst1 & "' and Region = '" & NewMsg.CmbRegion.Text & "'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
  ''26====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(0).Value = True Then
SqLst1 = "Select distinct Companyname From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Categoryname like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
  ''27====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(1).Value = True Then
SqLst1 = "Select distinct Companyname From Qcatcomp"
SqLst1 = SqLst1 & " WHERE Companyname like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
  ''28====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(2).Value = True Then
SqLst1 = "Select distinct Companyname From Qprocomp"
SqLst1 = SqLst1 & " WHERE Productname like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
  ''29====================================================
ElseIf NewMsg.Text1.Text <> "" And NewMsg.Option1(3).Value = True Then
SqLst1 = "Select distinct Companyname From Qtype"
SqLst1 = SqLst1 & " WHERE Typename like" & _
"'*" & NewMsg.Text1.Text & "*'"
Set rs1 = db.OpenRecordset(SqLst1)
GoTo endf
Exit Function
endf:
NewMsg.cmbCompanyname.Clear
NewMsg.Lstcompany.Clear
NewMsg.Lstpersonalname.Clear
y = 0
'NewMsg.cmbCompanyname.AddItem "All"
    Do While Not rs1.EOF
       NewMsg.cmbCompanyname.AddItem rs1("Companyname")
       NewMsg.Lstcompany.AddItem rs1("Companyname")
       NewMsg.Lstcompany.Selected(y) = True
       rs1.MoveNext
 y = y + 1
       Loop
End If

NewMsg.Chkcompany = True
'30
End Function
Public Function Add_1(ScrName As String)
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
        SqLst = "SELECT max(CategoryID) as maxs FROM Category"
        Set rs = db.OpenRecordset(SqLst)
        If IsNull(rs("maxs")) Then
            frmMainData.txtCategorycode.Text = 1
        Else
            frmMainData.txtCategorycode.Text = rs("maxs") + 1
        End If
           
            frmMainData.txtCategoryname.Text = ""
           frmMainData.txtCategoryname.SetFocus
    Case 1
        SqLst = "SELECT max(ProductID) as maxs FROM Product"
        Set rs = db.OpenRecordset(SqLst)
        If IsNull(rs("maxs")) Then
            frmMainData.txtProductID.Text = 1
        Else
            frmMainData.txtProductID.Text = rs("maxs") + 1
        End If
           
            frmMainData.txtProductname.Text = ""
             frmMainData.txtProductname.SetFocus
            'frmMainData.CmbCategoryname.Text = ""
           
       
    Case 2
     SqLst = "SELECT max(citycode) as maxs FROM city "
        Set rs = db.OpenRecordset(SqLst)
        If IsNull(rs("maxs")) Then
            frmMainData.txtcitycode.Text = 1
        Else
            frmMainData.txtcitycode.Text = rs("maxs") + 1
        End If
        
        'frmMainData.Cmbcountry.Text = ""
        'frmMainData.cmbCity.Text = ""
        frmMainData.txtRegion.Text = ""
        frmMainData.txtRegion.SetFocus
             
    Case 3
     SqLst = "SELECT max(TypeId) as maxs FROM Type"
        Set rs = db.OpenRecordset(SqLst)
        If IsNull(rs("maxs")) Then
            frmMainData.txtTypeID.Text = 1
        Else
            frmMainData.txtTypeID.Text = rs("maxs") + 1
        End If
           
            frmMainData.txtTypeName.Text = ""
               frmMainData.txtTypeName.SetFocus
     Case 4
     SqLst = "SELECT max(posTypeId) as maxs FROM PosType"
        Set rs = db.OpenRecordset(SqLst)
        If IsNull(rs("maxs")) Then
            frmMainData.txtTposID.Text = 1
        Else
            frmMainData.txtTposID.Text = rs("maxs") + 1
        End If
           
            frmMainData.txtTposname.Text = ""
               frmMainData.txtTposname.SetFocus
        
    End Select
    frmMainData.cmdCmovenext.Enabled = False
    frmMainData.cmdCmoveprevious.Enabled = False
    
    frmMainData.cmdCMovelast.Enabled = True
    frmMainData.cmdCmovefrist.Enabled = True
Case "frmCampany"
    SqLst = "SELECT max(CompanyId) as maxs FROM Company"
    Set rs = db.OpenRecordset(SqLst)
    frmCampany.Command1.Enabled = False
    If IsNull(rs("maxs")) Then
       frmCampany.txtcompanyID.Text = 1
    Else
       frmCampany.txtcompanyID.Text = rs("maxs") + 1
    End If
    frmCampany.LstProduct.Clear
     frmCampany.LstType.Clear
    Set rs = db.OpenRecordset("Type")
     Do While Not rs.EOF
       frmCampany.LstType.AddItem rs("Typename")
        rs.MoveNext
    Loop
        frmCampany.Lstcategory.Clear
         Set rs = db.OpenRecordset("Category")
     Do While Not rs.EOF
        frmCampany.Lstcategory.AddItem (rs("Categoryname"))
        rs.MoveNext
    Loop
        Add.CompanyAdd
        frmCampany.cmdCmovenext.Enabled = False
        frmCampany.cmdCmoveprevious.Enabled = False
        frmCampany.cmdCMovelast.Enabled = True
        frmCampany.cmdCmovefrist.Enabled = True
Case "frmpersonal"
    SqLst = "SELECT max(personalId) as maxs FROM Personal"
    Set rs = db.OpenRecordset(SqLst)
    If IsNull(rs("maxs")) Then
       frmpersonal.txtPersonalID.Text = 1
    Else
       frmpersonal.txtPersonalID.Text = rs("maxs") + 1
    End If
    frmpersonal.lst_typ.Clear
     'frmCampany.LstType.Clear
    Set rs = db.OpenRecordset("PosType")
     Do While Not rs.EOF
       frmpersonal.lst_typ.AddItem rs("posTypename")
        rs.MoveNext
    Loop
        frmpersonal.cmb_company.Clear
        SqLst1 = "Select distinct Companyname From Company"
        Set rs = db.OpenRecordset(SqLst1)
        Do While Not rs.EOF
        frmpersonal.cmb_company.AddItem rs("Companyname")
        rs.MoveNext
        Loop
      PersonalAdd
        
        frmpersonal.cmdCmovenext.Enabled = False
        frmpersonal.cmdCmoveprevious.Enabled = False
        frmpersonal.cmdCMovelast.Enabled = True
        frmpersonal.cmdCmovefrist.Enabled = True


End Select
Add.addmenu
End Function

Public Function Mode()
frmMain.Toolbar1.Buttons(4).Enabled = True
frmMain.Toolbar1.Buttons(1).Enabled = True
frmMain.Toolbar1.Buttons(2).Enabled = True
frmMain.Toolbar1.Buttons(6).Enabled = True
frmMain.Toolbar1.Buttons(3).Enabled = True
frmMain.Toolbar1.Buttons(5).Enabled = True
frmMain.MF.Enabled = True
frmMain.ML.Enabled = True
frmMain.mn.Enabled = True
frmMain.MP.Enabled = True
frmMain.Menu003001.Enabled = True
frmMain.Menu003002.Enabled = True
frmMain.Menu003003.Enabled = True

End Function
Public Function ModeP()
frmMain.Toolbar1.Buttons(4).Enabled = False
frmMain.Toolbar1.Buttons(1).Enabled = False
frmMain.Toolbar1.Buttons(2).Enabled = False
frmMain.Toolbar1.Buttons(6).Enabled = False
frmMain.Toolbar1.Buttons(3).Enabled = False
frmMain.Toolbar1.Buttons(5).Enabled = False
frmMain.MF.Enabled = False
frmMain.ML.Enabled = False
frmMain.mn.Enabled = False
frmMain.MP.Enabled = False
frmMain.Menu003001.Enabled = False
frmMain.Menu003002.Enabled = False
frmMain.Menu003003.Enabled = False
End Function
Public Function ModeC()
frmMain.Toolbar1.Buttons(4).Enabled = False
frmMain.Toolbar1.Buttons(1).Enabled = False
frmMain.Toolbar1.Buttons(2).Enabled = True
'frmMain.Toolbar1.Buttons(6).Enabled = True
frmMain.Toolbar1.Buttons(3).Enabled = False
frmMain.Toolbar1.Buttons(5).Enabled = True
frmMain.Menu003001.Enabled = False
frmMain.Menu003002.Enabled = True
frmMain.Menu003003.Enabled = False
End Function
Public Sub PersonalAdd()
'frmpersonal.cmbCompanyname.Visible = False
'frmpersonal.txtCompanyName.Visible = True
'frmpersonal.cmb_company.Text = ""
frmpersonal.txt_name.SetFocus
        frmpersonal.txt_tel.Text = ""
        frmpersonal.txt_fax.Text = ""
        frmpersonal.cmb_titel.Text = ""
        frmpersonal.chkpos.Value = 0
        frmpersonal.txt_pos.Text = ""
        frmpersonal.txt_notes.Text = ""
        frmpersonal.txt_name.Text = ""
        frmpersonal.txt_mail.Text = ""
        frmpersonal.txt_mobile.Text = ""
        frmpersonal.lst_person.Clear
       ' If frmpersonal.cmb_company.Text <> "" Then
       'SqLst = "Select name From Qpersonal"
'SqLst = SqLst & " WHERE Companyname = '" & frmpersonal.cmb_company.Text & "'"
'frmpersonal.lst_person.Clear
'Set rs = db.OpenRecordset(SqLst)

'Do While Not rs.EOF
'frmpersonal.lst_person.AddItem rs("name")
'rs.MoveNext

'Loop
'End If
End Sub

Public Sub PersonserAdd()
frmpersonal.cmbname.Visible = False
frmpersonal.txt_name.Visible = True
frmpersonal.cmbname.Text = ""
        frmpersonal.txt_tel.Text = ""
        frmpersonal.txt_fax.Text = ""
        frmpersonal.cmb_titel.Text = ""
        frmpersonal.txt_pos.Text = ""
        frmpersonal.chkpos.Value = 0
        frmpersonal.txt_notes.Text = ""
        frmpersonal.txt_name.Text = ""
        frmpersonal.txt_mail.Text = ""
        frmpersonal.txt_mobile.Text = ""
        frmpersonal.cmbname.SetFocus
End Sub
Public Sub CompanyAdd()
frmCampany.cmbCompanyname.Visible = False
frmCampany.txtCompanyName.Visible = True
frmCampany.cmbabbr.Visible = False
frmCampany.txtAbbreviation.Visible = True
frmCampany.txtCompanyName.Text = ""
frmCampany.txtCompanyName.SetFocus
        frmCampany.txtAbbreviation.Text = ""
        frmCampany.txt_addno.Text = ""
        frmCampany.txtaddress.Text = ""
        frmCampany.cmbCity.Text = ""
        frmCampany.Cmbcountry.Text = ""
        frmCampany.CmbRegion.Text = ""
        frmCampany.txtPostelCode.Text = ""
        frmCampany.txtPbox.Text = ""
        frmCampany.txttel1.Text = ""
        frmCampany.txttel2.Text = ""
        frmCampany.txttel3.Text = ""
        frmCampany.txttel4.Text = ""
        frmCampany.txttel5.Text = ""
        frmCampany.txtFax1.Text = ""
        frmCampany.txtFax2.Text = ""
        frmCampany.txtEmail.Text = ""
        frmCampany.txtEmail2.Text = ""
        frmCampany.txtweb.Text = ""
        frmCampany.txtNots.Text = ""
        frmCampany.lst_person.Clear
End Sub

Public Sub addsearch(ScrName As String)
If Right(App.Path, 1) = "\" Then
            PathPro = App.Path & "Mailing.mdb"
        Else
            PathPro = App.Path & "\Mailing.mdb"
        End If
On Error Resume Next
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
Select Case ScrName
   Case "frmCampany"
    frmCampany.cmbCompanyname.Clear
    Set rs = db.OpenRecordset("Company")
     Do While Not rs.EOF
       frmCampany.cmbCompanyname.AddItem rs("Companyname")
        rs.MoveNext
    Loop
    frmCampany.cmbabbr.Clear
    Set rs = db.OpenRecordset("Company")
     Do While Not rs.EOF
     If IsNull(rs("Abbreviation")) Or rs("Abbreviation") = "Empty" Then
      
     Else
       frmCampany.cmbabbr.AddItem rs("Abbreviation")
       
     End If
      rs.MoveNext
    Loop
    
frmCampany.Lstcategory.Clear
frmCampany.LstProduct.Clear
frmCampany.LstType.Clear
Add.CompanyAdd
frmCampany.txtCompanyName.Visible = False
frmCampany.cmbCompanyname.Visible = True
frmCampany.txtAbbreviation.Visible = False
frmCampany.cmbabbr.Visible = True
frmCampany.txtcompanyID.Text = ""
frmCampany.txtcompanyID.Enabled = True
frmCampany.txtcompanyID.Locked = False
frmCampany.txtcompanyID.SetFocus
frmCampany.cmdCmovenext.Enabled = False
frmCampany.cmdCmoveprevious.Enabled = False
frmCampany.cmdCMovelast.Enabled = True
frmCampany.cmdCmovefrist.Enabled = True
Case "frmpersonal"
frmpersonal.cmbname.Clear
    Set rs = db.OpenRecordset("Personal")
     Do While Not rs.EOF
       frmpersonal.cmbname.AddItem rs("name")
        rs.MoveNext
    Loop
    frmpersonal.cmb_company.Text = ""
    
    
frmpersonal.lst_person.Clear
frmpersonal.lst_typ.Clear

Add.PersonserAdd
frmpersonal.txt_name.Visible = False
frmpersonal.cmbname.Visible = True
'frmpersonal.txtAbbreviation.Visible = False
'frmpersonal.cmbabbr.Visible = True

frmpersonal.txtPersonalID.Text = ""
frmpersonal.txtPersonalID.Enabled = True
frmpersonal.txtPersonalID.Locked = False
'frmpersonal.txtPersonalID.SetFocus
frmpersonal.cmdCmovenext.Enabled = False

frmpersonal.cmdCmoveprevious.Enabled = False
frmpersonal.cmdCMovelast.Enabled = True
frmpersonal.cmdCmovefrist.Enabled = True
End Select

End Sub

'Public Sub mode2()
'On Error Resume Next
'Set ws = CreateWorkspace("", "admin", "")
'Set db = ws.OpenDatabase(PathPro)
'If frmpersonal.lst_typ.Selected(frmpersonal.lst_typ.ListIndex) = True Then

'If frmpersonal.lst_typ.List(frmpersonal.lst_typ.ListIndex) <> "-1" Then

'SqLst1 = "Select Productname From QProduct"
'SqLst1 = SqLst1 & " WHERE Categoryname = '" & frmpersonal.Lstcategory.List(frmCampany.Lstcategory.ListIndex) & "'"
'Set rs1 = db.OpenRecordset(SqLst1, dbOpenDynaset)

'    Do While Not rs1.EOF
'        frmCampany.LstProduct.AddItem rs1("Productname")
'       rs1.MoveNext
'    Loop
'    End If
'  Else
'  SqLst1 = "Select Productname From QProduct"
'SqLst1 = SqLst1 & " WHERE Categoryname = '" & frmCampany.Lstcategory.List(frmCampany.Lstcategory.ListIndex) & "'"
'Set rs1 = db.OpenRecordset(SqLst1, dbOpenDynaset)

' A = 0
'e:
' Do While Not rs1.EOF
' A = -1
' Do While A < frmCampany.LstProduct.ListCount
'    B = rs1("Productname")
'  If B = frmCampany.LstProduct.List(frmCampany.LstProduct.ListIndex) Then
'        frmCampany.LstProduct.RemoveItem (frmCampany.LstProduct.ListIndex)
'       rs1.MoveNext
'       GoTo e
       
'    End If
 
'    frmCampany.LstProduct.ListIndex = A
'        A = A + 1
'
'  Loop
'   rs1.MoveNext

'  Loop
  
'  End If
'End Sub
Public Sub mode1()
On Error Resume Next
If Right(App.Path, 1) = "\" Then
            PathPro = App.Path & "Mailing.mdb"
        Else
            PathPro = App.Path & "\Mailing.mdb"
        End If
Set ws = CreateWorkspace("", "admin", "")
Set db = ws.OpenDatabase(PathPro, True, False, ";pwd=eit")
If frmCampany.Lstcategory.Selected(frmCampany.Lstcategory.ListIndex) = True Then

If frmCampany.Lstcategory.List(frmCampany.Lstcategory.ListIndex) <> "-1" Then

SqLst1 = "Select Productname From QProduct"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & frmCampany.Lstcategory.List(frmCampany.Lstcategory.ListIndex) & "'"
Set rs1 = db.OpenRecordset(SqLst1, dbOpenDynaset)
' frmCampany.LstProduct.Clear
 'Set rs = db.OpenRecordset("Category")
    Do While Not rs1.EOF
        frmCampany.LstProduct.AddItem rs1("Productname")
       rs1.MoveNext
    Loop
    End If
  Else
  SqLst1 = "Select Productname From QProduct"
SqLst1 = SqLst1 & " WHERE Categoryname = '" & frmCampany.Lstcategory.List(frmCampany.Lstcategory.ListIndex) & "'"
Set rs1 = db.OpenRecordset(SqLst1, dbOpenDynaset)
' frmCampany.LstProduct.Clear
 'Set rs = db.OpenRecordset("Category")
 A = 0
e:
 Do While Not rs1.EOF
 A = -1
 Do While A < frmCampany.LstProduct.ListCount
frmCampany.LstProduct.ListIndex = A
    B = rs1("Productname")
  If B = frmCampany.LstProduct.List(frmCampany.LstProduct.ListIndex) Then
        frmCampany.LstProduct.RemoveItem (frmCampany.LstProduct.ListIndex)
       rs1.MoveNext
       GoTo e

    End If


        A = A + 1
    
  Loop
   rs1.MoveNext

  Loop
 
  End If
End Sub
Public Function compers(s As String)

Add.Add_1 (frmpersonal.Name)
If s <> "" Then
frmpersonal.cmb_company.Text = s
End If
End Function

Public Sub addmenu()
frmMain.MP.Enabled = False
     frmMain.mn.Enabled = False
     frmMain.ML.Enabled = True
     frmMain.MF.Enabled = True
End Sub
