<%@ LANGUAGE="VBScript" EnableSessionState=False%><%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: ListAddEdit.asp                                                 '
'  Purpose: The primary data entry screen for maintaining the appliction    '
'           lists for the dropdown listboxes.                               '
'           This form is only available to admin users.                     '
'==========================================================================='
%>
<!-- METADATA TYPE="TypeLib" UUID="00000206-0000-0010-8000-00AA006D2EA4"" -->
<!--#include file="IncCnn.asp"-->
<%
Dim madoCmd
Dim adRs
Dim madoRs
Dim mstrAction
Dim adCmd
Dim mintAliasID
Dim mstrParentIDs
Dim mstrCalledFrom
Dim mlngTypeID
Dim mstrErrorMessage, mstrReload
Dim intReturnID

mblnDelete = 1
mstrErrorMessage = ""
mstrReload = ""

Set adRs = Server.CreateObject("ADODB.Recordset")
Response.ExpiresAbsolute = Now - 5

mstrCalledFrom = Request.QueryString("CalledFrom")
If Len(mstrCalledFrom) = 0 Then mstrCalledFrom = "Self"

If mstrCalledFrom = "Self" Then
    mintAliasID = ReqForm("ID")
    mlngTypeID = ReqForm("TypeID")
Else
    mintAliasID = Request.QueryString("ID")
    mlngTypeID = Request.QueryString("TypeID")
End If
If Len(mlngTypeID) = 0 Then mlngTypeID = 0

Select Case ReqForm("FormAction")
    Case "Add"
        mintAliasID = -1
        
    Case "AddSave"
        Set adCmd = GetAdoCmd("spAliasAdd")
            AddParmIn adCmd, "@TypeID", adInteger, 0, ReqForm("TypeID")
            AddParmIn adCmd, "@Name", adVarChar, 50, ReqForm("Name")
            AddParmOut adCmd, "@ReturnID", adInteger, 0            
            adCmd.Execute
            mintAliasID = adCmd.Parameters("@ReturnID").Value
        Set adCmd = Nothing
        Call ProcessParents()
        mstrReload = "Y"
    Case "EditSave"
        Set adCmd = GetAdoCmd("spAliasUpd")
            AddParmIn adCmd, "@AliasID", adInteger, 0, mintAliasID
            AddParmIn adCmd, "@NewName", adVarChar, 255, ReqForm("Name")
            AddParmOut adCmd, "@ReturnID", adInteger, 0            
            adCmd.Execute
            intReturnID = adCmd.Parameters("@ReturnID").Value
            If intReturnID = -2 Then
                mstrErrorMessage = "Error encountered when trying to save record."
            ElseIf intReturnID = -1 Then
                'ok
                Call ProcessParents()
            End If
        Set adCmd = Nothing
        mstrReload = "Y"
    Case "Delete"
        intReturnID = -1
        Set adCmd = GetAdoCmd("spAliasDel")
            AddParmIn adCmd, "@AliasID", adInteger, 0, mintAliasID
            AddParmOut adCmd, "@ReturnID", adInteger, 0
            adCmd.Execute
            intReturnID = adCmd.Parameters("@ReturnID").Value
            Select Case intReturnID
                Case 125
                    mstrErrorMessage = "Could not delete Office Manager as it was used in at least 1 review."
                Case 126
                    mstrErrorMessage = "Could not delete Region as it was used in at least 1 review."
                Case 250
                    mstrErrorMessage = "Could not delete FIPs as it was used in at least 1 review."
                Case -1
                    mstrReload = "Y"
                    mintAliasID = -1
                Case -2
                    mstrErrorMessage = "Error encountered when trying to delete record."
            End Select
        Set adCmd = Nothing
    Case Else
        'First time load of the page.
End Select

If Not IsNumeric(mintAliasID) Then
    mintAliasID = -1
ElseIf mintAliasID = 0 Then
    mintAliasID = -1
End If

mstrParentIDs = ""
If mintAliasID <> -1 Then
    'Retrieve the values to display:
    Set adRs = Server.CreateObject("ADODB.Recordset")
    Set madoCmd = GetAdoCmd("spGetAlaisIDs")
        AddParmIn madoCmd, "@ID", adInteger, 0, mintAliasID
        AddParmIn madoCmd, "@TypeID", adInteger, 0, NULL
        AddParmIn madoCmd, "@ParentID", adInteger, 0, NULL
        AddParmIn madoCmd, "@Name", adVarchar, 50, NULL
    adRs.Open madoCmd, , adOpenForwardOnly, adLockReadOnly
    Set madoCmd = Nothing 
    Do While Not adRs.EOF
        mstrParentIDs = mstrParentIDs & adRs.Fields("ParentID").Value & "*"
        adRs.MoveNext
    Loop
    adRs.MoveFirst
End If

Sub ProcessParents()
    Dim intI, strRecord
    If ReqForm("FormAction") = "EditSave" Then
        Set adCmd = GetAdoCmd("spAliasLinkDel")
            AddParmIn adCmd, "@AliasID", adInteger, 0, ReqForm("ID")
            adCmd.Execute
        Set adCmd = Nothing
    End If
    For intI = 1 To 100
        strRecord = Parse(ReqForm("ParentIDs"),"*",intI)
        If strRecord = "" Then Exit For
        Set adCmd = GetAdoCmd("spAliasLinkAdd")
            AddParmIn adCmd, "@AliasID", adInteger, 0, mintAliasID
            AddParmIn adCmd, "@ParentID", adInteger, 0, strRecord
            adCmd.Execute
        Set adCmd = Nothing
    Next
End Sub
%>
<HTML>
<HEAD>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Dim mstrOriginalText

Sub window_onload()
    If "<%=mstrErrorMessage%>" <> "" Then
        MsgBox "<%=mstrErrorMessage%>",vbInformation,"Upper Management"
        Form.FormAction.Value = "Edit"
    End If
    'If Form.FormAction.Value = "Delete" Or Form.FormAction.Value = "AddSave" Or Form.FormAction.Value = "EditSave" Then
    If "<%=mstrReload%>" = "Y" Then
        Top.Form.AliasTypeID.Value = Form.TypeID.Value
        Top.Form.AliasID.Value = Form.ID.Value
        window.parent.form.AliasID.value = Form.id.value
        Top.Form.Action = "AliasSelect.asp"
        top.mblnSetFocusToMain = False
        Top.Form.Submit
        Exit Sub
    End if

    Call FillScreen
    mstrOriginalText = Form.Name.Value 
    txtAliasName1.disabled = True
    txtAliasName2.disabled = True
    
    Form.FormAction.Value = ""
    PageFrame.disabled = False
    Call SetButtons("PageLoad")
End Sub

Sub SetButtons(strMode)
    Select Case strMode
        Case "PageLoad","Cancel"
            cmdAdd.disabled = False
            cmdEdit.disabled = False
            cmdDelete.disabled = False
            cmdSave.disabled = True
            cmdCancel.disabled = True
            Call DisableParents(True)
        Case "Add","Edit"
            cmdAdd.disabled = True
            cmdEdit.disabled = True
            cmdDelete.disabled = True
            cmdSave.disabled = False
            cmdCancel.disabled = False
            Call DisableParents(False)
    End Select
End Sub

Sub cmdSave_onclick()
    Dim intLp, intI
    Dim blnFound
     
    blnFound = False
    For intLp = 0 To Window.parent.cboAliasID.options.length - 1
        If Window.parent.cboAliasID.options(intLp).Text = txtAliasName1.value Then
            blnFound = true
            Exit For
        End If
    Next
    If blnFound = True Then
        MsgBox "A duplicate value would be created.  Please modify the Value Text.", vbinformation, "Save"
        txtAliasName1.focus
        Exit Sub
    End If
    If Form.FormAction.Value = "Add" Then
        Form.FormAction.Value = "AddSave"
    ElseIf Form.FormAction.Value = "Edit" Then
        Form.FormAction.Value = "EditSave"
    End If
    
    If Form.TypeID.value = 125 Then
        Form.Name.value = txtAliasName1.value & ", " & txtAliasName2.value
    Else
        Form.Name.value = txtAliasName1.value
    End If
    
    Form.ParentIDs.value = ""
    If Form.TypeID.value = 125 Or Form.TypeID.value = 250 Then
        For intI = 0 To hidLastParentID.value
            If document.all("lblParent" & intI).style.fontWeight = "bold" Then
                Form.ParentIDs.value = Form.ParentIDs.value & document.all("hidRowInfo" & intI).value & "*"
            End If
        Next
    End If
    Form.action = "AliasAddEdit.asp"
    Form.submit
End Sub

Sub cmdCancel_onclick()
    Dim intParentTypeID
    
    txtAliasID.value = Form.ID.Value
    txtAliasName1.Value = Trim(Parse(mstrOriginalText,",",1))
    txtAliasName2.Value = Trim(Parse(mstrOriginalText,",",2))
    txtAliasName1.disabled = True
    txtAliasName2.disabled = True

    If Form.TypeID.value = 125 Then
        intParentTypeID = 250
    ElseIf Form.TypeID.value = 250 Then
        intParentTypeID = 126
    ElseIf Form.TypeID.value = 126 Then
        intParentTypeID = 0
    End If
        
    Call BuildParentList(intParentTypeID, Form.ParentIDs.value)
    Call SetButtons("Cancel")
End Sub

Sub DisableParents(blnVal)
    Dim intI
    
    If divParents.innerHTML = "" Then Exit Sub
    
    For intI = 0 To hidLastParentID.value
        If blnVal Then
            document.all("lblParent" & intI).style.color = "gray"
            document.all("lblParent" & intI).style.cursor = "default"
        Else
            document.all("lblParent" & intI).style.color = "black"
            document.all("lblParent" & intI).style.cursor = "hand"
        End If
    Next
End Sub

Sub cmdAdd_onclick()
    Dim intParentTypeID
    
    txtAliasID.value = ""
    txtAliasName1.value = ""
    txtAliasName1.disabled = False
    txtAliasName2.value = ""
    txtAliasName2.disabled = False
    txtAliasName1.focus
    
    Call SetButtons("Add")

    txtAliasName1.focus
    If Form.TypeID.value = 125 Then
        intParentTypeID = 250
    ElseIf Form.TypeID.value = 250 Then
        intParentTypeID = 126
    ElseIf Form.TypeID.value = 126 Then
        intParentTypeID = 0
    End If
        
    Call BuildParentList(intParentTypeID, "")
    Form.FormAction.Value = "Add"   
End Sub

Sub cmdEdit_onclick()
    Call SetButtons("Edit")
    txtAliasName1.disabled = False
    txtAliasName2.disabled = False
    txtAliasName1.focus
    txtAliasName1.select
    Form.FormAction.Value = "Edit"
End Sub

Sub cmdDelete_onclick()
    Dim intResp
    
    intResp = MsgBox("Delete this " & window.parent.cboAliasType.options(window.parent.cboAliasType.selectedIndex).Text & "?", vbQuestion + vbYesNo, "Delete")
    If intResp = vbNo Then Exit Sub
    
    Form.ID.Value = txtAliasID.value
    Form.FormAction.Value = "Delete"
    Form.action = "AliasAddEdit.asp"
    Form.submit
End Sub

Sub FillScreen()
    Dim oAliasID
    Dim strHTML, strRecord
    Dim intParentTypeID, intI, intTop, strFont, intLeft, intJ

    If Form.TypeID.value = 125 Then
        lblAliasName1.innerText = "Last Name"
        lblAliasName2.style.left = 140
        txtAliasName2.style.left = 140
        txtAliasName1.Value = Trim(Parse(Form.Name.Value,",",1))
        txtAliasName2.Value = Trim(Parse(Form.Name.Value,",",2))
        intParentTypeID = 250
        lblParentHeading.innerText = "FIPs for Office Manager:"
        divParents.style.left = 10
    ElseIf Form.TypeID.value = 250 Then
        lblAliasName1.innerText = "FIPs"
        lblAliasName2.style.left = -1000
        txtAliasName2.style.left = -1000
        txtAliasName1.Value = Form.Name.Value
        intParentTypeID = 126
        lblParentHeading.innerText = "Region FIPs Located In:"
        divParents.style.left = 10
    ElseIf Form.TypeID.value = 126 Then
        lblAliasName1.innerText = "Region"
        lblAliasName2.style.left = -1000
        txtAliasName2.style.left = -1000
        txtAliasName1.Value = Form.Name.Value
        intParentTypeID = 0
        lblParentHeading.innerText = ""
        divParents.style.left = -1000
    End If
    txtAliasID.Value = Form.ID.Value
    
    Call BuildParentList(intParentTypeID, Form.ParentIDs.value)
End Sub

Sub BuildParentList(intParentTypeID, strParentList)
    Dim intI, intTop, strFont, intLeft, intJ
    
    intI = 0
    If intParentTypeID > 0 Then
        intTop = 0
        intLeft = 5
        intJ = 0
        For Each oAliasID In window.parent.mdctAliasIDs
            strRecord = window.parent.mdctAliasIDs(oAliasID)
            If CInt(Parse(strRecord,"^",1)) = CInt(intParentTypeID) Then
                'strHTML = strHTML & Parse(strRecord,"^",2) & vbCrLf
                strFont = "normal"
                If InStr("*" & strParentList,"*" & oAliasID & "*") > 0 Then
                    strFont = "bold"
                End If
                strHTML = strHTML & "<SPAN id=lblParent" & intI & " onclick=Parent_OnClick(" & intI & ") class=DefLabel style=""cursor:hand;LEFT:" & intLeft & ";WIDTH:100;TOP:" & intTop & ";font-weight:" & strFont & """>" & Parse(strRecord,"^",2) & "</SPAN>" & vbCrLf
                strHTML = strHTML & "<input type=hidden id=hidRowInfo" & intI & " value=" & oAliasID & ">" & vbCrLf
                intI = intI + 1
                intJ = intJ + 1
                If intJ = 10 Then
                    intLeft = intLeft + 100
                    intTop = -15
                    intJ = 0
                End If
                intTop = intTop + 15
            End If
        Next
    Else
        strHTML = ""
    End If
    intI = intI - 1
    
    strHTML = strHTML & "<input type=hidden id=hidLastParentID value=" & intI & ">" & vbCrLf
    divParents.innerHTML = strHTML
End Sub

Sub Gen_onkeydown
    If window.event.keyCode = 13 Then
        If cmdSave.disabled = false Then
            Call cmdSave_onclick
        End If
    ElseIf window.event.keyCode = 27 Then
        If cmdCancel.disabled = false Then
            Call cmdCancel_onclick
        End If
    End If
End Sub

Sub Parent_OnClick(intRowID)
    Dim intI
    
    If txtAliasName1.disabled = True Then Exit Sub
    
    If document.all("lblParent" & intRowID).style.fontWeight = "bold" Then
        document.all("lblParent" & intRowID).style.fontWeight = "normal"
    Else
        document.all("lblParent" & intRowID).style.fontWeight = "bold"
    End If
    If Form.TypeID.value = 250 Then
        Call ToggleRegion(intRowID)
    End If
End Sub

Sub ToggleRegion(intRowID)
    Dim intOtherRowID

    intOtherRowID = Abs(CInt(intRowID) - 1)

    If document.all("lblParent" & intRowID).style.fontWeight = "bold" Then
        document.all("lblParent" & intOtherRowID).style.fontWeight = "normal"
    Else
        document.all("lblParent" & intOtherRowID).style.fontWeight = "bold"
    End If
End Sub

</SCRIPT>
<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody style="Overflow:visible;background-color:<%=gstrBackColor%>">
    
    <DIV id=PageFrame class=DefPageFrame disabled=true style="Overflow:visible; LEFT:325; HEIGHT:340; WIDTH:400; BORDER:none">

        <SPAN id=lblAliasID class=DefLabel style="LEFT:10; WIDTH:75; TOP:20">
            ID:
        </SPAN>
        <INPUT type=text id=txtAliasID 
            style="LEFT:10; WIDTH:65; TOP:35; BACKGROUND-COLOR: buttonface" 
            onkeydown="Gen_onkeydown"
            tabIndex=-1 disabled=true cols=26 NAME="txtAliasID">

        <SPAN id=lblAliasName1 class=DefLabel style="LEFT:10; WIDTH:75; TOP:70">
            Name
        </SPAN>
        <INPUT type=text id=txtAliasName1
            style="LEFT:10; WIDTH:120; TOP:85" 
            onkeydown="Gen_onkeydown"
            tabIndex=10 NAME="txtAliasName1"> 
        <SPAN id=lblAliasName2 class=DefLabel style="LEFT:140; WIDTH:75; TOP:70">
            First Name
        </SPAN>
        <INPUT type=text id=txtAliasName2
            style="LEFT:140; WIDTH:100; TOP:85" 
            onkeydown="Gen_onkeydown"
            tabIndex=11 NAME="txtAliasName2">
        <SPAN id=lblParentHeading class=DefLabel style="LEFT:10; WIDTH:275; TOP:115">
        </SPAN>
        <DIV id=divParents class=DefPageFrame style="Overflow:auto;top:130; LEFT:10; HEIGHT:170; WIDTH:310; BORDER-STYLE:thin">
        </DIV>
    </DIV>
    <BUTTON class=DefBUTTON id="cmdAdd"
        style="LEFT:10; TOP:310;WIDTH:70;HEIGHT:20" accessKey=A tabIndex=12>
        <U>A</U>dd
    </BUTTON>
    <BUTTON class=DefBUTTON id=cmdEdit 
        style="LEFT:85; TOP:310;WIDTH:70;HEIGHT:20" accessKey=E tabIndex=13>
        <U>E</U>dit
    </BUTTON>
    <BUTTON class=DefBUTTON id="cmdDelete"
        style="LEFT:160;TOP:310;WIDTH:70;HEIGHT:20;"
        accessKey=S tabIndex=14>
        <U>D</U>elete
    </BUTTON>
    <BUTTON class=DefBUTTON id=cmdSave
        style="LEFT:235;TOP:310;WIDTH:70;HEIGHT:20" accessKey=S tabIndex=15>
        <U>S</U>ave
    </BUTTON>
    <BUTTON class=DefBUTTON id=cmdCancel
        style="LEFT:325;TOP:310;WIDTH:70;HEIGHT:20" accessKey=C tabIndex=16>
        <U>C</U>ancel
    </BUTTON>
</BODY>
<FORM NAME="Form" METHOD="Post" STYLE="VISIBILITY: hidden" ACTION="AliasSelect.ASP" ID=Form>
    <%
    Call CommonFormFields()
    WriteFormField "FormAction", mstrAction
    IF mintAliasID = -1 Then
        WriteFormField "ID", ""
        WriteFormField "Name", ""
        WriteFormField "TypeID", mlngTypeID
        WriteFormField "ParentIDs", "" 
    Else
        WriteFormField "ID", adRs.Fields("alsID").Value
        WriteFormField "Name", adRs.Fields("alsName").Value
        WriteFormField "TypeID", adRs.Fields("alsTypeID").Value
        WriteFormField "ParentIDs", mstrParentIDs
    End If
    IF mintAliasID <> -1 Then
        adRs.Close
    End if
    Set adRs = Nothing
    gadoCon.Close
    Set gadoCon = Nothing
    %>
</FORM>
</HTML>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
