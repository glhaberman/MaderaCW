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
Dim mlngNxtID
Dim adCmd
Dim mlngNxtLstID
Dim mblnDelete
Dim mstrListName
Dim mstrMemberValue
Dim mintEdit
Dim mstrCalledFrom

mblnDelete = 1

Set adRs = Server.CreateObject("ADODB.Recordset")
Response.ExpiresAbsolute = Now - 5

mstrCalledFrom = Request.QueryString("CalledFrom")
If Len(mstrCalledFrom) = 0 Then mstrCalledFrom = "Self"

If mstrCalledFrom = "Self" Then
    mstrAction = ReqForm("FormAction")
    mlngNxtLstID = ReqForm("ID")
    mstrListName = ReqForm("ListName")
    mstrMemberValue = ReqForm("MemberValue")
    mintEdit = ReqForm("Edit")
Else
    mlngNxtLstID = Request.QueryString("ID")
    mstrListName = Request.QueryString("ListName")
    mstrMemberValue = Request.QueryString("MemberValue")
    mintEdit = Request.QueryString("Edit")
End If

Select Case ReqForm("FormAction")
	Case "Add"
		mlngNxtLstID = -1
		
    Case "AddSave"
        Set adCmd = GetAdoCmd("spListValueAdd")
            AddParmIn adCmd, "@ListName", adVarChar, 50, ReqForm("ListName")
            AddParmIn adCmd, "@ListValue", adVarChar, 255, ReqForm("MemberValue")
            AddParmIn adCmd, "@EditID", adInteger, 0, ReqForm("Edit")
            adCmd.Execute
        Set adCmd = Nothing
    Case "EditSave"
        Set adCmd = GetAdoCmd("spListValueUpd")
            AddParmIn adCmd, "@ListID", adInteger, 0, ReqForm("ID")
            AddParmIn adCmd, "@ListValue", adVarChar, 255, ReqForm("MemberValue")
            AddParmIn adCmd, "@EditID", adInteger, 0, ReqForm("Edit")
            adCmd.Execute
        Set adCmd = Nothing
    Case "Delete"
		mlngNxtLstID = -1
        Set adCmd = GetAdoCmd("spListValueDel")
			AddParmIn adCmd, "@ListID", adInteger, 0, ReqForm("ID")
            AddParmIn adCmd, "@ListName", adVarChar, 50, ReqForm("ListName")
            AddParmOut adCmd, "@USED", adInteger, 0
            adCmd.Execute
            mblnDelete = adCmd.Parameters("@USED").Value
            If mblnDelete = 0 Then
                mlngNxtLstID = ReqForm("ID")
            End If
        Set adCmd = Nothing
    Case Else
        'First time load of the page.
End Select

If Not IsNumeric(mlngNxtLstID) Then
    mlngNxtLstID = -1
ElseIf mlngNxtLstID = 0 Then
    mlngNxtLstID = -1
End If

If mlngNxtLstID <> -1 Then
    'Retrieve the values to display:
	Set adRs = Server.CreateObject("ADODB.Recordset") 
	Set adCmd = GetAdoCmd("spGetListValues")
		AddParmIn adCmd, "@ListName", adVarchar, 50, NULL
		AddParmIn adCmd, "@ValueID", adInteger, 0, mlngNxtLstID
		'Call ShowCmdParms(adCmd) '***DEBUG
	'Open a recordset from the query:
	Call adRs.Open(adCmd, , adOpenForwardOnly, adLockReadOnly)
	Set adCmd = Nothing
End If

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
	If <%=mblnDelete%> = 0 Then
		MsgBox "Unable to delete because the list value is being used.",vbInformation,"Drop Down Lists"
		Form.FormAction.Value = "Edit"
	End If
    If Form.FormAction.Value = "Delete" Or Form.FormAction.Value = "AddSave" Or Form.FormAction.Value = "EditSave" Then
		Top.Form.ListName.Value = Form.ListName.VAlue
		Top.Form.Action = "ListSelect.asp"
        top.mblnSetFocusToMain = False
		Top.Form.Submit
		Exit Sub
	End if

	Call FillScreen
	mstrOriginalText = Form.MemberValue.Value 
	txtValueText.disabled = True
    
	Form.FormAction.Value = ""
    PageFrame.disabled = False
    If Not IsNumeric(Form.Edit.value) Then Form.Edit.value = 0
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
            If Form.Edit.value <= 0 Then
	            cmdAdd.disabled = True
	            cmdDelete.disabled = True
            End If
            If Form.Edit.value = 0 Then
	            cmdEdit.disabled = True
            End If
        Case "Add","Edit"
	        cmdAdd.disabled = True
	        cmdEdit.disabled = True
	        cmdDelete.disabled = True
	        cmdSave.disabled = False
	        cmdCancel.disabled = False
    End Select
End Sub

Sub cmdSave_onclick()
    Dim intLp
    Dim blnFound
	 
    blnFound = False
   	For intLp = 0 To Window.parent.lstValues.options.length - 1
        If Window.parent.lstValues.options(intLp).Text = txtValueText.value Then
            blnFound = true
            Exit For
        End If
    Next
    If blnFound = True Then
        MsgBox "A duplicate value would be created.  Please modify the Value Text.", vbinformation, "Save"
        txtValueText.focus
        Exit Sub
    End If
    If Form.FormAction.Value = "Add" Then
        Form.FormAction.Value = "AddSave"
    ElseIf Form.FormAction.Value = "Edit" Then
        Form.FormAction.Value = "EditSave"
    End If
    
    Form.MemberValue.value = txtValueText.value
    Form.action = "ListAddEdit.asp"
    Form.submit
End Sub

Sub cmdCancel_onclick()
	txtValueID.value = Form.ID.Value
	txtValueText.value = mstrOriginalText
    txtValueText.disabled = True

    Call SetButtons("Cancel")
End Sub

Sub cmdAdd_onclick()
    txtValueID.value = ""
    txtValueText.value = ""
    txtValueText.disabled = False
    txtValueText.focus
    
    Call SetButtons("Add")

	txtValueText.focus
    Form.FormAction.Value = "Add"   
End Sub

Sub cmdEdit_onclick()
    Call SetButtons("Edit")
	txtValueText.disabled = False
	txtValueText.focus
	txtValueText.select
	Form.FormAction.Value = "Edit"
End Sub

Sub cmdDelete_onclick()
    Dim intResp
    
    intResp = MsgBox("Delete this " & window.parent.lstLists.options(window.parent.lstLists.selectedIndex).Text & "?", vbQuestion + vbYesNo, "Delete")
    If intResp = vbNo Then Exit Sub
    
    Form.ID.Value = txtValueID.value
    Form.FormAction.Value = "Delete"
    Form.action = "ListAddEdit.asp"
    Form.submit
End Sub

Sub FillScreen()
	txtValueID.Value = Form.ID.Value
	txtValueText.Value = Form.MemberValue.Value
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
</SCRIPT>
<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody style="Overflow:visible;background-color:<%=gstrBackColor%>">
    
    <DIV id=PageFrame class=DefPageFrame disabled=true style="Overflow:visible; LEFT:325; HEIGHT:340; WIDTH:400; BORDER:none">

        <SPAN id=lblValueID class=DefLabel style="LEFT:10; WIDTH:75; TOP:20">
            Value ID:
        </SPAN>
        <TEXTAREA id=txtValueID title="List value ID (read only)"
            style="LEFT:10; WIDTH:65; TOP:35; BACKGROUND-COLOR: buttonface" 
            onkeydown="Gen_onkeydown"
            tabIndex=2 disabled=true cols=26 NAME="txtValueID"></TEXTAREA> 

        <SPAN id=lblValueText class=DefLabel style="LEFT:10; WIDTH:75; TOP:70">
            Value Text:
        </SPAN>
        <TEXTAREA id=txtValueText title="Enter value text"
            style="LEFT:10; WIDTH:185; TOP:85" 
            onkeydown="Gen_onkeydown"
            tabIndex=4 cols=26 NAME="txtValueText"></TEXTAREA> 
	</DIV>
    <BUTTON class=DefBUTTON id="cmdAdd" title="Add a new drop down list item" 
	    style="LEFT:10; TOP:310;WIDTH:70;HEIGHT:20"accessKey=StabIndex=6>
	    <U>A</U>dd
    </BUTTON>
    <BUTTON class=DefBUTTON id=cmdEdit title="Edit drop down list item" 
		style="LEFT:85; TOP:310;WIDTH:70;HEIGHT:20"accessKey=StabIndex=6>
		<U>E</U>dit
    </BUTTON>
    <BUTTON class=DefBUTTON id="cmdDelete" title="Delete drop down list item" 
		style="LEFT:160;TOP:310;WIDTH:70;HEIGHT:20;"
        accessKey=StabIndex=6>
		<U>D</U>elete
    </BUTTON>
    <BUTTON class=DefBUTTON id=cmdSave title="Save changes to the drop down list item" 
		style="LEFT:235;TOP:310;WIDTH:70;HEIGHT:20"accessKey=StabIndex=6>
		<U>S</U>ave
    </BUTTON>
    <BUTTON class=DefBUTTON id=cmdCancel title="Cancel changes to the drop down list item" 
        style="LEFT:325;TOP:310;WIDTH:70;HEIGHT:20"accessKey=CtabIndex=7>
        <U>C</U>ancel
    </BUTTON>
</BODY>
<FORM NAME="Form" METHOD="Post" STYLE="VISIBILITY: hidden" ACTION="ListSelect.ASP" ID=Form>
    <%
    Call CommonFormFields()
    WriteFormField "FormAction", mstrAction
    IF mlngNxtLstID = -1 Then
		WriteFormField "ID", ""
		WriteFormField "ListName", mstrListName
		WriteFormField "MemberValue", ""
		WriteFormField "Edit", 0 
    Else
		WriteFormField "ID", adRs.Fields("lstID").Value
		WriteFormField "MemberValue", adRs.Fields("lstMemberValue").Value
		WriteFormField "ListName", adRs.Fields("lstName").Value
		WriteFormField "Edit", adRs.Fields("lstEdit").Value
    End If
    IF mlngNxtLstID <> -1 Then
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
