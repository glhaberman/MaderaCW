<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: ReviewTypeSelect.asp                                            '
'  Purpose: This screen allows the user to add/update/delete a reviwe type  '
'           definitions.                                                    '
'==========================================================================='
%>
<!-- METADATA TYPE="TypeLib" UUID="00000206-0000-0010-8000-00AA006D2EA4"" -->
<!--#include file="IncCnn.asp"-->
<%
Dim madoCmd
Dim adRs
Dim adRsElms
Dim strSQL
Dim mstrPageTitle
Dim strHTML
Dim madoRs
Dim mstrAction
Dim mlngNxtID
Dim adCmd
Dim mblnSaveError
Dim mlngNxtRvwID
Dim mblnDelete
Dim mlngProgramID
Dim mstrElementList

Set adRs = Server.CreateObject("ADODB.Recordset")
Response.ExpiresAbsolute = Now - 5
mstrAction = ReqForm("FormAction")
mlngNxtRvwID = Request.QueryString("ReviewTypeID")
mlngProgramID = Request.QueryString("Program")
If Len(mlngProgramID) = 0 Then mlngProgramID = 0
If mstrAction <> "" Then mlngProgramID = ReqForm("ProgramID")

mblnDelete = 0
Select Case mstrAction
	Case "Add"
		mlngNxtRvwID = -1
	
	Case "AddSave"
		Set adRs = Server.CreateObject("ADODB.Recordset")
		Set adCmd = GetAdoCmd("spReviewTypeAdd")
			addParmIn adCmd, "@ReviewTypeName", adVarChar, 255, ReqForm("ReviewTypeName")
			addParmIn adCmd, "@StartDate", adDBTimeStamp, 0, ReqIsDate("StartDate")
			addParmIn adCmd, "@EndDate", adDBTimeStamp, 0, ReqIsDate("EndDate")
			addParmIn adCmd, "@ProgramID", adInteger, 0, ReqForm("ProgramID")
			addParmIn adCmd, "@ElementList", adVarchar, 1000, ReqForm("ElementList")
			AddParmOut adCmd, "@NxtID", adInteger, 0
			adCmd.Execute
			mlngNxtRvwID = adCmd.Parameters("@NxtID").Value
	Case "EditSave"
		Set adRs = Server.CreateObject("ADODB.Recordset")
		Set adCmd = GetAdoCmd("spReviewTypeUpd")
			addParmIn adCmd, "@ReviewTypeID", adInteger, 0, ReqForm("ReviewTypeID")
			addParmIn adCmd, "@ReviewTypeName", adVarChar, 255, ReqForm("ReviewTypeName")
			addParmIn adCmd, "@StartDate", adDBTimeStamp, 0, ReqIsDate("StartDate")
			addParmIn adCmd, "@EndDate", adDBTimeStamp, 0, ReqIsDate("EndDate")
			addParmIn adCmd, "@ProgramID", adInteger, 0, ReqForm("ProgramID")
			addParmIn adCmd, "@ElementList", adVarchar, 1000, ReqForm("ElementLIst")
			adCmd.Execute
			mlngNxtRvwID = ReqForm("ReviewTypeID")
	Case "GetRecord"
        mlngNxtRvwID = ReqForm("ReviewTypeID")
        
    Case "Delete"
        'Delete an existing case review:
        Set adCmd = GetAdoCmd("spReviewTypeDel")
            AddParmIn adCmd, "@rteRecordID", adInteger, 0, ReqForm("ReviewTypeID")
			AddParmOut adCmd, "@UseCheck", adInteger, 0
            
        adCmd.Execute
        mblnDelete = adCmd.Parameters("@UseCheck").Value
        If mblnDelete = 0 Then 
			mlngNxtRvwID = 0
		Else
			mlngNxtRvwID = ReqForm("ReviewTypeID")
		End If
    Case Else
        'First time load of the page.
End Select

If Not IsNumeric(mlngNxtRvwID) Then
    mlngNxtRvwID = -1
ElseIf mlngNxtRvwID = 0 Then
    mlngNxtRvwID = -1
End If

If mlngNxtRvwID <> -1 Then
    'Retrieve the case to display:
	Set adRs = Server.CreateObject("ADODB.Recordset") 
	Set adCmd = GetAdoCmd("spGetReviewTypeDefs")
		AddParmIn adCmd, "@ReviewTypeID", adInteger, 0, mlngNxtRvwID
		AddParmin adCmd, "@ReviewTypeName", adVarchar, 100, NULL
	    AddParmIn adCmd, "@StartDate", adDBTimeStamp, 0, NULL
		AddParmIn adCmd, "@EndDate", adDBTimeStamp, 0, NULL
	'Open a recordset from the query:
	Call adRs.Open(adCmd, , adOpenForwardOnly, adLockReadOnly)
	Set adCmd = Nothing
End If

Set adRsElms = Server.CreateObject("ADODB.Recordset") 
Set adCmd = GetAdoCmd("spGetElements")
    AddParmIn adCmd, "@ProgramID", adInteger, 0, mlngProgramID
    AddParmIn adCmd, "@TypeID", adInteger, 0, 2
    Call adRsElms.Open(adCmd, , adOpenForwardOnly, adLockReadOnly)
Set adCmd = Nothing
mstrElementList = "|"
Do While Not adRsElms.EOF
    mstrElementList = mstrElementList & adRsElms.Fields("elmID").Value & "^" & _
        adRsElms.Fields("elmShortName").Value & "|" '& "^" & adRsElms.Fields("elmSortOrder").Value & "|"
    adRsElms.MoveNext
Loop
%>

<HTML>
<HEAD>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Dim mdctElements

Sub window_onload()
    Dim intPrg

    Set mdctElements = CreateObject("Scripting.Dictionary")
    Set mdctElements = LoadDictionaryObject("<% = mstrElementList %>")
    
    Call FillScreen
    
    Call DisableControls(True)
    
    If <%=mblnDelete%> = 1 Then
		MsgBox "Unable to delete because a review was entered using the review type.",vbInformation,"Review Types"
		Form.FormAction.Value = "Edit"
	Else
		If Form.FormAction.Value = "Delete" Then
			Form.FormAction.Value = "Add"
			top.Form.Action = "ReviewTypeSelect.asp"
            top.mblnSetFocusToMain = False
			top.Form.Submit
		End If
	End If
	
	If Form.FormAction.Value = "Delete" OR Form.FormAction.Value = "AddSave" OR Form.FormAction.Value = "EditSave" Then
		top.Form.ReviewTypeID.Value = <%=mlngNxtRvwID%>
		top.Form.Action = "ReviewTypeSelect.asp"
        top.mblnSetFocusToMain = False
		top.Form.Submit
	ElseIf Form.formAction.Value = "Add" Then
		Call DisableControls(False)
		cmdEdit.disabled = True
		cmdDelete.disabled = True
		cmdAdd.disabled = True
		txtTypeName.focus
    Else
        cmdAdd.disabled = False
	    If Form.ReviewTypeID.value = 0 Then
	        ' no review type exists
	    Else
		    Form.FormAction.Value = "Edit"
	        cmdEdit.disabled = False
	        cmdDelete.disabled = False
	    End If
	End if
	
	If Form.ProgramID.value > 0 Then
	    If Asc(Left(window.parent.cboProgram.options(window.parent.cboProgram.selectedindex).text,1)) = 160 Then
	        lblProgramSelected.innerHTML = Mid(window.parent.cboProgram.options(window.parent.cboProgram.selectedindex).text,3)
	    Else
	        lblProgramSelected.innerHTML = Trim(window.parent.cboProgram.options(window.parent.cboProgram.selectedindex).text)
	    End If
	    lblProgramSelected.style.fontWeight = "bold"
	Else
        cmdAdd.disabled = True
	End If
End Sub

Sub DisableControls(blnDisable)
	txtTypeName.disabled = blnDisable
	txtStartDate.disabled = blnDisable
	txtEndDate.disabled = blnDisable
	lstEligElements.disabled = blnDisable
	cmdSave.disabled = blnDisable 
	cmdCancel.disabled = blnDisable 
	cmdDelete.disabled = blnDisable
	cmdAddElement.disabled = blnDisable
	cmdRemoveElement.disabled = blnDisable
	lstEligElemSelected.disabled = blnDisable
	cmdAdd.disabled = blnDisable
	cmdEdit.disabled = blnDisable
End Sub

Sub cmdEdit_onclick()
	Call DisableControls(False)
	cmdAdd.disabled = True
	cmdEdit.disabled = True
	cmdDelete.disabled = True
	Form.FormAction.Value = "Edit"
	lblReadOnly.style.left = -1000
	If IsDate(Form.LastUsed.value) Then
	    ' Reviews exist for this review type, only allow edit of name
	    txtStartDate.disabled = True
	    'txtEndDate.disabled = True
	    lstEligElements.disabled = True
	    lstEligElemSelected.disabled = True
	    cmdAddElement.disabled = True
	    cmdRemoveElement.disabled = True
	    lblReadOnly.style.left = 290
	End If
	txtTypeName.select
	txtTypeName.focus
End Sub

Sub cmdAdd_onclick()
	Form.ReviewTypeName.Value = ""
	Form.StartDate.Value = ""
	Form.EndDate.Value = ""
	Form.ElementList.Value = ""
	'Form.ProgramID.Value = 0
	Form.FormAction.Value = "Add"
	Form.Action = "ReviewTypeAddEdit.asp"
    Form.Submit
End Sub

Sub cmdSave_onclick()
	Dim strMsg
	Dim blnValidationFailed
	
    window.parent.Form.ProgramID.Value = window.parent.cboProgram.Value
    
	blnValidationFailed = False
	strMsg = "The following items must be completed before the review type definition can be saved:" & space(10) & vbCrLf
	
	If txtTypeName.Value = "" Then
		strMsg = strMsg & vbCrLf & space(4) & "Review Type Name:  Please enter a name for the review type definition." & space(10)
		
		If Not blnValidationFailed Then
			txtTypeName.focus            
			blnValidationFailed = True
		End If
	End If
		
	If txtStartDate.Value = "" Then
		strMsg = strMsg & vbCrLf & space(4) & "Review Type Start Date:  Please enter a Start Date for the review type definition." & space(10)
		
		If Not blnValidationFailed Then
			txtStartDate.focus            
			blnValidationFailed = True
		End If
	End If
	
	If txtEndDate.Value <> "" Then
	    If CDate(txtStartDate.Value) > CDate(txtEndDate.Value)  Then
		    strMsg = strMsg & vbCrLf & space(4) & "Review Type End Date:  An End Date must take place after the Start Date." & space(10)
    		
		    If Not blnValidationFailed Then
			    txtEndDate.focus            
			    blnValidationFailed = True
		    End If
	    End If
    End If
    
    If lstEligElemSelected.options.length = 0 Then
		strMsg = strMsg & vbCrLf & space(4) & "Review Type Elements:  At least one Element must be selected." & space(10)
		
		If Not blnValidationFailed Then
			blnValidationFailed = True
		End If
	End If
	
	If blnValidationFailed Then
		MsgBox strMsg, vbInformation, "Save Review Type"
        Exit Sub
     End If
	
	Call FillForm()
    If Form.FormAction.Value = "Add" Then
        Form.FormAction.Value = "AddSave"
    ElseIf Form.FormAction.Value = "Edit" Then
        Form.FormAction.Value = "EditSave"
    End If
    Form.Action = "ReviewTypeAddEdit.asp"
    Form.Submit
End Sub

Sub cmdCancel_onclick()
	lblReadOnly.style.left = -1000
    window.parent.Form.ProgramID.Value = window.parent.cboProgram.Value
	top.Form.ReviewTypeID.Value = <%=mlngNxtRvwID%>
	top.Form.Action = "ReviewTypeSelect.asp"
    top.mblnSetFocusToMain = False
	top.Form.Submit
End Sub

Sub cmdDelete_onclick()
    Dim intResp
    window.parent.Form.ProgramID.Value = window.parent.cboProgram.Value
    intResp = MsgBox("Delete this review type?", vbYesNo + vbQuestion, "Delete Review Type")
    If intResp = vbNo Then
        Exit Sub
    End If
	Form.FormAction.Value = "Delete"
	Form.Action = "ReviewTypeAddEdit.asp"
	Form.Submit
End Sub

Sub FillForm()
	Dim intI
	Dim strEligElemList

	Form.ReviewTypeName.Value = txtTypeName.Value
	Form.StartDate.Value = txtStartDate.value
	Form.EndDate.Value = txtEndDate.Value

	strEligElemList = ""
	Form.ElementList.Value = ""
	For intI = 0 to lstEligElemSelected.options.length - 1
		strEligElemList = strEligElemList & "[" & lstEligElemSelected.options(intI).value & "]"
	Next
	
	Form.ElementList.Value = strEligElemList 
End Sub

Sub FillScreen()
	txtTypeName.value = Form.ReviewTypeName.Value
	txtStartDate.value = Form.StartDate.Value
	txtEndDate.value = Form.EndDate.Value
	Call FillElements()
End Sub

Sub FillElements()
	Dim oElm
	Dim strRecord
	Dim oOption

	For Each oElm In mdctElements
	    strRecord = mdctElements(oElm)

	    Set oOption = Document.createElement("OPTION")
	    oOption.Value = Parse(strRecord,"^",1)
	    oOption.Text = Parse(strRecord,"^",2)

	    If InStr(Form.ElementList.Value,"[" & Parse(strRecord,"^",1) & "]") > 0 Then
	        ' Selected
    	    lstEligElemSelected.options.Add oOption
    	Else
	        ' Available
    	    lstEligElements.options.Add oOption
    	End If
    	Set oOption = Nothing
	Next
End Sub

Sub Gen_onkeydown
    If window.event.keyCode = 13 Then
        call cmdFind_onclick()
    End If
End Sub

Sub Gen_focus(txtBox)
    txtBox.select
End Sub

Sub cmdAddElement_onclick()
	Dim oOption
	Dim intIndex
	Dim intPrg
    Dim intI
	If Not IsNumeric(lstEligElements.SelectedIndex) Then
        Exit Sub
    End If
    If lstEligElements.SelectedIndex = -1 Then
        Exit Sub
    End If
    
	intIndex = lstEligElements.SelectedIndex 
	Set oOption = Document.createElement("OPTION")
    oOption.Value = lstEligElements.options.Item(intIndex).Value
    oOption.Text = lstEligElements.options.Item(intIndex).Text
    lstEligElemSelected.options.Add oOption
    'Call SortcboList(lstEligElemSelected, "Text")
    
    lstEligElements.options.Remove(intIndex)
    lstEligElements.selectedindex = intIndex
End Sub

Sub cmdRemoveElement_onclick()
	Dim oOption
    Dim intIndex
    Dim intPrg
    Dim intI
	If Not IsNumeric(lstEligElemSelected.SelectedIndex) Then
        Exit Sub
    End If
    If lstEligElemSelected.SelectedIndex = -1 Then
        Exit Sub
    End If
    
	intIndex = lstEligElemSelected.SelectedIndex
    Set oOption = Document.createElement("OPTION")
    oOption.Value = lstEligElemSelected.options.Item(intIndex).Value
    oOption.Text = lstEligElemSelected.options.Item(intIndex).Text
    lstEligElements.options.Add oOption
    'Call SortcboList(lstEligElements, "Text")
    
    lstEligElemSelected.options.Remove(intIndex)
    lstEligElemSelected.SelectedIndex = intIndex
End Sub

'priority sorts either by "Value" or "Text", depending on which is passed
Sub SortcboList(cboList,strPriority)
	Dim strTemp1
	Dim strTemp2
	Dim strTemp3
	Dim arrList(2,500)
	
	Dim intI
	Dim intJ
	
	For intI = 0 to 500 Step 1
		arrList(0, intI) = 0
		arrList(1, intI) = ""
		arrList(2, intI) = 0
	Next
	For intI = 0 To cboList.options.Length - 1 Step 1
		arrList(0 , intI) = Parse(cboList.Options.Item(intI).Value, ":", 1)		
		arrList(1 , intI) = Parse(cboList.Options.Item(intI).Text, ":", 1)
		arrList(2 , intI) = Parse(cboList.Options.Item(intI).Value, ":", 2)
	Next
	
	If strPriority  = "Value" Then
		For intI = 0 To cboList.options.Length Step 1
			For intJ = 0 To cboList.options.Length Step 1
				If cint(arrList(0, intJ)) > cint(arrList(0, intJ + 1)) Then
					If Not cint(arrList(0, intJ)) = 0 Or cint(arrList(0, intJ + 1)) = 0 Then
						strTemp1 = cint(arrList(0, intJ + 1))
						strTemp2 = arrList(1, intJ + 1)
						strTemp3 = arrList(2, intJ + 1)
						
						arrList(0, intJ + 1) = cint(arrList(0, intJ))
						arrList(1, intJ + 1) = arrList(1, intJ)
						arrList(2, intJ + 1) = arrList(2, intJ)
						
						arrList(0, intJ) = cint(strTemp1)
						arrList(1, intJ) = strTemp2
						arrList(2, intJ) = strTemp3
					End IF
				End If
			Next
		Next
		
		intJ = 0
		For intI = 0 To 20 Step 1
			If Not arrList(0, intI) = 0 Then
				cboList.Options.Item(intJ).Value = arrList(0 , intI) & ":" & arrList(2 , intI)
				cboList.Options.Item(intJ).Text = arrList(1 , intI)
				intJ = intJ + 1
			End If
		Next
		
	ElseIf strPriority  = "Text" Then
		'Step through list beginning to end	
		For intI = 0 To cboList.options.Length Step 1
			For intJ = 0 To cboList.options.Length Step 1
				If StrComp(LCase(arrList(1, intJ)), LCase(arrList(1, intJ + 1)), 1) = 1 Then
					strTemp1 = arrList(0, intJ + 1)
					strTemp2 = arrList(1, intJ + 1)
					
					arrList(0, intJ + 1) = arrList(0, intJ)
					arrList(1, intJ + 1) = arrList(1, intJ)
					
					arrList(0, intJ) = strTemp1
					arrList(1, intJ) = strTemp2
				End If
			Next
		Next
		intJ = 0
		For intI = 0 To 20 Step 1
			If Not arrList(0, intI) = 0 Then
				cboList.Options.Item(intJ).Value = arrList(0 , intI)& ":" & arrList(2 , intI)
				cboList.Options.Item(intJ).Text = arrList(1 , intI)
				intJ = intJ + 1
			End IF
		Next
	End IF
	
End Sub

Sub txtStartDate_onkeypress()
    If txtStartDate.value = "(MM/DD/YYYY)" Then
        txtStartDate.value = ""
    End If
    Call TextBoxOnKeyPress(window.event.keyCode,"D")
End Sub
Sub txtStartDate_onblur()
    If Trim(txtStartDate.value) = "(MM/DD/YYYY)" Then
        txtStartDate.value = ""
        Exit Sub
    End If
    If Not ValidDate(txtStartDate.value) Then
        MsgBox "The Start Date must be a valid date - MM/DD/YYYY.", vbInformation, "Review Type Entry"
        txtStartDate.focus
        Exit Sub
    End If
    If Len(txtStartDate.value) > 0 Then
        ' If it gets here, a valid date has been entered.  Ensure that it is not a date in the past
        If CDate(txtStartDate.value) < CDate(FormatDateTime(Now(),2)) Then
            MsgBox "The Start Date cannot be before today's date.", vbInformation, "Review Type Entry"
            txtStartDate.focus
            Exit Sub
        End If
    End If
End Sub
Sub txtStartDate_onfocus()
    If Trim(txtStartDate.value) = "" Then
        txtStartDate.value = "(MM/DD/YYYY)"
    End If
    txtStartDate.select
End Sub

Sub txtEndDate_onkeypress()
    If txtEndDate.value = "(MM/DD/YYYY)" Then
        txtEndDate.value = ""
    End If
    Call TextBoxOnKeyPress(window.event.keyCode,"D")
End Sub
Sub txtEndDate_onblur()
    If Trim(txtEndDate.value) = "(MM/DD/YYYY)" Then
        txtEndDate.value = ""
        Exit Sub
    End If
    If Not ValidDate(txtEndDate.value) Then
        MsgBox "The End Date must be a valid date - MM/DD/YYYY.", vbInformation, "Review Type Entry"
        txtEndDate.focus
        Exit Sub
    End If
    If IsDate(Form.LastUsed.value) And IsDate(txtEndDate.value) Then
        ' End Date cannot be on or before the last date used in a review
        If CDate(txtEndDate.value) < CDate(Form.LastUsed.value) Then
            MsgBox "The End Date must be on or after last date used in a review."  & vbCrLf & "Last used " & Form.LastUsed.value & ".", vbInformation, "Review Type Entry"
            txtEndDate.focus
            Exit Sub
        End If
    End If
End Sub
Sub txtEndDate_onfocus()
    If Trim(txtEndDate.value) = "" Then
        txtEndDate.value = "(MM/DD/YYYY)"
    End If
    txtEndDate.select
End Sub

Sub txtTypeName_onkeypress()
    Call TextBoxOnKeyPress(window.event.keyCode,"X")
End Sub
</SCRIPT>
<!--#include file="IncCmnCliFunctions.asp"-->
<BODY id=PageBody>
    
        <DIV id=ReviewTypeSpec class=DefPageFrame style="LEFT:0; HEIGHT:100; WIDTH:483; TOP:0; BORDER:none">
			<SPAN id=lblTypeName class=DefLabel style="LEFT:20;WIDTH:100;TOP:10; TEXT-ALIGN: LEFT">
				Review Type Name
    			<INPUT ID=txtTypeName TYPE=text VALUE="" STYLE="LEFT:0; TOP:15" SIZE=25 NAME="txtTypeName">
			</SPAN>

			<SPAN id=lblProgram class=DefLabel style="LEFT:185; TOP:10; WIDTH:200">
				Function
			</SPAN>
			<SPAN id=lblProgramSelected class=DefLabel style="LEFT:185; TOP:25; WIDTH:200">
			</SPAN>
			
			<SPAN id=lblStartDate class=DefLabel style="LEFT:20; WIDTH:100; TOP:50">
				Start Date
    			<INPUT ID=txtStartDate TYPE=text VALUE="" STYLE="LEFT:0; TOP:15" SIZE=15 NAME="txtStartDate">
			</SPAN>
			
			<SPAN id=lblEndDate class=DefLabel style="LEFT:185;WIDTH:100;TOP:50">
				End Date
    			<INPUT ID=txtEndDate TYPE=text VALUE="" STYLE="LEFT:0; TOP:15" SIZE=15 NAME="txtEndDate">
			</SPAN>
			<SPAN id=lblReadOnly class=DefLabel style="LEFT:-1000; TOP:40; WIDTH:180;height:60;color:red;text-align:center">
				<B>Review Type has been used in a review, therefore only the Name and End Date can be changed.</B>
			</SPAN>
		</DIV>
		<DIV id=ProgElem class=DefPageFrame style="LEFT:0; HEIGHT:255; WIDTH:483; TOP:100; BORDER:none; BORDER-TOP:1 solid #a9a9a9; BORDER-BOTTOM:1 solid">	
			<SPAN id=lblEligElements class=DefLabel style="LEFT:20;WIDTH:200;TOP:30">
			    <SELECT id=lstEligElements style="LEFT:0;WIDTH:200;TOP:15;HEIGHT:175"size=8 NAME="lstEligElements">
			    </SELECT>
			</SPAN>

			<BUTTON id=cmdAddElement STYLE="LEFT:230; WIDTH:25; TOP:100; HEIGHT:20">
			    &gt
			</BUTTON>
			<BUTTON id=cmdRemoveElement STYLE="LEFT:230; WIDTH:25; TOP:125; HEIGHT:20" >
			    &lt
			</BUTTON>
			
			<SPAN id=lblEligElemSelected class=DefLabel style="LEFT:265;WIDTH:200;TOP:30">
    			Screens selected for this review type:
			    <SELECT id=lstEligElemSelected style="LEFT:0;WIDTH:200;TOP:15;HEIGHT:175" size=2 NAME="lstEligElemSelected"></SELECT>
			</SPAN>
		</DIV>
		<BUTTON class=DefBUTTON id=cmdAdd title="Add a new Review Type Definition"
			style="LEFT:10; TOP:365;WIDTH:70;HEIGHT:20"accessKey=StabIndex=6>
			<U>A</U>dd
        </BUTTON>
        <BUTTON class=DefBUTTON id=cmdEdit title="Add a new Review Type Definition"
			style="LEFT:85; TOP:365;WIDTH:70;HEIGHT:20"accessKey=StabIndex=6>
			<U>E</U>dit
        </BUTTON>
        <BUTTON class=DefBUTTON id=cmdDelete title="Delete Review Type Definition"
			style="LEFT:160;TOP:365;WIDTH:70;HEIGHT:20"accessKey=StabIndex=6>
			<U>D</U>elete
        </BUTTON>
        <BUTTON class=DefBUTTON id=cmdSave title="Save changes to the Review Type Definition"
			style="LEFT:235;TOP:365;WIDTH:70;HEIGHT:20"accessKey=StabIndex=6>
			<U>S</U>ave
        </BUTTON>
        <BUTTON class=DefBUTTON id=cmdCancel title="Cancel changes to the Review Type Definition"
            style="LEFT:325;TOP:365;WIDTH:70;HEIGHT:20"accessKey=CtabIndex=7>
            <U>C</U>ancel
        </BUTTON>
        
</BODY>

<FORM NAME="Form" METHOD="Post" STYLE="VISIBILITY: hidden" ACTION="ReviewTypeSelect.ASP" ID=Form2>
     <%
    Call CommonFormFields()
    WriteFormField "FormAction", mstrAction
    
    WriteFormField "ProgramID", mlngProgramID
    If mlngNxtRvwID = -1 Then
		WriteFormField "ReviewTypeID", 0
		WriteFormField "ReviewTypeName", ""
		WriteFormField "StartDate", ""
		WriteFormField "EndDate", ""
		WriteFormField "ElementList", ""
		WriteFormField "LastUsed", 0
    Else
		WriteFormField "ReviewTypeID", adRs.Fields("ReviewTypeID").Value
		WriteFormField "ReviewTypeName", TRIM(adRs.Fields("ReviewTypeName").Value)
		WriteFormField "StartDate", adRs.Fields("StartDate").Value
		WriteFormField "EndDate", adRs.Fields("EndDate").Value
		WriteFormField "ElementList", adRs.Fields("ElementList").Value
		WriteFormField "LastUsed", adRs.Fields("LastUsed").Value
		
    End If
    If mlngNxtRvwID <> -1 Then
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
