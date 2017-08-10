<%@ LANGUAGE="VBScript" EnableSessionState=False%><%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: ListSelect.asp                                                  '
'  Purpose: This screen allows the user to select a list to modify.         '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<%
Dim adRs
Dim strSQL
Dim mstrPageTitle
Dim mstrList

mstrPageTitle = "Drop Down List Maintenance"

Set adRs = Server.CreateObject("ADODB.Recordset")
Set madoCmd = GetAdoCmd("spGetListValues")
    AddParmIn madoCmd, "@ListName", adVarchar, 50, NULL
    AddParmIn madoCmd, "@ValueID", adInteger, 0, NULL
adRs.Open madoCmd, , adOpenForwardOnly, adLockReadOnly
Set madoCmd = Nothing 
mstrList = "|"
Dim intEdit
Do While Not adRs.EOF
    Select Case adRs.Fields("lstName").Value
        Case "ReviewType" ', "ArrearageStatus","ElemProgStatus"
            ' Do not include
        Case Else
            If IsNull(adRs.Fields("lstEdit").Value) Then
                intEdit = 0
            Else
                intEdit = adRs.Fields("lstEdit").Value
            End If
            mstrList = mstrList & adRs.Fields("lstID").Value & "^" & _
                adRs.Fields("lstName").Value & "^" & _
                adRs.Fields("lstMemberValue").Value & "^" & _
                intEdit & "|"
    End Select
    adRs.MoveNext
Loop 
adRs.Close

%>
<HTML>
<HEAD>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
</HEAD>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Dim mdctLists
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnSetFocusToMain  <%'Flag used througout page to determine if focus should be set to main when window is closed.%>

Sub window_onload()
	Call SizeAndCenterWindow(767, 520, True)
    
    mblnSetFocusToMain = True
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>
    Set mdctLists = CreateObject("Scripting.Dictionary")
    Set mdctLists = LoadDictionaryObject("<% = mstrList %>")
    
    Call FillLists()

	Call lstLists_onChange
    PageFrame.disabled = False
End Sub
<%'If timer detects that Main has been closed, this sub will be called. %>
Sub MainClosed()
    mblnSetFocusToMain = False
    window.close
End Sub
Sub cmdClose_onclick()
    Call window.opener.ManageWindows(6,"Close")
End Sub
<%'If Main has not been closed, set focus back to it.%>
Sub window_onbeforeunload()
    If mblnSetFocusToMain = True Then
        window.opener.focus
    End If
End Sub

Sub lstLists_onchange()
	Dim intI
	Dim strOptionText
	Dim oOption
	Dim oList
	Dim strRecord

	lstValues.options.length = Null
    For Each oList In mdctLists
        strRecord = mdctLists(oList)
        strOptionText = Parse(strRecord,"^",3)
        If Parse(strRecord,"^",2) = lstLists.options(lstLists.selectedIndex).Text Then
			Set oOption = Document.createElement("OPTION")
			oOption.Value = oList
			oOption.Text = strOptionText
			lstValues.options.add oOption
			Set oOption = Nothing
        End If
    Next
	
	lstValues.selectedIndex = 0
	Call lstValues_onChange
End Sub

Sub FillLists()
	Dim intI
	Dim oOption
	Dim strLoaded
	Dim oList
	Dim strRecord
	Dim strOptionText
	Dim intDefault

    strLoaded = "|"
    intI = 0
    intDefault = 0
    For Each oList In mdctLists
        strRecord = mdctLists(oList)
        strOptionText = Parse(strRecord,"^",2)
        If InStr(strLoaded, "|" & strOptionText & "|") = 0 Then
            ' If option not in Loaded string, add it to drop down and string
            strLoaded = strLoaded & strOptionText & "|"

			Set oOption = Document.createElement("OPTION")
			oOption.Value = oList
			oOption.Text = strOptionText
			lstLists.options.add oOption
			Set oOption = Nothing
			If Trim(Form.ListName.value) = Trim(strOptionText) Then
			    intDefault = intI
			End If
			intI = intI + 1
        End If
    Next
    lstLists.selectedIndex = intDefault
End Sub

Sub lstValues_onChange()
    Dim lngID
    Dim strListName
    Dim strMemberValue
    Dim intEdit
    Dim strRecord

    strRecord = mdctLists(CLng(lstValues.value)) 'lstMasterList.options(lstValues.selectedIndex).Text
    lngID = lstValues.value
    strListName = Parse(strRecord,"^",2)
    strMemberValue = Parse(strRecord,"^",3)
    intEdit = Parse(strRecord,"^",4)
    Form.ListID.value = lstValues.value

	fraRev.frameElement.src = "ListAddEdit.asp?CalledFrom=Select&ID=" & lngID & _
	    "&ListName=" & strListName & _
	    "&MemberValue=" & strMemberValue & _
	    "&Edit=" & intEdit
End Sub


Sub Gen_onkeydown
    If window.event.keyCode = 13 Then
        Call cmdFind_onclick
    ElseIf window.event.keyCode = 27 Then
        Call cmdCancel_onclick
    End If
End Sub

Sub NavigateFix(strAction)
    If strAction = "Open" Then
        lblProgram.style.top = 47
        lstLists.style.top = 62
        lblValues.style.top = 85
        lstValues.style.top = 100
        lstValues.style.height = 167
    Else
        lblProgram.style.top = 20
        lstLists.style.top = 35
        lblValues.style.top = 70
        lstValues.style.top = 85
        lstValues.style.height = 176
    End If
End Sub
</SCRIPT>
<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody bottomMargin=5 leftMargin=5 topMargin=5 rightMargin=5 >
    
    <DIV id=Header class=DefTitleArea style="WIDTH:737; height:40">

        <SPAN id=lblAppTitleHiLight class=DefTitleTextHiLight style="WIDTH:737">
            <%=mstrPageTitle%>
        </SPAN>
        
        <SPAN id=lblAppTitle class=DefTitleText style="WIDTH:737">
            <%=mstrPageTitle%>
        </SPAN>
        <DIV id=divNavigateButton style="left:4;top:4;font-size:12;width:75;
            FONT-WEIGHT:bold;FONT-FAMILY:tahoma;COLOR:darkolivegreen;">
            Navigate
        </DIV>
    </DIV>
    <% Call WriteNavigateControls(-2,30,gstrBackColor) %>
            
    <DIV id=PageFrame class=DefPageFrame disabled=true style="WIDTH:737; HEIGHT:400; TOP:51">
		<SPAN id=lblinstructions class=DefLabel style="LEFT:50; TOP:10; WIDTH:600">
			Click a value in the list to modify or delete it. Click [Add] to add a new value to the list. 
			Click [Delete] to remove an item from the value list. Click [Save] to keep your changes. 
			Click [Cancel] to abandon changes. 
		</SPAN>
		<DIV id=divList class=DefPageFrame style="LEFT:50; HEIGHT:300; WIDTH:400; TOP:50; BORDER:NONE">
			
			<SPAN id=lblProgram class=DefLabel style="LEFT:10; TOP:20; WIDTH:200">
				Drop Down List Category
			</SPAN>		
			
			<SELECT id=lstLists  title="Select a list to modify" TYPE="select-one"
				style="LEFT:10; WIDTH:220; TOP:35" tabIndex=1 NAME="lstLists"></SELECT>
				
			<SELECT id=lstMasterList title="Select a list to modify" TYPE="select-one"
				style="LEFT:10; WIDTH:220; TOP:50; visibility:hidden" tabIndex=1 NAME="lstMasterList">
				<% Set madoCmd = GetAdoCmd("spGetListValues")
					AddParmIn madoCmd, "@ListName", adVarchar, 50, NULL
					AddParmIn madoCmd, "@ValueID", adInteger, 0, NULL
				adRs.Open madoCmd, , adOpenForwardOnly, adLockReadOnly
				Set madoCmd = Nothing 
				Do While Not adRs.EOF
					Response.Write "<OPTION VALUE=" & adRs.Fields("lstID").Value & ">" & adRs.Fields("LstName").Value & "^" & adRs.Fields("lstMemberValue").Value & "^" & adRs.Fields("lstEdit").Value
					adRs.MoveNext
				Loop 
				adRs.Close%>
			</SELECT>
			
			<SPAN id=lblValues class=DefLabel style="LEFT:10; WIDTH:185; TOP:70">
				Values in Selected List:
			</SPAN>
			<SELECT id=lstValues title="Values in selected list" TYPE="select-one"
				style="LEFT:10; WIDTH:220; TOP:85; HEIGHT:176" 
				onkeydown="Gen_onkeydown"
				tabIndex=1 size=13 NAME="lstValues">
			</SELECT>
			
		    <IFRAME ID=fraRev 
				SRC="ListAddEdit.asp?ValueID=0&ValueText=""" 
				STYLE="Left:10; WIDTH:650; HEIGHT:340;background-color:<%=gstrbackcolor%>" 
				FRAMEBORDER=0></IFRAME>

            <BUTTON id=cmdClose class=DefBUTTON title="Close the form" 
                style="LEFT:575; TOP:310; HEIGHT: 20; WIDTH:70"
                tabIndex=4>Close
            </BUTTON>
        </DIV>
    </DIV>
</BODY>

<FORM NAME="Form" METHOD="Post" STYLE="VISIBILITY: hidden" ACTION="Main.ASP" ID=Form>
    <%
    Call CommonFormFields()
    WriteFormField "ListID", ReqForm("ListID")
    WriteFormField "ListName", ReqForm("ListName")
    WriteFormField "FormAction", ""
    WriteFormField "ValueID", ReqForm("ValueID")
    WriteFormField "ValueText", ReqForm("ValueText")
    WriteFormField "EditID", ReqForm("EditID")
    gadoCon.Close
    Set gadoCon = Nothing
    %>
</FORM>
</HTML>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncNavigateControls.asp"-->
