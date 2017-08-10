<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>

<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->

<%
Dim adRsResults  
Dim strSQL
Dim adRs
Dim intResultCnt
Dim madoCmdStf
Dim mstrPrgRevTypes
Dim strHTML
Dim strTmp
Dim intLine
Dim adCmd
Dim strRowStart
Dim mstrPageTitle
Dim mstrVisible
Dim mintTblWidth
Dim mintTblCols

mstrPageTitle = "Review Type Selection"

%>
<HTML>
<HEAD>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncTableStyles.asp"-->
</HEAD>


<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Dim mblnLoadChild
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnSetFocusToMain  <%'Flag used througout page to determine if focus should be set to main when window is closed.%>

Sub window_onload()
    Dim intPrg
    Dim intReviewTypeID
    
    mblnLoadChild = False
    mblnSetFocusToMain = True
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>
	Call SizeAndCenterWindow(767, 520, True)
	If Form.ProgramID.Value = "" Then
		Form.ProgramID.Value = 0
	End If
	cboProgram.value = Form.ProgramID.Value
	If cboProgram.options.length = 2 Then
	    cboProgram.selectedIndex = 1
	    cboProgram.disabled = True
	End If
	Call cboProgram_onchange()
	
	mblnLoadChild = True
	intPrg = Parse(cboProgram.value,":",1)
	If Form.ReviewTypeID.Value <= 0 Then
	    If lstPrgRevTypes.options.length > 0 Then
		    lstPrgRevTypes.selectedIndex = 0
		    intReviewTypeID = lstPrgRevTypes.options(0).value
		Else
		    intReviewTypeID = 0
		End If
        fraRev.frameElement.src = "ReviewTypeAddEdit.asp?Program=" & intPrg & _
            "&ReviewTypeID=" & intReviewTypeID & "&ProgramName=&ReviewTypeName="
	Else
		' If Form.ReviewTypeID.Value > 0 and lstPrgRevTypes.options.length = 0, action was an add or a delete
		' If there are any items in the list, set selectedindex to first one, otherwise skip this section
        If lstPrgRevTypes.options.length > 0 Then
		    lstPrgRevTypes.Value = Form.ReviewTypeID.Value
		    If lstPrgRevTypes.SelectedIndex = -1 Then lstPrgRevTypes.SelectedIndex = 0
		    Call lstPrgRevTypes_onchange()
		Else
            fraRev.frameElement.src = "ReviewTypeAddEdit.asp?Program=" & intPrg & _
                "&ReviewTypeID=0&ProgramName=&ReviewTypeName="
		End If
	End IF
	
    PageFrame.disabled = False
    PageBody.style.cursor = "default"
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

Sub cboProgram_onchange() 
	Dim intI
    Dim intPrg
	Dim oOption
    
	intPrg = Parse(cboProgram.Value, ":", 1)
	If CInt(intPrg) < 0 Then
	    cboProgram.selectedIndex = cboProgram.selectedIndex + 1
	    intPrg = Parse(cboProgram.Value, ":", 1)
	End If
    lstPrgRevTypes.options.length = Null
	For intI = 0 To lstMasterPrgRev.options.length - 1
		If Parse(lstMasterPrgRev.options(intI).value, ":", 1) = intPrg Then
			Set oOption = Document.createElement("OPTION")
			oOption.Value = Parse(lstMasterPrgRev.options(intI).value, ":", 2)
			oOption.Text = lstMasterPrgRev.options(intI).Text
			lstPrgRevTypes.options.add oOption
			Set oOption = Nothing
		End If
	Next
	
	If lstPrgRevTypes.options.length > 0 Then
	    lstPrgRevTypes.selectedIndex = 0
	    Call lstPrgRevTypes_onchange()
	Else
	    If mblnLoadChild = True Then
            fraRev.frameElement.src = "ReviewTypeAddEdit.asp?Program=" & intPrg & "&ReviewTypeID=0&ProgramName=&ReviewTypeName="
        End If
	End If
End Sub

Sub lstPrgRevTypes_onchange()
    Dim strPrgID
    Dim strRevID
    Dim strPrgName
    Dim strRevName
    
    strPrgID = Parse(cboProgram.Value, ":", 1)
    strRevID = lstPrgRevTypes.value
    strPrgName = Trim(cboProgram.options.item(cboProgram.selectedIndex).text)
    strRevName = lstPrgRevTypes.options.item(lstPrgRevTypes.selectedIndex).text
    Form.ProgramID.Value = cboProgram.Value
    If mblnLoadChild = True Then
        fraRev.frameElement.src = "ReviewTypeAddEdit.asp?Program=" & strPrgID & "&ReviewTypeID=" & strRevID & "&ProgramName=" & strPrgName & "&ReviewTypeName=" & strRevName
    End If
End Sub

Sub Gen_onkeydown
    If window.event.keyCode = 13 Then
        Call cmdFind_onclick
    ElseIf window.event.keyCode = 27 Then
        Call cmdCancel_onclick
    End If
End Sub

Sub Gen_onfocus(txtBox)
    txtBox.select
End Sub

Sub NavigateFix(strAction)
    If strAction = "Open" Then
        divProgram.style.left = 20
        cboProgram.style.left = -1000
    Else
        divProgram.style.left = -1000
        cboProgram.style.left = 20
    End If
End Sub

</SCRIPT>

<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody bottomMargin=5 leftMargin=5 topMargin=5 rightMargin=5 style="cursor:wait">
    
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

		<DIV id=divProgramRevType class=DefPageFrame style="LEFT:0; HEIGHT:398; WIDTH:250; TOP:0; BORDER-COLOR: <%=gstrDefButtonColor%>">
			<SPAN id=lblinstructions class=DefLabel style="LEFT:10; TOP:10; WIDTH:230">
				Select a <B> Function </B> to display the Function's Review Types,
				Select <B> "All" </B> to display all the defined Review Types,  
				or Click on <B>Add</B> to create a new Review Type.
			</SPAN>
			<SPAN id=lblProgram class=DefLabel style="LEFT:10; TOP:90; WIDTH:200">
				Functions
			</SPAN>
		    <DIV id=divProgram style="LEFT:-1000; WIDTH:200; TOP:105;height:19;border-style:solid;border-width:1;
		        background-color:white;color:gray">
		        <DIV id=divProgramBtn style="LEFT:183; WIDTH:15; TOP:0"><IMG src="downclickbutton.bmp"></DIV>
		    </DIV>
			<SELECT id=cboProgram title="Select the Program" style="Width:200; Left:20; Top: 105" NAME="cboProgram">
                <% 
	            Dim adCmdPrg
                Dim adRsPrg
                Dim blnIndent
                Dim intLastGroup
                
                Set adRsPrg = Server.CreateObject("ADODB.Recordset")
                Set adCmdPrg = GetAdoCmd("spGetProgramList")
                    AddParmIn adCmdPrg, "@PrgID", adVarchar, 255, NULL ' mstrProgramsSelected
                    'Call ShowCmdParms(adCmdPrg) '***DEBUG
                    adRsPrg.Open adCmdPrg, , adOpenForwardOnly, adLockReadOnly
                Set adCmdPrg = Nothing

                strHTML = "<OPTION value=0 selected></OPTION>"
                intLastGroup = -1
                blnIndent = False
                Do While Not adRsPrg.EOF
                    If adRsPrg.Fields("prgID").Value < 50 And adRsPrg.Fields("prgID").Value <> 6 Then
                        strHTML = strHTML & "<OPTION VALUE=" & adRsPrg.Fields("prgID").Value & ">" & adRsPrg.Fields("prgShortTitle").Value
                    End If
                    adRsPrg.MoveNext 
                Loop 
                adRsPrg.Close
                Set adRsPrg = Nothing

                Response.Write strHTML
                strHTML = ""
                %>
			</SELECT>
			
			<SPAN id="Span1" class=DefLabel style="LEFT:10; TOP:130; WIDTH:200">
				Review Types
			</SPAN>
			<%
				Set adRs = Server.CreateObject("ADODB.Recordset")
				Set adCmd = GetAdoCmd("spGetALLReviewTypeDefs")

				Call adRs.Open(adCmd, , adOpenForwardOnly, adLockReadOnly)
				Set adCmd = Nothing
				mstrPrgRevTypes = ""
				Do While Not adRs.EOF
				    mstrPrgRevTypes = mstrPrgRevTypes & "<OPTION VALUE=" & adRs.Fields("ProgramID").Value & ":" & adRs.Fields("ReviewTypeID").Value & ">" & adRs.Fields("ReviewTypeName").Value
				    adRs.MoveNext 
				Loop 
			%>
			<SELECT id=lstPrgRevTypes style="LEFT:20;HEIGHT:auto;Width:200; TOP:145" size=10 NAME="lstPrgRevTypes">
			</SELECT>
			<SELECT id=lstMasterPrgRev style="LEFT:20;HEIGHT:auto;Width:200; TOP:200;visibility:hidden" size=10 NAME="lstMasterPrgRev">
				<%=mstrPrgRevTypes%>
			</SELECT>
		</DIV>
        <DIV id=divEdit class=DefPageFrame style="LEFT:250; HEIGHT:398; WIDTH:483; TOP:0; BORDER-COLOR: <%=gstrDefButtonColor%>">
            <IFRAME ID=fraRev src="Blank.html" STYLE="positon:absolute; LEFT:0; WIDTH:483; HEIGHT:396; TOP:0; BORDER:none" FRAMEBORDER=0></IFRAME>

            <BUTTON id=cmdClose class=DefBUTTON title="Close the form" 
                style="LEFT:400; TOP:365; HEIGHT: 20; WIDTH:70"
                tabIndex=4>Close
            </BUTTON>
        </DIV>
        
    </DIV>

    <%
    Response.Write "<FORM NAME=""Form"" METHOD=""Post"" STYLE=""VISIBILITY:hidden"" ACTION=""Main.ASP"" ID=Form>" & vbCrLf

    Call CommonFormFields()
	
    If ReqForm("ReviewTypeID") = "" Then
		WriteFormField "ProgramID", 0
        WriteFormField "ReviewTypeID", 0
        WriteFormField "ReviewTypeName", ""
        WriteFormField "StartDate", ""
        WriteFormField "EndDate", ""
    Else
		WriteFormField "ProgramID", ReqForm("ProgramID")
        WriteFormField "ReviewTypeID", ReqForm("ReviewTypeID")
		WriteFormField "ReviewTypeName", ReqForm("ReviewTypeName")
        WriteFormField "StartDate", ReqForm("StartDate")
        WriteFormField "EndDate", ReqForm("EndDate")
    End If
    WriteFormField "FormAction", ""
    If intLine > 0 Then
        WriteFormField "SelectedIndex", "1"
    Else
        WriteFormField "SelectedIndex", ""
    End if
    WriteFormField "ResultsCount", intLine - 1
    Response.Write Space(4) & "</FORM>"

    gadoCon.Close
    Set gadoCon = Nothing
    %>
</BODY>
</HTML>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncBuildList.asp"-->
<!--#include file="IncWriteTableCell.asp"-->
<!--#include file="IncNavigateControls.asp"-->