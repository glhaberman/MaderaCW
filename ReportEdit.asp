<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: RepordEdit.asp                                                  '
'  Purpose: The primary admin data entry screen for maintaining the report  '
'			title and descriptions.											'
'           This form is only available to admin users.                     '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<%
Dim adRs
Dim strSQL
Dim mstrPageTitle
Dim adCmd
Dim mstrOptions
Dim mblnDeleteFailed
Dim madoRs
Set madoRs = Server.CreateObject("ADODB.Recordset")
mstrPageTitle = "Edit Reports"

Set adRs = Server.CreateObject("ADODB.Recordset")

Select Case ReqForm("FormAction")
    Case "Edit"
		Set adCmd = GetAdoCmd("spReportUpd")
			AddParmIn adCmd, "@rptID", adInteger, 0, ReqForm("rptID")
            AddParmIn adCmd, "@rptTitle", adVarChar, 255, ReqForm("rptTitle")
            AddParmIn adCmd, "@rptDescription", adVarChar, 5000, ReqForm("rptDescription")
            'Call ShowCmdParms(adCmd) '***DEBUG
		Set adRs = GetAdoRs(adCmd)
    Case "Disable"
        Set adCmd = GetAdoCmd("spReportAble")
            AddParmIn adCmd, "@rptID", adInteger, 0, ReqForm("rptID")
            AddParmIn adCmd, "@Function", adInteger, 0, 0
            'Call ShowCmdParms(adCmd) '***DEBUG
		Set adRs = GetAdoRs(adCmd)
    Case "Enable"
        Set adCmd = GetAdoCmd("spReportAble")
            AddParmIn adCmd, "@rptID", adInteger, 0, ReqForm("rptID")
            AddParmIn adCmd, "@Function", adInteger, 0, 1
            'Call ShowCmdParms(adCmd) '***DEBUG
		Set adRs = GetAdoRs(adCmd)
End Select
%>
<HTML>
<HEAD>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Dim mstrOriginalText
Dim mstrOriginalCode
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnSetFocusToMain  <%'Flag used througout page to determine if focus should be set to main when window is closed.%>

Sub window_onload()

    mblnSetFocusToMain = True
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>
    If Form.CalledFrom.Value = "FactorSelect.asp" Then
		Call SizeAndCenterWindow(767, 520, True)
    Else
		Call SizeAndCenterWindow(767, 520, False)
    End If
	Form.FormAction.Value = ""
    ButtonFrame.disabled = False
    Call FillReportList
    If <%=gblnUserQA%> Then
		cmdDisable.style.visibility = "visible"
		cmdEnable.style.visibility = "visible"
		lblQAInstructions.style.visibility = "visible"
	End If
	IF form.rptID.value <> "" Then
		lstReports.value = Form.rptID.value
	Else
		lstReports.selectedIndex = 0
	End If
    Call lstReports_onclick()
    lstReports.focus
    cmdSave.disabled = false
    cmdCancel.disabled = false
    cmdClose.disabled = false
End Sub

<%'If timer detects that Main has been closed, this sub will be called. %>
Sub MainClosed()
    mblnSetFocusToMain = False
    window.close
End Sub

<%'If Main has not been closed, set focus back to it.%>
Sub window_onbeforeunload()
    If mblnSetFocusToMain = True Then
        window.opener.focus
    End If
End Sub

Sub cmdSave_onclick()
    Dim intLp
    Dim intCntCode
    Dim intCntText
    
    Form.rptID.value = lstReports.value    
    Form.rptTitle.Value = Trim(txtTitle.value)
    Form.rptDescription.value = Trim(txtDescription.value)
    Form.FormAction.Value = "Edit"
    mblnSetFocusToMain = False
    Form.submit
End Sub

Sub cmdCancel_onclick()
	txtTitle.value = Parse(cboReportmaster.options(lstReports.selectedIndex).Text, "^", 1)    
	txtDescription.value = Parse(cboReportMaster.options(lstReports.selectedIndex).Text, "^", 2)
End Sub

Sub cmdClose_onclick()
    Call window.opener.ManageWindows(6,"Close")
End Sub

Sub cmdAble_onclick(intCmd)
    If intCmd = 0 Then
		Form.FormAction.Value = "Disable"
	Else
		Form.FormAction.Value = "Enable"
	End IF
    Form.rptID.value = lstReports.value
    mblnSetFocusToMain = False
    Form.submit    
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

Sub FillReportList
	Dim intI
	Dim oOption
	
	lstReports.options.length = Null
	
	For intI = 0 To cboReportMaster.options.length - 1
		Set oOption = Document.createElement("OPTION")
		oOption.Value = cboReportMaster.options(intI).value
		oOption.Text = Parse(cboReportMaster.options(intI).Text, "^", 1)
		If Instr(Parse(cboReportMaster.options(intI).Text, "^", 3), "[999]") > 0 Then
			oOption.style.Color = "#64865C"
		End If
		lstReports.options.add oOption
		Set oOption = Nothing
	Next
End Sub

Sub lstReports_onclick()
	txtTitle.value = Parse(cboReportmaster.options(lstReports.selectedIndex).Text, "^", 1)    
	txtDescription.value = Parse(cboReportMaster.options(lstReports.selectedIndex).Text, "^", 2)
	If Instr(Parse(cboReportMaster.options(lstReports.selectedIndex).Text, "^", 3), "[999]") > 0 Then
		cmdDisable.disabled = True
		cmdEnable.disabled = False
	Else
		cmdEnable.disabled = True
		cmdDisable.disabled = False
	End IF
	Call FillPrograms(Parse(cboReportmaster.options(lstReports.selectedIndex).Text, "^", 3))    
	txtOrgTitle.value = Parse(cboReportMaster.options(lstReports.selectedIndex).Text, "^", 4)
End Sub

Sub lstReports_onkeyup
	txtTitle.value = Parse(cboReportmaster.options(lstReports.selectedIndex).Text, "^", 1)    
	txtDescription.value = Parse(cboReportMaster.options(lstReports.selectedIndex).Text, "^", 2)
	If Instr(Parse(cboReportMaster.options(lstReports.selectedIndex).Text, "^", 3), "[999]") > 0 Then
		cmdDisable.disabled = True
		cmdEnable.disabled = False
	Else
		cmdEnable.disabled = True
		cmdDisable.disabled = False
	End IF
	Call FillPrograms(Parse(cboReportmaster.options(lstReports.selectedIndex).Text, "^", 3))
	txtOrgTitle.value = Parse(cboReportMaster.options(lstReports.selectedIndex).Text, "^", 4)
End Sub

Sub FillPrograms(strProgramList)
	Dim intI
	Dim strTmpList
	Dim oOption
	Dim intJ
	
	intI = 2
	strTmpList = Parse(strProgramList, "[", intI)
	lstPrograms.options.length = Null
	Do While strTmpList <> ""
		For intJ = 0 To cboProgramMaster.options.length - 1
			'Parse(strTmpList, "]", 1)
			IF cboProgramMaster.options(intJ).Value = Parse(strTmpList, "]", 1) Then
				Set oOption = Document.createElement("OPTION")
				oOption.Value = cboProgramMaster.options(intJ).value
				oOption.Text = Parse(cboProgramMaster.options(intJ).Text, "^", 1)
				lstPrograms.options.add oOption
				Set oOption = Nothing
			End If
		Next
		intI = intI + 1
		strTmpList = Parse(strProgramList, "[", intI)		
	Loop
End Sub

Sub NavigateFix(strAction)
    If strAction = "Open" Then
        lblReports.style.top = 105
        lstReports.style.top = 130
        lstReports.style.height = 265
    Else
        lblReports.style.top = 60
        lstReports.style.top = 75
        lstReports.style.height = 320
    End If
End Sub

</SCRIPT>
<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody bottomMargin=5 leftMargin=5 topMargin=5 rightMargin=5 style="BACKGROUND-COLOR: <%=gstrPageColor%>" >
    
    <DIV id=Header
		class=DefTitleArea 
        style="POSITION: absolute;
        BORDER-STYLE: solid;
        BORDER-WIDTH: 1px;
        BORDER-COLOR: <%=gstrBorderColor%>;
        BACKGROUND-COLOR: <%=gstrBackColor%>;
        COLOR: black;
        HEIGHT: 40;
        WIDTH: 747;">
       
        <SPAN id=lblAppTitle
            style="POSITION: absolute;
            COLOR: <%=gstrAccentColor%>;
            FONT-SIZE: <%=gstrTitleFontSize%>;
            FONT-FAMILY: <%=gstrTitleFont%>;
            HEIGHT: 20;
            TOP: 5;
            LEFT: 7;
            WIDTH: 733;
            TEXT-ALIGN: center;
            FONT-WEIGHT: bold"><%=mstrPageTitle%>
        </SPAN>
        <SPAN id=lblAppTitleHiLight
            style="POSITION: absolute;
            FONT-SIZE: <%=gstrTitleFontSize%>;
            FONT-FAMILY: <%=gstrTitleFont%>;
            COLOR: <%=gstrTitleColor%>;
            HEIGHT: 20;
            TOP: 4;
            LEFT: 6;
            WIDTH: 735;
            TEXT-ALIGN: center;
            FONT-WEIGHT: bold"><%=mstrPageTitle%>
        </SPAN>
        <DIV id=divNavigateButton style="left:4;top:4;font-size:12;width:75;
            FONT-WEIGHT:bold;FONT-FAMILY:tahoma;COLOR:darkolivegreen;">
            Navigate
        </DIV>
    </DIV>
    <% Call WriteNavigateControls(-2,30,gstrBackColor) %>
            
    <DIV id=ButtonFrame
		class=DefPageFrame
        disabled=true
        style="POSITION: absolute;
        BORDER-STYLE: solid;
        BORDER-WIDTH: 1px;
        BORDER-COLOR: <%=gstrBorderColor%>;
        BACKGROUND-COLOR: <%=gstrBackColor%>;
        COLOR: black;
        HEIGHT: 425;
        WIDTH: 747;
        FONT-SIZE: 14pt;
        TOP: 51">
		
		<%'Hidden for further developement on assigning programs to reports%>
		 <SPAN id=lblProgram class=DefLabel style="LEFT:350; WIDTH:46; TOP:280;FONT-WEIGHT: bold;">
            Programs:
        </SPAN>

        <SELECT id=cboProgramMaster title="Program" style="LEFT:40; WIDTH:225; TOP:25;visibility:hidden" tabIndex=7 NAME="cboProgramMaster">
            <%
				Set adRs = Server.CreateObject("ADODB.Recordset")
				Set adCmd = GetAdoCmd("spGetProgramList")
					AddParmIn adCmd, "@PrgID", adVarchar, 255, NULL
					'Call ShowCmdParms(adCmd) '***DEBUG
					adRs.Open adCmd, , adOpenForwardOnly, adLockReadOnly
				Set adCmd = Nothing
            mstrOptions = ""
			Do While Not adRs.EOF
				mstrOptions = mstrOptions & "<OPTION VALUE=" & adRs.Fields("prgID").Value & ">" & adRs.Fields("prgShortTitle").Value
				adRs.MoveNext 
			Loop 
			adRs.Close
			Set adRs = Nothing
			Response.Write mstrOptions
			%>
        </SELECT>
        
        <SELECT id=lstPrograms title="Program" style="LEFT:350;WIDTH: 240; TOP:295; Height:98; BACKGROUND-COLOR: <%=gstrPageColor%>" size=2 NAME="lstPrograms"></SELECT>
		
        <SELECT id=cboReportMaster style="VISIBILITY:hidden; WIDTH:225; LEFT:500; TOP:25" tabIndex=0 NAME="cboReportMaster">
            <%
            Set adCmd = Server.CreateObject("ADODB.Command")
                With adCmd
                   .ActiveConnection = gadoCon
                   .CommandType = adCmdStoredProc
                   .CommandText = "spGetReportList"
                   .Parameters.Append .CreateParameter("@Programs", adVarChar,adParamInput, 255, NULL)
                End With
                'Open a recordset from the query:
                Set adRs = Server.CreateObject("ADODB.Recordset") 
                Call adRs.Open(adCmd, , adOpenForwardOnly, adLockReadOnly)
                mstrOptions = ""
				Do While Not adRs.EOF And Not adRs.BOF
					mstrOptions = mstrOptions & "<OPTION VALUE=" & adRs.Fields("rptRecordID").Value & ">" & adRs.Fields("rptReportTitle").Value & "^" & adRs.Fields("rptDescription").Value & "^" & adRs.Fields("rptProgramList").Value & "^" & adRs.Fields("rptOriginalTitle").Value
                    adRs.MoveNext
                Loop
                adRs.Close
                Set adCmd = Nothing
                Response.Write mstrOptions%>
        </SELECT>

        <SPAN class=DefLabel 
            id=lblReports
            style="POSITION: absolute;
                LEFT: 40; 
                WIDTH: 185; 
                FONT-WEIGHT: bold; 
                TOP: 60">Reports:
        </SPAN>
        <SELECT id=lstReports
            title="Report List"
            style="Z-INDEX: 2; 
                LEFT: 40; 
                WIDTH: 240; 
                POSITION: absolute; 
                TOP: 75; 
                HEIGHT:320" 
                tabIndex=1 
                size=2 
                TYPE="select-one" NAME="lstReports">
               
        </SELECT>

        <SPAN class=DefLabel 
            id=lblInstructions
            style="POSITION: absolute;
                FONT-WEIGHT: bold;
                LEFT: 40; 
                WIDTH: 280; 
                TOP: 10">Instructions:
        </SPAN>
        <SPAN class=DefLabel 
            id=lblInstructionsText
            style="POSITION: absolute;
                LEFT: 40; 
                WIDTH: 600;
                HEIGHT: 75; 
                OVERFLOW: hidden;
                TOP: 25">Select a report in the list to modify. Click [Save] to keep your changes. Click [Cancel] to abandon changes.
        </SPAN>
		<SPAN class=DefLabel 
            id="lblQAInstructions"
            style="POSITION: absolute;
                LEFT: 40; 
                WIDTH: 600;
                HEIGHT: 75; 
                OVERFLOW: hidden;
                TOP: 40;
				VISIBILITY: hidden">Click [Disable] to Disable a report. Click [Endable] to to enable a report. All Disabled reports appear in green.
        </SPAN>
        <SPAN class=DefLabel 
            id=lblTitle
            style="POSITION: absolute;
                LEFT: 350; 
                WIDTH: 150;
                FONT-WEIGHT: bold; 
                TOP: 60">Report Title:
        </SPAN>
        <TEXTAREA id=txtTitle
            title="Report Title"
            style="POSITION: absolute;
                LEFT: 350; 
                WIDTH: 240; 
                TOP: 75" 
                tabIndex=2
                cols=26 NAME="txtTitle"></TEXTAREA> 
		
		<SPAN class=DefLabel 
            id=lblOrgTitle
            style="POSITION: absolute;
                LEFT: 350; 
                WIDTH: 150;
                FONT-WEIGHT: bold; 
                TOP: 100">Original Report Title:
        </SPAN>
        <TEXTAREA id=txtOrgTitle
            title="Original Report Title"
            Disabled=True
            style="POSITION: absolute;
                LEFT: 350; 
                WIDTH: 240; 
                TOP: 115;
                COLOR: #000000;
                BACKGROUND-COLOR: <%=gstrPageColor%>" 
                tabIndex=2
                cols=26 NAME="txtOrgTitle"></TEXTAREA> 
                
        <SPAN class=DefLabel 
            id=lblDescription
            style="POSITION: absolute;
                LEFT: 350; 
                WIDTH: 150; 
                FONT-WEIGHT: bold;
                TOP: 140">Report Description:
        </SPAN>
        <TEXTAREA id=txtDescription
            title="Report Description"
            style="POSITION: absolute;
				WIDTH:240; 
				HEIGHT:115; 
				overflow:auto
                TEXT-ALIGN: left;
                LEFT:350; 
                TOP: 155" 
                tabIndex=4
                cols=26 NAME="txtDescription"></TEXTAREA> 

         <BUTTON class=DefBUTTON
            id=cmdSave title="Save changes for the current selected list" 
            style="LEFT:630;
                POSITION: absolute;
                TOP: 60;
                WIDTH: 70;
                HEIGHT: 20"
            accessKey=S
            disabled=true
            tabIndex=7><U>S</U>ave
        </BUTTON>
        <BUTTON class=DefBUTTON
            id=cmdCancel title="Cancel the current change" 
            style="LEFT: 630;
                POSITION: absolute;
                TOP:100;
                HEIGHT: 20;
                WIDTH: 70"
            accessKey=C
            disabled=true
            tabIndex=7>Cance<U>l</U>
        </BUTTON>
        <BUTTON class=DefBUTTON
            id=cmdDisable title="Disable the selected value" 
            style="LEFT:630;
                POSITION: absolute;
                TOP: 140;
                HEIGHT: 20;
                WIDTH: 70;
				visibility:hidden"
			onclick="cmdAble_onclick(0)"
            accessKey=D
            tabIndex=7><U>D</U>isable
        </BUTTON>
        
        <BUTTON class=DefBUTTON
            id="cmdEnable" title="Enable the selected value" 
            style="LEFT:630;
                POSITION: absolute;
                TOP: 180;
                HEIGHT: 20;
                WIDTH: 70;
                visibility:hidden"
            onclick="cmdAble_onclick(1)"
            accessKey=E
            tabIndex=7><U>E</U>nable
        </BUTTON>      
       
        <BUTTON class=DefBUTTON
            id=cmdClose title="Abandon changes" 
            style="LEFT: 630;
                POSITION: absolute;
                TOP: 220;
                HEIGHT: 20;
                WIDTH: 70"
            accessKey=C
            tabIndex=8><U>C</U>lose
        </BUTTON>
    
    </DIV>

</BODY>
<FORM NAME="Form" METHOD="Post" STYLE="VISIBILITY: hidden" ACTION="ReportEdit.ASP" ID=Form>
    <%
    Call CommonFormFields()
    Response.Write "<INPUT TYPE=""hidden"" Name=""FormAction"" VALUE="""" ID=FormAction>"
    Response.Write "<INPUT TYPE=""hidden"" Name=""RptID"" VALUE=""" & ReqForm("RptID") & """ ID=RptID>"
    Response.Write "<INPUT TYPE=""hidden"" Name=""RptTitle"" VALUE=""" & ReqForm("RptTitle") & """ ID=RptTitle>"
    Response.Write "<INPUT TYPE=""hidden"" Name=""RptDescription"" VALUE=""" & ReqForm("RptDescription") & """ id=RptDescription>"
    %>
</FORM>
</HTML>
<%
gadoCon.Close
Set gadoCon = Nothing
%>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncNavigateControls.asp"-->