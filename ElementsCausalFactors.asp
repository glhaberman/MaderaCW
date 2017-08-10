<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: FactorAddEditAssign.asp                                         '
'  Purpose: The primary admin data entry screen for maintaining the causal  '
'           factors for each eligibility element.                           '
'           This form is only available to admin users.                     '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<%
Dim mstrWhoCalled
Dim madoRs, mstrProgramList

mstrWhoCalled = Request.Form("WhoCalled")

Set madoRs = Server.CreateObject("ADODB.Recordset")
Set gadoCmd = GetAdoCmd("spGetProgramList")
    AddParmIn gadoCmd, "@PrgID", adVarchar, 255, NULL
    madoRs.Open gadoCmd, , adOpenForwardOnly, adLockReadOnly
Set gadoCmd = Nothing

mstrProgramList =  "<option value=0></option>"
madoRs.Filter = "prgID<50"
Do While Not madoRs.EOF
    mstrProgramList = mstrProgramList & "<option value=" & madoRs.Fields("prgID").Value & ">" & madoRs.Fields("prgShortTitle").Value & "</option>"
    madoRs.MoveNext
Loop
madoRs.Close

%>
<HTML>
<HEAD>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Option Explicit

Dim mdctWindows        <%'Holds reference to Print windows when they are opened.%>
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnSetFocusToMain  <%'Flag used througout page to determine if focus should be set to main when window is closed.%>

Sub window_onload()
    Dim intI
    
    mblnSetFocusToMain = True
    If "<%=mstrWhoCalled%>" = "Factors" Then
        Call ShowDiv("Factors")
    Else
        divElementsLoading.style.left = -1000
        divPageFrame.style.left = 0
        divHeader.style.left = 0
        Call ShowDiv("SubMenu")
        cboTabID.value = 2 'Installation does not use Tabs, so default to Tab 2.
        Call cboTabID_onchange()
    End If
End Sub

Sub window_onbeforeunload()
    If mblnSetFocusToMain = True Then
        window.opener.focus
    End If
End Sub

Function TabAndFunctionSelected()
    Dim strMissing
    
    strMissing = ""
    If cboTabID.value = 0 Then
        strMissing = "Tab"
    End If
    If cboProgramID.value = 0 Then
        If strMissing = "" Then
            strMissing = "Program"
        Else
            strMissing = strMissing & vbCrLf & Space(10) & "Program"
        End If
    End If
    If cboProgramID.value = 6 And cboTabID.value = 2 And cboElementID.value = 0 Then
        If strMissing = "" Then
            strMissing = "Action"
        Else
            strMissing = strMissing & vbCrLf & Space(10) & "Action"
        End If
    End If
    If strMissing <> "" Then
        MsgBox "The following item(s) have not been selected: " & vbCrLf & Space(10) & strMissing, vbOkOnly, "Case Review System Admin"
        TabAndFunctionSelected = False
    Else
        TabAndFunctionSelected = True
    End If
End Function

Sub ShowDiv(strShow)
    If strShow = "Elements" Then
        If TabAndFunctionSelected = False Then Exit Sub
        divFactors.style.left = -1000
        divHeader.style.left = -1000
        divPageFrame.style.left = -1000
        divElementsLoading.style.left = 0
        If cboProgramID.value = 6 And cboTabID.value = 2 Then
            ' Enforment Remedies (programID=6) has sub-programs for Data Integrity.  
            Form.ElementID.value = cboElementID.value            
            If fraElements.Form.ProgramID.value = cboElementID.value And fraElements.Form.TabID.value = cboTabID.value Then
            Else
                fraElements.frameElement.src = "ElementAddEdit.asp?" & _
                "ProgramID=" & cboElementID.value & _
                "&ProgramName=" & cboElementID.options(cboElementID.selectedIndex).text & _
                "&TabID=" & cboTabID.value
            End If
        Else
            If fraElements.Form.ProgramID.value = cboProgramID.value And fraElements.Form.TabID.value = cboTabID.value Then
            Else
                fraElements.frameElement.src = "ElementAddEdit.asp?" & _
                "ProgramID=" & cboProgramID.value & _
                "&ProgramName=" & cboProgramID.options(cboProgramID.selectedIndex).text & _
                "&TabID=" & cboTabID.value
            End If
        End If
    ElseIf strShow = "Factors" Then
        'If TabAndFunctionSelected = False Then Exit Sub
        divElements.style.left = -1000
        divFactors.style.left = 0
        divHeader.style.left = -1000
        divPageFrame.style.left = -1000
        divFactors.style.left = 0
    Else
        divElements.style.left = -1000
        divFactors.style.left = -1000
        divHeader.style.left = 0
        divPageFrame.style.left = 0
    End If
    Call cboTabID_onchange()
End Sub

Sub cboTabID_onchange()
    Select Case cboTabID.value
        Case 0
            cmdElements.value = "Edit Items"
        Case 1
            cmdElements.value = "Edit Actions"
        Case 2
            cmdElements.value = "Edit Elements"
        Case 3
            cmdElements.value = "Edit Questions"
    End Select
    Call CheckActions()
End Sub

Sub cboProgramID_onChange()
    Call CheckActions()
End Sub

Sub CheckActions()
    Dim dctReturn, oReturn
    Dim oOption

    If cboTabID.value = "2" And cboProgramID.value = "6" Then
        cboElementID.options.length = Null
        Set oOption = Document.createElement("OPTION")
        oOption.Value = 0
        oOption.Text = ""
        cboElementID.options.add oOption
        Set oOption = Nothing
        Set dctReturn = window.showModalDialog("ElementSort.asp?Action=GetActions&ProgramID=6&TabID=1", , "dialogWidth:210px;dialogHeight:120px;scrollbars:no;center:yes;border:thin;help:no;status:no")
        If dctReturn.Count > 0 Then
            For Each oReturn In dctReturn
                Set oOption = Document.createElement("OPTION")
                oOption.Value = Parse(dctReturn(oReturn),"^",1)
                oOption.Text = Parse(dctReturn(oReturn),"^",2)
                cboElementID.options.add oOption
                Set oOption = Nothing
            Next
        End If
        cboElementID.style.left = 10
        lblElementID.style.left = 10
        cmdElements.style.top = 200
        'cmdFactors.style.top = 200
        cboElementID.value = Form.ElementID.value
    Else
        cboElementID.style.left = -1000
        lblElementID.style.left = -1000
        cmdElements.style.top = 150
        'cmdFactors.style.top = 150
    End If
End Sub

Sub cmdFactors_onclick()
    Call ShowDiv("Factors")
End Sub

Sub cmdElements_onclick()
    Call ShowDiv("Elements")
End Sub

Sub cmdPrint_onclick()
    Call window.showModalDialog("RptElementFactorList.asp?ProgramID=" & cboProgramID.Value,, "dialogWidth:770px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")
End Sub

Sub cmdClose_onclick()
    Call window.opener.ManageWindows(6,"Close")
End Sub

Sub CloseMe()
End Sub
</SCRIPT>
<!--#include file="IncCmnCliFunctions.asp"-->

<BODY ID=PageBody style="OVERFLOW: auto; POSITION: absolute; BACKGROUND-COLOR: <%=gstrPageColor%>" 
    bottomMargin=0 leftMargin=0 topMargin=0 rightMargin=0>
    
    <DIV id=divElements style="POSITION: absolute;HEIGHT:476;WIDTH:737;LEFT:-1000;TOP:0;">
        <IFRAME ID=fraElements src="ElementAddEdit.asp?ProgramID=0&TabID=0"
            STYLE="positon:absolute;LEFT:0;WIDTH:735;HEIGHT:474;TOP:0;BORDER-style:none" FRAMEBORDER=0>
        </IFRAME>
    </DIV>
    <DIV id=divElementsLoading style="POSITION: absolute;HEIGHT:476;WIDTH:737;LEFT:0;TOP:0;">
        <SPAN id=lblDatabaseStatus class=DefLabel
            STYLE="LEFT:5; WIDTH:200; TOP:100; TEXT-ALIGN:left">
            Accessing Database...
        </SPAN>
    </DIV>
    <DIV id=divFactors style="POSITION: absolute;HEIGHT:476;WIDTH:737;LEFT:-1000;TOP:0;">
        <IFRAME ID=fraFactors src="FactorAddEdit.asp?Action=Load"
            STYLE="positon:absolute;LEFT:0;WIDTH:735;HEIGHT:474;TOP:0;BORDER-style:none" FRAMEBORDER=0>
        </IFRAME>
    </DIV>
    <DIV id=divHeader
        style="POSITION: absolute;BORDER-STYLE: solid;BORDER-WIDTH: 1px;BORDER-COLOR: <%=gstrBorderColor%>;
        BACKGROUND-COLOR: <%=gstrBackColor%>;COLOR: black;HEIGHT: 40;WIDTH: 730;LEFT:-1000;TOP:10">
       
        <SPAN id=lblAppTitle
            style="POSITION: absolute;FONT-SIZE: <%=gstrTitleFontSize%>;
            FONT-FAMILY: <%=gstrTitleFont%>;COLOR: <%=gstrTitleColor%>;
            HEIGHT: 20;TOP: 4;LEFT: 6;WIDTH: 720;TEXT-ALIGN: center;FONT-WEIGHT: bold">Case Review Elements</SPAN>
    </DIV>

    <DIV id=divPageFrame Class=ControlDiv style="WIDTH:730;TOP:51;height:420;LEFT:-1000;BORDER-STYLE: solid;BORDER-WIDTH: 1px;BORDER-COLOR: <%=gstrBorderColor%>;
        BACKGROUND-COLOR: <%=gstrBackColor%>;">
        <SPAN id=lblinstructions class=DefLabel style="LEFT:50; TOP:10; WIDTH:600">
            Select a Program for editing.  Click on the &lt;Edit Elements&gt; button to modify the list of elements for the selected program.
        </SPAN>
        
        <SPAN id=lblTabID class=DefLabel style="display:none; LEFT:20; TOP:50; WIDTH:200">
            Tab
        </SPAN>        
        
        <SELECT id=cboTabID TYPE="select-one"
            style="display:none; LEFT:20; WIDTH:220; TOP:65" tabIndex=0 NAME="cboTabID">
            <option value=0></option>
            <option value=1>Action Integrity</option>
            <option value=2>Data Integrity</option>
            <option value=3>Information Gathering</option>
        </SELECT>
        
        <SPAN id=lblProgramID class=DefLabel style="LEFT:20; TOP:100; WIDTH:200">
            Select a Program
        </SPAN>        
        
        <SELECT id=cboProgramID TYPE="select-one"
            style="LEFT:20; WIDTH:220; TOP:115" tabIndex=0 NAME="cboProgramID">
            <%=mstrProgramList%>
        </SELECT>
        
        <SPAN id=lblElementID class=DefLabel style="LEFT:-1020; TOP:150; WIDTH:200">
            Action
        </SPAN>        
        <SELECT id=cboElementID TYPE="select-one"
            style="LEFT:-1020; WIDTH:220; TOP:165" tabIndex=0 NAME="cboElementID">
            <option selected value=0></option>
        </SELECT>
        
        <BUTTON class=DefBUTTON id=cmdElements 
            style="LEFT:20; POSITION: absolute; TOP: 150;HEIGHT: 20;WIDTH: 105" 
            tabIndex=1>Edit Items
        </BUTTON>

        <BUTTON class=DefBUTTON id=cmdPrint 
            style="LEFT:145; POSITION: absolute; TOP: 150;HEIGHT: 20;WIDTH: 105" 
            tabIndex=1>Print Elements
        </BUTTON>
        
        <BUTTON class=DefBUTTON id=cmdFactors
            style="LEFT: -1125;POSITION: absolute; TOP: 150;HEIGHT: 20;WIDTH: 105" 
            tabIndex=2>Factors
        </BUTTON>
        <BUTTON class=DefBUTTON id=cmdClose title="Close" style="LEFT: 645;POSITION: absolute;
                TOP: 390;HEIGHT: 20;WIDTH: 70" accessKey=R
            tabIndex=3>Close
        </BUTTON>
    </DIV>
</BODY>
<FORM NAME="Form" METHOD="Post" STYLE="VISIBILITY: hidden" ACTION="" ID=Form>
<%
    Call CommonFormFields()
    WriteFormField "WhoCalled", mstrWhoCalled
    WriteFormField "ElementID", 0
%>
</FORM>
</HTML>
<%
gadoCon.Close
Set gadoCon = Nothing
%>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
