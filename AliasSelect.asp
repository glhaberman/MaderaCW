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
Dim mdctAliasIDs, moAliasID

mstrPageTitle = "Upper Level Management"
Set adRs = Server.CreateObject("ADODB.Recordset")
Set madoCmd = GetAdoCmd("spGetAlaisIDs")
    AddParmIn madoCmd, "@ID", adInteger, 0, NULL
    AddParmIn madoCmd, "@TypeID", adInteger, 0, NULL
    AddParmIn madoCmd, "@ParentID", adInteger, 0, NULL
    AddParmIn madoCmd, "@Name", adVarchar, 50, NULL
adRs.Open madoCmd, , adOpenForwardOnly, adLockReadOnly
Set madoCmd = Nothing 
Set mdctAliasIDs = CreateObject("Scripting.Dictionary")
Do While Not adRs.EOF
    If mdctAliasIDs.Exists(adRs.Fields("alsID").Value) Then
        mdctAliasIDs(adRs.Fields("alsID").Value) = mdctAliasIDs(adRs.Fields("alsID").Value) & adRs.Fields("ParentID").Value & "*"
    Else
        mdctAliasIDs.Add adRs.Fields("alsID").Value, adRs.Fields("alsTypeID").Value & "^" & _
            adRs.Fields("alsName").Value & "^" & _
            adRs.Fields("ParentID").Value & "*"
    End If
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
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnSetFocusToMain  <%'Flag used througout page to determine if focus should be set to main when window is closed.%>
Dim mdctAliasIDs

Sub window_onload()
    Call SizeAndCenterWindow(767, 520, True)
    
    mblnSetFocusToMain = True
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>
    Set mdctAliasIDs = CreateObject("Scripting.Dictionary")
   <%
    For Each moAliasID In mdctAliasIDs
        Response.Write "mdctAliasIDs.Add " & moAliasID & ", """ & mdctAliasIDs(moAliasID) & """" & vbCrLf
    Next
   %>
    Call cboAliasType_onChange
    'stop
    If Form.AliasID.value <> 0 Then
        For intI = 0 To cboAliasID.options.length - 1
            If CLng(cboAliasID.options(intI).value) = CLng(Form.AliasID.value) Then
                cboAliasID.selectedIndex = intI
                Call cboAliasID_onchange()
     'msgbox cboAliasID.options(intI).text
                Exit For
            End If
        Next
    End If
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

Sub cboAliasType_onchange()
    Dim intI
    Dim strOptionText
    Dim oOption
    Dim oAliasID
    Dim strRecord

    cboAliasID.options.length = Null
    For Each oAliasID In mdctAliasIDs
        strRecord = mdctAliasIDs(oAliasID)
        If CInt(Parse(strRecord,"^",1)) = CInt(cboAliasType.Value) Then
            Set oOption = Document.createElement("OPTION")
            oOption.Value = oAliasID
            oOption.Text = Parse(strRecord,"^",2)
            cboAliasID.options.add oOption
            Set oOption = Nothing
        End If
    Next
    
    cboAliasID.selectedIndex = 0
    Call cboAliasID_onChange
End Sub

Sub cboAliasID_onChange()
    Dim lngID, lngTypeID
    Dim strName, strParentIDs
    Dim strRecord

    strRecord = mdctAliasIDs(CLng(cboAliasID.value))
    lngID = cboAliasID.value
    lngTypeID = cboAliasType.value
    strName = Parse(strRecord,"^",2)
    strParentIDs = Parse(strRecord,"^",3)
    Form.AliasID.value = cboAliasID.value

    fraRev.frameElement.src = "AliasAddEdit.asp?CalledFrom=Select&ID=" & lngID & _
        "&Name=" & strName & _
        "&ParentIDs=" & strParentIDs & _
        "&TypeID=" & lngTypeID
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
        cboAliasType.style.top = 62
        lblValues.style.top = 85
        cboAliasID.style.top = 100
        cboAliasID.style.height = 167
    Else
        lblProgram.style.top = 20
        cboAliasType.style.top = 35
        lblValues.style.top = 70
        cboAliasID.style.top = 85
        cboAliasID.style.height = 176
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
                Management Type
            </SPAN>        
            
            <SELECT id=cboAliasType TYPE="select-one"
                style="LEFT:10; WIDTH:220; TOP:35" tabIndex=1 NAME="cboAliasType">
                <option value=125>Office Managers</option>
                <option value=126>Regions</option>
                <option value=250>FIPs</option>
            </SELECT>
            
            <SPAN id=lblValues class=DefLabel style="LEFT:10; WIDTH:185; TOP:70">
                Values in Management Type:
            </SPAN>
            <SELECT id=cboAliasID title="Values in selected list" TYPE="select-one"
                style="LEFT:10; WIDTH:220; TOP:85; HEIGHT:176" 
                onkeydown="Gen_onkeydown"
                tabIndex=2 size=13 NAME="cboAliasID">
            </SELECT>
            
            <IFRAME ID=fraRev 
                SRC="AliasAddEdit.asp?AliasID=0&AliasName=""" 
                STYLE="Left:10; WIDTH:650; HEIGHT:340;background-color:<%=gstrbackcolor%>" 
                FRAMEBORDER=0 ></IFRAME>

            <BUTTON id=cmdClose class=DefBUTTON title="Close the form" 
                style="LEFT:575; TOP:310; HEIGHT: 20; WIDTH:70"
                tabIndex=3>Close
            </BUTTON>
        </DIV>
    </DIV>
</BODY>

<FORM NAME="Form" METHOD="Post" STYLE="VISIBILITY: hidden" ACTION="Main.ASP" ID=Form>
    <%
    Call CommonFormFields()
    WriteFormField "AliasTypeID", ReqForm("AliasTypeID")
    WriteFormField "AliasTypeName", ReqForm("AliasTypeName")
    WriteFormField "FormAction", ""
    WriteFormField "AliasID", ReqForm("AliasID")
    WriteFormField "AliasName", ReqForm("AliasName")
    WriteFormField "EditID", ReqForm("EditID")
    gadoCon.Close
    Set gadoCon = Nothing
    %>
</FORM>
</HTML>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncNavigateControls.asp"-->
