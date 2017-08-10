<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: AppOptionSelect.asp                                             '
'  Purpose: This screen allows the system admin user to modify application  '
'           options and settings.                                           '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<%
Dim madoRs
Dim strSQL
Dim mstrPageTitle
Dim mstrOptions
Dim mintLine
Dim mstrHtml
Dim mstrRowStart
Dim mstrRecords
Dim strTmp

mstrPageTitle = "Select Option to Edit"

'If this is a post-back, save the information to the database:
If ReqForm("FormAction") = "Save" Then
    Set gadoCmd = GetAdoCmd("spUpdOption")
        Call AddParmIn(gadoCmd, "@SettingName", adVarChar, 50, ReqForm("SettingName"))
        Call AddParmIn(gadoCmd, "@SettingValue", adVarChar, 255, ReqForm("SettingValue"))
        Call AddParmIn(gadoCmd, "@Description", adVarChar, 255, ReqForm("Description"))
        gadoCmd.Execute
    Set gadoCmd = nothing
End If 

Set madoRs = Server.CreateObject("ADODB.Recordset")
Set gadoCmd = GetAdoCmd("spGetOptionNames")
    madoRs.Open gadoCmd, , adOpenForwardOnly, adLockReadOnly
    
mstrRecords = "~"
Do While Not madoRs.EOF
    mstrRecords = mstrRecords & madoRs("SettingName").Value & "*" & _
        madoRs("SettingValue").Value & "*" & _
        madoRs("SettingGlobal").Value & "*" & _
        madoRs("InputMask").Value & "*" & _
        madoRs("Description").Value & "*" & _
        madoRs("ClientCanEdit").Value & "*" & _
        madoRs("SettingValueCaption").Value & "*" & _
        madoRs("MaxLength").Value & "*" & _
        madoRs("Explanation").Value & "*" & _
        madoRs("Category").Value & "~"
    
    madoRs.MoveNext
Loop
madoRs.Close

%>
<HTML>
<HEAD>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE><%=gstrAppName%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncTableStyles.asp"-->
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Dim mintSelectedRow 'Keeps track of which item is selected in the list.
Dim mdctAppSettings
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnSetFocusToMain  <%'Flag used througout page to determine if focus should be set to main when window is closed.%>

Sub window_onload()

    Call SizeAndCenterWindow(555, 335, False)
    mblnSetFocusToMain = True
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>

    Set mdctAppSettings = CreateObject("Scripting.Dictionary")
    Call LoadDictionary()
    Call LoadTable()
    
    mintSelectedRow = -1
End Sub

<%'If timer detects that Main has been closed, this sub will be called. %>
Sub MainClosed()
    mblnSetFocusToMain = False
    window.close
End Sub

Sub window_onbeforeunload()
    If mblnSetFocusToMain = True Then
        window.opener.focus
    End If
End Sub

Sub LoadDictionary()
    Dim strValue
    Dim intLast
    Dim intCnt
    Dim intDelim
    Dim strRecords

    strRecords ="<% = mstrRecords %>"

    If Len(strRecords) <= 1 Then Exit Sub
	intLast = 1
	intCnt = -1
	' Load array of items from string value
	Do While True
		intDelim = Instr(intLast + 1,strRecords,"~")
		strValue = Mid(strRecords, intLast + 1, intDelim - (intLast + 1))
		intCnt = intCnt + 1
        mdctAppSettings.Add intCnt,strValue
		intLast = intDelim
		If intLast = Len(strRecords) Then Exit Do
	Loop
End Sub

Sub LoadTable()
    Dim strHTML
    Dim intI
    Dim strTable
    Dim strRecord
    Dim blnSysAdmin
    Dim strColor
    Dim strCategory
    Dim intCat
    Dim strDisplayCategory

'SettingName	1
'SettingValue	2
'SettingGlobal	3
'InputMask	4
'Description	5
'ClientCanEdit	6
'SettingValueCaption	7
'MaxLength	8
'Explanation	9
'Category	10
    
    strHTML = tblSettings.outerHTML
    intEnd = InStr(strHTML,"</THEAD>")

    strTable = "</THEAD>" & vbCrLf
    intI = 0
    intCat = 0

    strCategory = ".."
    For Each strRecord In mdctAppSettings
        If strCategory <> Parse(mdctAppSettings(strRecord),"*",10) Then
            ' New Category, add a row to the table as a heading
            strDisplayCategory = Parse(mdctAppSettings(strRecord),"*",10)
            If InStr(strDisplayCategory,"^") > 0 Then strDisplayCategory = Parse(strDisplayCategory,"^",2)
            strTable = strTable & "<TR ID=tbrSettingsH" & intCat & " class=TableDetail > " & vbCrLf
            strTable = strTable & "<TD ID=tbcSettingsC0RH" & intCat & " class=TableDetail style=""font-weight:bold;color:black;background-color:khaki"">" & strDisplayCategory & "</TD>" & vbCrLf
            strTable = strTable & "<TD ID=tbcSettingsC1RH" & intCat & " class=TableDetail style=""color:beige"">" & "" & "</TD>" & vbCrLf
            strTable = strTable & "</TR>"
            intCat = intCat + 1
            strCategory = Parse(mdctAppSettings(strRecord),"*",10)
        End If
        strColor = "silver"
        If Parse(mdctAppSettings(strRecord),"*",6) = "1" Then strColor = "black"
        strTable = strTable & "<TR ID=tbrSettings" & intI & " class=TableDetail onclick=Result_onclick(" & intI & ") > " & vbCrLf
        strTable = strTable & "<TD ID=tbcSettingsC0R" & intI & " class=TableDetail style=""cursor:hand;padding-left:10;color:black"">" & Parse(mdctAppSettings(strRecord),"*",5) & "</TD>" & vbCrLf
        strTable = strTable & "<TD ID=tbcSettingsC1R" & intI & " class=TableDetail style=""cursor:hand;color:" & strColor & """>" & Parse(mdctAppSettings(strRecord),"*",2) & "</TD>" & vbCrLf
        strTable = strTable & "</TR>"
        intI = intI + 1
    Next
    strTable = strTable & "</TABLE>" & vbCrLf
        
    tblSettings.outerHTML = Left(strHTML,intEnd-1) & strTable
End Sub

Sub Result_onclick(intRow)
    Dim strRow
    Dim intI
    Dim oOption
    Dim strInputMask
    Dim strInputMaskType
    Dim strInputMask2
    Dim strInputMaskType2
    Dim strOptionValue
    Dim strValue
    Dim strText

    If Not IsObject(tblSettings) Then
        Exit Sub 
    End If
    If PageFrame.disabled Or tblSettings.Rows.Length = 0 Then
        Exit Sub
    End If
    
    If mintSelectedRow >= 0 Then
        ' Check if there are unsaved changes
        strInputMask = Parse(mdctAppSettings(mintSelectedRow),"*",4)
        If InStr(strInputMask,"|") > 0 Then
            strValue = cboSettingValue.value & txtSettingValue.value
        Else
            If txtSettingValue.style.visibility = "visible" Then
                strValue = txtSettingValue.value
            ElseIf cboSettingValue.style.visibility = "visible" Then
                strValue = cboSettingValue.value
            End If
        End If

        If txtDescription.value <> Parse(mdctAppSettings(mintSelectedRow),"*",5) Or _
            strValue <> Parse(mdctAppSettings(mintSelectedRow),"*",2) Then
            
            intI = MsgBox("Changes for the setting " & Replace(Parse(mdctAppSettings(mintSelectedRow),"*",7),"^"," / ") & " have not been saved.  If you continue you will lose these changes.  Continue?",vbYesNo,"Application Settings")
            If intI = vbNo Then Exit Sub
        End If

        strRow = "tbrSettings" & mintSelectedRow
        tblSettings.Rows(strRow).className = "TableRow"    
        tblSettings.Rows(strRow).cells(0).tabindex = -1
        tblSettings.Rows(strRow).cells(0).style.color = "black"
        If tblSettings.Rows(strRow).cells(1).style.color <> "silver" Then tblSettings.Rows(strRow).cells(1).style.color = "black"
    End If
    strValue = ""
    
    strRow = "tbrSettings" & intRow
    tblSettings.Rows(strRow).className = "TableSelectedRow"    
    tblSettings.Rows(strRow).cells(0).focus
    tblSettings.Rows(strRow).cells(0).tabindex = 9
    tblSettings.Rows(strRow).cells(0).style.color = "white"
    If tblSettings.Rows(strRow).cells(1).style.color <> "silver" Then tblSettings.Rows(strRow).cells(1).style.color = "white"

    mintSelectedRow = intRow
    txtDescription.value = Parse(mdctAppSettings(intRow),"*",5)
    txtExplanation.value = Parse(mdctAppSettings(intRow),"*",9)
    lblSettingValue2.innerHTML = Parse(mdctAppSettings(intRow),"*",7)
    divSelectColor.style.visibility = "hidden"
    cmdColor.style.visibility = "hidden"
    cboSettingValue.style.visibility = "hidden"
    txtSettingValue.style.visibility = "hidden"
    lblSettingValue.style.visibility = "visible"
    lblSettingValue2.style.visibility = "visible"
    lblSettingValue3.style.visibility = "hidden"
    txtSettingValue.style.width = 260
    txtSettingValue.style.left = 1
    txtSettingValue.disabled = False
    cboSettingValue.disabled = False
    cmdColor.disabled = False
    
    strInputMask = Parse(mdctAppSettings(intRow),"*",4)
    If InStr(strInputMask,"^") > 0 Then
        If InStr(strInputMask,"|") > 0 Then
            strInputMaskType = "MULTILIST"
        Else
            strInputMaskType = "SINGLELIST"
        End If
    Else
        strInputMaskType = strInputMask
        Form.InputMask.value = strInputMask
        If IsNumeric(strInputMask) Then Form.InputMask.value = "NUMBER"
    End If
    Select Case strInputMaskType
        Case "COLOR"
            txtSettingValue.style.visibility = "visible"
            txtSettingValue.value = Parse(mdctAppSettings(intRow),"*",2)
            txtSettingValue.style.width = 322
            cmdColor.style.visibility = "visible"
            txtSettingValue.focus
            Form.InputMask.value = "TEXT"
        Case "FONTSIZE"
            cboSettingValue.style.visibility = "visible"
            cboSettingValue.options.length = Null
            For intI = 6 To 48
                Set oOption = Document.createElement("OPTION")
                oOption.Value = intI & "pt"
                oOption.Text = intI & "pt"
                cboSettingValue.add oOption
            Next
            
            cboSettingValue.value = Parse(mdctAppSettings(intRow),"*",2)
            cboSettingValue.style.width = 70
            cboSettingValue.focus
        Case "BOOLEAN"
            cboSettingValue.style.visibility = "visible"
            cboSettingValue.options.length = Null
            Set oOption = Document.createElement("OPTION")
            oOption.Value = ""
            oOption.Text = ""
            cboSettingValue.add oOption
            Set oOption = Document.createElement("OPTION")
            oOption.Value = "Yes"
            oOption.Text = "Yes"
            cboSettingValue.add oOption
            Set oOption = Document.createElement("OPTION")
            oOption.Value = "No"
            oOption.Text = "No"
            cboSettingValue.add oOption
            
            cboSettingValue.value = Parse(mdctAppSettings(intRow),"*",2)
            cboSettingValue.style.width = 70
            cboSettingValue.focus
        Case "SINGLELIST"
            cboSettingValue.style.visibility = "visible"
            cboSettingValue.options.length = Null
            strInputMask = strInputMask & "^*~*"
            intI = 1
            Do While True
                strOptionValue = Parse(strInputMask,"^",intI)
                If strOptionValue = "*~*" Then Exit Do

                Set oOption = Document.createElement("OPTION")
                If InStr(strOptionValue,"+") > 0 Then
                    strValue = Parse(strOptionValue,"+",1)
                    strText = Parse(strOptionValue,"+",2)
                Else
                    strValue = strOptionValue
                    strText = strOptionValue
                End If
                oOption.Value = strValue
                oOption.Text = strText
                cboSettingValue.add oOption
                
                intI = intI + 1
            Loop
            cboSettingValue.value = Parse(mdctAppSettings(intRow),"*",2)
            cboSettingValue.style.width = 200
            cboSettingValue.focus
        Case "MULTILIST"
            strInputMask2 = Parse(strInputMask,"|",2)
            strInputMask = Parse(strInputMask,"|",1)
            
            lblSettingValue3.innerHTML = Parse(lblSettingValue2.innerHTML,"^",2)
            lblSettingValue2.innerHTML = Parse(lblSettingValue2.innerHTML,"^",1)
            cboSettingValue.style.visibility = "visible"
            lblSettingValue3.style.visibility = "visible"
            cboSettingValue.options.length = Null
            strInputMask = strInputMask & "^*~*"
            intI = 1
            Do While True
                strOptionValue = Parse(strInputMask,"^",intI)
                If strOptionValue = "*~*" Then Exit Do

                Set oOption = Document.createElement("OPTION")
                oOption.Value = Parse(strOptionValue,"+",1)
                oOption.Text = Parse(strOptionValue,"+",2)
                cboSettingValue.add oOption
                
                intI = intI + 1
            Loop
            cboSettingValue.style.width = 170
            cboSettingValue.focus

            txtSettingValue.style.visibility = "visible"
            txtSettingValue.style.width = 50
            txtSettingValue.style.left = 190
            txtSettingValue.maxLength = Len(strInputMask2)

            If Len(Parse(mdctAppSettings(intRow),"*",2)) > 1 Then
                cboSettingValue.value = Left(Parse(mdctAppSettings(intRow),"*",2),1)
                txtSettingValue.value = Mid(Parse(mdctAppSettings(intRow),"*",2),2,Len(Parse(mdctAppSettings(intRow),"*",2)) - 1)
            End If
            Form.InputMask.value = "NUMBER"

        Case Else
            If InStr(strInputMaskType,"9") > 0 Then
                txtSettingValue.maxLength = Len(strInputMaskType)
                If Len(strInputMaskType) <= 23 Then
                    txtSettingValue.style.width = Len(strInputMaskType) * 15
                Else
                    txtSettingValue.style.width = 345
                End If
            ElseIf strInputMaskType = "DATE" Then
                txtSettingValue.maxLength = 10
                txtSettingValue.style.width = 90
            Else
                txtSettingValue.maxLength = 255
                txtSettingValue.style.width = 345
                If IsNumeric(Parse(mdctAppSettings(intRow),"*",8)) Then txtSettingValue.maxLength = Parse(mdctAppSettings(intRow),"*",8)
            End If
            txtSettingValue.style.visibility = "visible"
            txtSettingValue.value = Parse(mdctAppSettings(intRow),"*",2)
            txtSettingValue.focus
    End Select

    If Parse(mdctAppSettings(intRow),"*",6) = "1" Then
    Else
        txtSettingValue.disabled = True
        cboSettingValue.disabled = True
        If cmdColor.style.visibility = "visible" Then cmdColor.disabled = True
    End If
End Sub

Sub txtSettingValue_onkeypress()
    Dim strType
    Dim strMask

    strType = ""
    If window.event.keyCode >= 48 And window.event.keyCode <=57 Then
        strType = "NUMBER"
    ElseIf window.event.keyCode >= 65 And window.event.keyCode <=90 Then
        strType = "TEXT"
    ElseIf window.event.keyCode >= 97 And window.event.keyCode <=122 Then
        strType = "TEXT"
    ElseIf window.event.keyCode = 32 Then
        strType = "SPACE"
    End If
    
    Select Case Form.InputMask.value
        Case "DATE"
            If txtSettingValue.value = "(MM/DD/YYYY)" Then
                txtSettingValue.value = ""
            End If
            Call TextBoxOnKeyPress(window.event.keyCode,"D")
        Case "NUMBER"
            If strType <> "NUMBER" Then window.event.keyCode = 0
        Case "TEXT"
            If strType = "" Then window.event.keyCode = 0
        Case Else
            If InStr(Form.InputMask.value,"TEXT") > 0 Then
                strMask = Mid(Form.InputMask.value,5,Len(Form.InputMask.value) - 4)
                
                If Len(txtSettingValue.value) = Len(strMask) Then
                    window.event.keyCode = 0
                    Exit Sub
                End If

                Select Case Mid(strMask,Len(txtSettingValue.value) + 1, 1)
                    Case "X"
                        If strType <> "TEXT" Then window.event.keyCode = 0
                    Case "9"
                        If strType <> "NUMBER" Then window.event.keyCode = 0
                End Select
            End If
    End Select
End Sub

Sub txtSettingValue_onblur
    If Form.InputMask.value <> "DATE" Then Exit Sub
    If Trim(txtSettingValue.value) = "(MM/DD/YYYY)" Then
        txtSettingValue.value = ""
        Exit Sub
    End If
    If Not ValidDate(txtSettingValue.value) Then
        MsgBox Parse(mdctAppSettings(mintSelectedRow),"*",7) & " must be a valid date - MM/DD/YYYY.", vbInformation, "Application Setting"
        txtSettingValue.focus
    End If
End Sub

Sub txtSettingValue_onfocus
    If Trim(txtSettingValue.value) = "" And Form.InputMask.value = "DATE" Then
        txtSettingValue.value = "(MM/DD/YYYY)"
    End If
    txtSettingValue.select
End Sub

Sub cmdColor_onclick()
    divSelectColor.style.visibility = "visible"
    fraColors.frameElement.src = "SelectColor.asp"
End Sub
Sub lstResults_onkeydown()
    <%
    'This code controls the behavior in the results DIV when the Up arrow,
    'Down arrow, Home, and End keys are pressed.  This code changes the
    'selected item as the user moves up and down in the list:
    %>
    If IsNumeric(mintSelectedRow) Then
        Select Case Window.Event.keyCode
            Case 36 'home
                Window.event.returnValue = False
                tblSettings.rows(0).scrollIntoView
                Call Result_onclick(1)
            Case 35 'end
				Window.event.returnValue = False
                Call Result_onclick(tblSettings.Rows.Length - 1)
            Case 38 'Up
                If mintSelectedRow > 1 Then
					Window.event.returnValue = False
                    Call Result_onclick(mintSelectedRow - 1)
                End If
            Case 40 'Down
                If Cint(mintSelectedRow) < CInt(tblSettings.Rows.Length - 2) Then
					Window.event.returnValue = False
                    Call Result_onclick(mintSelectedRow + 1)
                End If
        End Select
    End If
End Sub

Sub cmdClose_onclick()
    mblnSetFocusToMain = False
    Form.Action = "Admin.asp"
    Form.Submit
End Sub

Sub cmdSave_onclick()
    Dim strRecord
 
    If mintSelectedRow < 0 Then Exit Sub
    
    strRecord = mdctAppSettings(mintSelectedRow)
    
    Form.SettingName.Value = Parse(strRecord,"*",1)
    strInputMask = Parse(strRecord,"*",4)
    If InStr(strInputMask,"|") > 0 Then
        Form.SettingValue.Value = cboSettingValue.value & txtSettingValue.value
    Else
        If txtSettingValue.style.visibility = "visible" Then
            Form.SettingValue.Value = txtSettingValue.value
        ElseIf cboSettingValue.style.visibility = "visible" Then
            Form.SettingValue.Value = cboSettingValue.value
        End If
    End If
    Form.Description.value = txtDescription.value
    Form.FormAction.Value = "Save"
    Form.Action = "AppOptionSelect.asp"
    mblnSetFocusToMain = False
    Form.Submit
End Sub

Sub Gen_onkeydown
    If window.event.keyCode = 13 Then
        If cmdSave.disabled = false Then
            window.event.keyCode = 0
            Call cmdSave_onclick
        End If
    ElseIf window.event.keyCode = 27 Then
        If cmdCancel.disabled = false Then
            Call cmdClose_onclick
        End If
    End If
End Sub

Sub SelectColor(strColor)
    divSelectColor.style.visibility = "hidden"
    txtSettingValue.style.visibility = "visible"
    If strColor <> "" Then txtSettingValue.value = strColor
End Sub

</SCRIPT>
<%'----------------------------------------------------------------------------
'  Client side include files:
'----------------------------------------------------------------------------%>
<!--#include file="IncCmnCliFunctions.asp"-->

<BODY bottomMargin=5 leftMargin=5 topMargin=5 rightMargin=5>
    
    <DIV id=Header class=DefTitleArea style="WIDTH:737; HEIGHT:40">
    
        <SPAN id=lblAppTitleHiLight class=DefTitleTextHiLight 
            style="WIDTH:737"><%=mstrPageTitle%>
        </SPAN>

        <SPAN id=lblAppTitle class=DefTitleText
            style="WIDTH:737"><%=mstrPageTitle%>
        </SPAN>
        <DIV id=divNavigateButton style="left:4;top:4;font-size:12;width:75;
            FONT-WEIGHT:bold;FONT-FAMILY:tahoma;COLOR:darkolivegreen;">
            Navigate
        </DIV>
    </DIV>

    <% Call WriteNavigateControls(-1,30,gstrBackColor) %>            
    <DIV id=PageFrame class=DefPageFrame style="HEIGHT:425; WIDTH:737; TOP:51">

        <DIV id=lstResults class=TableDivArea
            style="LEFT:10; WIDTH:717; TOP:10; HEIGHT:250" tabIndex=1>
            <%
            Response.Write "<TABLE ID=tblSettings Border=0 Width=680 CellSpacing=0 Style=""overflow: hidden; TOP:0""> " & vbCrLf
            Response.Write "<TBODY ID=tbdSettings> " & vbCrLf
            Response.Write "<THEAD ID=thdSettings>" & vbCrLf
            Response.Write "<TR ID=thrSettings>" & vbCrLf
            Response.Write "<TD class=CellLabel ID=thcSettingCol0 style=""width:380"">Description</TD>" & vbCrLf
            Response.Write "<TD class=CellLabel ID=thcSettingCol1 style=""width:300"">Setting</TD>" & vbCrLf
            Response.Write "</TR>" & vbCrLf
            Response.Write "</THEAD>" & vbCrLf
            Response.Write "</TABLE>" & vbCrLf
            %>
         </DIV>

        <SPAN id=lblDescription class=DefLabel style="LEFT:10; WIDTH:340; TOP:270">
            Setting Description:
            <INPUT type=text id=txtDescription title="Setting Description" style="LEFT:1; WIDTH:350;TOP:15;TEXT-ALIGN:LEFT"
                tabIndex=1 maxlength=500 NAME="txtDescription">
            <SPAN id=lblExplanation class=DefLabel style="LEFT:1; WIDTH:260; TOP:35">
                Explanation:
            </SPAN>
            <TEXTAREA id=txtExplanation title="Explanation of Setting"
                style="TOP:50;LEFT:1; WIDTH:350; BACKGROUND-COLOR:<%=gstrAltBackColor%>;HEIGHT:95;OVERFLOW:AUTO"
                tabIndex=-1 readOnly NAME="txtExplanation"></TEXTAREA>
        </SPAN>
        <SPAN id=lblSettingValue class=DefLabel style="LEFT:380; WIDTH:300; TOP:270;visibility:hidden">
            <SPAN id=lblSettingValue2 class=DefLabel style="LEFT:1; WIDTH:260; TOP:1">
                Setting Value:
            </SPAN>
            <SPAN id=lblSettingValue3 class=DefLabel style="LEFT:190;visibility:hidden;WIDTH:170; TOP:1">
                Setting Value:
            </SPAN>
            <INPUT type=text id=txtSettingValue style="LEFT:1; WIDTH:260;TOP:15;TEXT-ALIGN:LEFT"
                tabIndex=1 maxlength=200 NAME="txtSettingValue">

            <SELECT id=cboSettingValue title="Setting Value"
                style="LEFT:1; TOP:15;WIDTH:195;visibility:hidden" NAME="cboSettingValue">
                <OPTION VALUE=0 SELECTED>
            </SELECT>

            <BUTTON id=cmdColor class=DefBUTTON title="Display Color Chart" 
                style="LEFT:325; TOP:14; WIDTH:20;height:19"
                tabIndex=7>...
            </BUTTON>
        </SPAN>
        <DIV id=divSelectColor class=DefPageFrame style="LEFT:320; HEIGHT:105; WIDTH:405; TOP:270; 
            visibility:hidden; BORDER-COLOR: <%=gstrDefButtonColor%>">
            <IFRAME ID=fraColors src="Blank.html" STYLE="positon:absolute; LEFT:0; WIDTH:405; HEIGHT:105; TOP:0; BORDER:none" 
                FRAMEBORDER=0>
            </IFRAME>
        </DIV>
        <BUTTON id=cmdSave class=DefBUTTON title="Save the value for the selected option" 
            style="LEFT:550; TOP:390; WIDTH:75"
            tabIndex=7>Save
        </BUTTON>
        <BUTTON id=cmdClose class=DefBUTTON title="" 
            style="LEFT:635; TOP:390; WIDTH:75"
            tabIndex=8>Close
        </BUTTON>
 
    </DIV>
</BODY>

<FORM NAME=Form METHOD="Post" STYLE="VISIBILITY: hidden" ACTION="Main.ASP" ID=Form>
    <%
    Call CommonFormFields()
    WriteFormField "Description", ReqForm("Description")
    WriteFormField "SettingValue", ReqForm("SettingValue")
    WriteFormField "SettingName", ReqForm("SettingName")
    WriteFormField "InputMask", ""

    strTmp = "<INPUT TYPE=""hidden"" Name="
    Response.Write strTmp & """FormAction"" value=""" & ReqForm("FormAction") & """>"
    %>
</FORM>

</HTML>
<%
gadoCon.Close
Set gadoCon = Nothing
%>

<%'----------------------------------------------------------------------------
'  Server side include files:
'----------------------------------------------------------------------------%>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncNavigateControls.asp"-->