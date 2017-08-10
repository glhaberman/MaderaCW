<SCRIPT LANGUAGE=vbscript>
<!--

'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: CrtAppProcessing.asp                                            '
'  Purpose: The report criteria selection screen for the Application        '
'           Processing report (medicaid).                                   '
'==========================================================================='
'Hides all criteria
Sub Hide_Criteria()
	lblDirector.style.visibility="hidden"
	cboDirector.style.visibility="hidden"
	cboDirCopy.style.visibility="hidden"
	lblOffice.style.visibility="hidden"
	cboOffice.style.visibility="hidden"
	cboOffCopy.style.visibility="hidden"
	lblProgramManager.style.visibility="hidden"
	cboProgramManager.style.visibility="hidden"
	cboMgrCopy.style.visibility="hidden"
	lblSupervisor.style.visibility="hidden"
	cboSupervisor.style.visibility="hidden"
	cboSupCopy.style.visibility="hidden"
	lblReviewer.style.visibility="hidden"
	cboReviewer.style.visibility="hidden"
	cboRvwCopy.style.visibility="hidden"
	lblAuthBy.style.visibility="hidden"
	cboAuthBy.style.visibility="hidden"
	cboAuthCopy.style.visibility="hidden"
	lblWorker.style.visibility="hidden"
	cboWorker.style.visibility="hidden"
	cboWkrCopy.style.visibility="hidden"
	
	lblProgram.style.visibility="hidden"
	cboProgram.style.visibility="hidden"
	lblCalWORKS.style.visibility="hidden"
	cboCalWORKS.style.visibility="hidden"
	lblMedical.style.visibility="hidden"
	cboMedical.style.visibility="hidden"
	lblEligElement.style.top = 95
	cboEligElement.style.top = 110
	lblEligElement.innerText="Eligibility Element"
	lblEligElement.style.visibility="hidden"
	cboEligElement.style.visibility="hidden"
	
	lblCaseAction.style.visibility="hidden"
	cboCaseAction.style.visibility="hidden"
	lblErrorDiscovery.style.visibility="hidden"
	cboErrorDiscovery.style.visibility="hidden"
	lblCompliance.style.visibility="hidden"
	cboCompliance.style.visibility="hidden"
	lbldivSubmitted.style.visibility="hidden"
	divSubmitted.style.visibility="hidden"
	lblCaseNumber.style.visibility="hidden"
	txtCaseNumber.style.visibility="hidden"
	lblDetail.style.visibility="hidden"
	lblMinDays.style.visibility="hidden"
	txtMinDays.style.visibility="hidden"
	lblHousehold.style.visibility="hidden"
	cboHousehold.style.visibility="hidden"
	lblPartHours.style.visibility="hidden"
	cboPartHours.style.visibility="hidden"
	lblResponse.style.visibility="hidden"
	cboResponse.style.visibility="hidden"

	divReviewType.style.visibility="hidden"
End Sub

'Display the appropriate Criteria for the selected Report
Sub Display_Criteria()
	Dim strReport
	Dim intI
	Dim intLength
	
	strReport = Parse(lstReports.value, ":", 1)
	Call Hide_Criteria()
	If optReportMode2.checked  Then'
		mblnUseAuthBy = True
	Else 
		mblnUseAuthBy = False
	End If
	
	Select Case strReport
        Case 26 'Case Accuracy Summary
			Call SetMgrSupWkrAbyRT(mblnUseAuthby)
			lblCaseAction.style.visibility="visible"
			cboCaseAction.style.visibility="visible"
			lblProgram.style.visibility="visible"
			cboProgram.style.visibility="visible"
			If lstProgram.selectedIndex = 0 Then
				lstProgram.selectedIndex = 1
				Call cboProgram_onchange
			Else
				Call cboProgram_onchange
			End If
		Case 27 'Case Review Detail
            Call SetMgrSupWkrAbyRT(mblnUseAuthby)
			lblCaseAction.style.visibility="visible"
			cboCaseAction.style.visibility="visible"
			lblCaseNumber.style.visibility="visible"
			txtCaseNumber.style.visibility="visible"
			lblDetail.innerHTML="Show Element Detail" & "<INPUT type=checkbox id=chkDetail style=""LEFT:160; WIDTH:20; HEIGHT:20; TOP:-2"" tabIndex=1>"
			lblDetail.style.top=110
			lblDetail.style.visibility="visible"
			Call ListReviewTyps(-1)
        Case 28 'Causal Factor Summary
            Call SetMgrSupWkrAbyRT(mblnUseAuthby)
            lblCaseAction.style.visibility="visible"
			cboCaseAction.style.visibility="visible"
			lblProgram.style.visibility="visible"
			cboProgram.style.visibility="visible"
			If lstProgram.selectedIndex = 0 Then
				lstProgram.selectedIndex = 1
			End If
			Call cboProgram_onchange
			lblEligElement.style.visibility="visible"
			cboEligElement.style.visibility="visible"
        Case 29 'Eligibility Element Detail
            Call SetMgrSupWkrAbyRT(mblnUseAuthby)
            lblCaseAction.style.visibility="visible"
			cboCaseAction.style.visibility="visible"
			lblProgram.style.visibility="visible"
			cboProgram.style.visibility="visible"
			If lstProgram.selectedIndex = 0 Then
				lstProgram.selectedIndex = 1
			End If
			Call cboProgram_onchange
			lblEligElement.innerText="Eligibility Element(required)"
			lblEligElement.style.visibility="visible"
			cboEligElement.style.visibility="visible"
        Case 30 'Eligibility Element Summary
            Call SetMgrSupWkrAbyRT(mblnUseAuthby)
			lblCaseAction.style.visibility="visible"
			cboCaseAction.style.visibility="visible"
			lblProgram.style.visibility="visible"
			cboProgram.style.visibility="visible"
			If lstProgram.selectedIndex = 0 Then
				lstProgram.selectedIndex = 1
			End If
			Call cboProgram_onchange
			lblEligElement.style.visibility="visible"
			cboEligElement.style.visibility="visible"
        Case 31 'Payment Error Summary
            Call SetMgrSupWkrAbyRT(mblnUseAuthby)
			lblCaseAction.style.visibility="visible"
			cboCaseAction.style.visibility="visible"
			lblErrorDiscovery.style.visibility="visible"
			cboErrorDiscovery.style.visibility="visible"
			lblCalWORKS.style.visibility="visible"
			cboCalWORKS.style.visibility="visible"
			If lstCalWORKS.selectedIndex = 0 Then
				lstCalWORKS.selectedIndex = 1
			End If
			Call cboCalWorks_onChange
        Case 32 'Reviewer Case Staus
            Call SetMgrRvw()
			divReviewType.style.visibility="visible"
			Call ListReviewTyps(-1)
        Case 33 'Unsubmitted Reviews
            Call SetMgrRvw()
        Case 34 'Response Due
            Call SetMgrSupWkrAbyRT(mblnUseAuthby)
			lbldivSubmitted.style.visibility="visible"
			divSubmitted.style.visibility="visible"
			lblResponse.style.visibility="visible"
			cboResponse.style.visibility="visible"
			Call ListReviewTyps(-1)
        Case 35 'Reviewer Case Count
            Call SetMgrRvw()
        Case 45 'Participation Hours Summary
            Call SetMgrSupWkrAbyRT(mblnUseAuthby)
			lblCaseAction.style.visibility="visible"
			cboCaseAction.style.visibility="visible"
			lblProgram.style.visibility="visible"
			cboProgram.style.visibility="visible"
			If lstProgram.selectedIndex = 0 Then
				lstProgram.selectedIndex = 1
			End If
			Call cboProgram_onchange
			lblHousehold.style.visibility="visible"
			cboHousehold.style.visibility="visible"
			lblPartHours.style.visibility="visible"
			cboPartHours.style.visibility="visible"
			lblDetail.innerHTML="Show Detail" & "<INPUT type=checkbox id=chkDetail style=""LEFT:160; WIDTH:20; HEIGHT:20; TOP:-2"" tabIndex=1>"
			lblDetail.style.left = 235
			lblDetail.style.top = 200
			lblDetail.style.visibility="visible"
        Case 47 'Supervisor Compliance Report
			Call SetMgrRvw()
			lblReviewer.innerText="Supervisor"
			lblCompliance.style.visibility="visible"
			cboCompliance.style.visibility="visible"
		Case 49 'Application Processing Report
            Call SetMgrSupWkrAbyRT(mblnUseAuthby)
			lblMedical.style.visibility="visible"
			cboMedical.style.visibility="visible"
			lstMedical.selectedIndex = 1
			Call cboMedical_onChange
			txtMedical.value = lstMedical.options.item(lstMedical.selectedIndex).text
			lblMinDays.style.visibility="visible"
			txtMinDays.style.visibility="visible"
        Case 50 'Timely Processing Report
            Call SetMgrSupWkrAbyRT(mblnUseAuthby)
			lblCaseAction.style.visibility="visible"
			cboCaseAction.style.visibility="visible"
			lblMedical.style.visibility="visible"
			cboMedical.style.visibility="visible"
			lstMedical.selectedIndex = 1
			Call cboMedical_onChange
        Case 51 'Significant Error Summary
            Call SetMgrSupWkrAbyRT(mblnUseAuthby)
			lblCaseAction.style.visibility="visible"
			cboCaseAction.style.visibility="visible"
			lblErrorDiscovery.style.visibility="visible"
			cboErrorDiscovery.style.visibility="visible"
			lblProgram.style.visibility="visible"
			cboProgram.style.visibility="visible"
			If lstProgram.selectedIndex = 0 Then
				lstProgram.selectedIndex = 1
			End If
			Call cboProgram_onchange
        Case 52 'Benefit Accuracy Summary
			Call SetMgrSupWkrAbyRT(mblnUseAuthby)
			lblCaseAction.style.visibility="visible"
			cboCaseAction.style.visibility="visible"
			lblMedical.style.visibility="visible"
			cboMedical.style.visibility="visible"
			lstMedical.selectedIndex = 1
			Call cboMedical_onChange
		Case 53 'Procedural Error Summary
			Call SetMgrSupWkrAbyRT(mblnUseAuthby)
			lblCaseAction.style.visibility="visible"
			cboCaseAction.style.visibility="visible"
			lblMedical.style.visibility="visible"
			cboMedical.style.visibility="visible"
			lstMedical.selectedIndex = 1
			Call cboMedical_onChange
		Case 55 'Evaluation Accuracy Summary
			Call SetMgrRvw()
			lblProgram.style.visibility="visible"
			cboProgram.style.visibility="visible"
			If lstProgram.selectedIndex = 0 Then
				lstProgram.selectedIndex = 1
			End If
			Call cboProgram_onchange
			lblEligElement.style.top = 50
			cboEligElement.style.top = 65
			lblEligElement.style.visibility="visible"
			cboEligElement.style.visibility="visible"
		Case 56 'Evaluation Eligibility Element Summary
			Call SetMgrRvw()
			lblProgram.style.visibility="visible"
			cboProgram.style.visibility="visible"
			If lstProgram.selectedIndex = 0 Then
				lstProgram.selectedIndex = 1
			End If
			Call cboProgram_onchange
		Case 57 'Case Accuracy Detail
			Call SetMgrSupWkrAbyRT(mblnUseAuthby)
			lblCaseAction.style.visibility="visible"
			cboCaseAction.style.visibility="visible"
			lblProgram.style.visibility="visible"
			cboProgram.style.visibility="visible"
			lstProgram.selectedIndex = 1
			Call cboProgram_onchange
		Case 58 'Payment Error Detail
			Call SetMgrSupWkrAbyRT(mblnUseAuthby)
			lblCaseAction.style.visibility="visible"
			cboCaseAction.style.visibility="visible"
			lblErrorDiscovery.style.visibility="visible"
			cboErrorDiscovery.style.visibility="visible"
			lblCalWORKS.style.visibility="visible"
			cboCalWORKS.style.visibility="visible"
			lstCalWORKS.selectedIndex = 1
			Call cboCalWorks_onChange
		Case 60 'Supervisor Compliance By Supervisor
			Call SetMgrRvw()
			lblReviewer.innerText="Supervisor"
			lblCompliance.style.visibility="visible"
			cboCompliance.style.visibility="visible"
			lblDetail.innerHTML="Show Submitted Reviews Only" & "<INPUT type=checkbox id=chkDetail style=""LEFT:170; WIDTH:20; HEIGHT:20; TOP:-2"" tabIndex=1>"
			lblDetail.style.left = 235
			lblDetail.style.top=50
			lblDetail.style.visibility="visible"
			
    End Select
	
End SUB

Sub SetMgrSupWkrAbyRT(mblnUseAuthby)
	lblDirector.style.visibility="visible"
	cboDirCopy.style.visibility="visible"
	lblOffice.style.visibility="visible"
	cboOffCopy.style.visibility="visible"
	lblProgramManager.style.visibility="visible"
	cboMgrCopy.style.visibility="visible"
	lblSupervisor.style.visibility="visible"
	cboSupCopy.style.visibility="visible"
	If mblnUseAuthby Then 
		lblAuthBy.style.visibility="visible"
		cboAuthCopy.style.visibility="visible"
	Else
		lblWorker.style.visibility="visible"
		cboWkrCopy.style.visibility="visible"
	End If
	divReviewType.style.visibility="visible"
End Sub

Sub SetMgrRvw()
	lblDirector.style.visibility="visible"
	cboDirCopy.style.visibility="visible"
	lblOffice.style.visibility="visible"
	cboOffCopy.style.visibility="visible"
	lblProgramManager.style.visibility="visible"
	cboMgrCopy.style.visibility="visible"
	lblReviewer.style.visibility="visible"
	cboRvwCopy.style.visibility="visible"
End Sub

'Hides all criteria
Sub cmdClearCriteria_onclick()
	Dim intI
    Dim oCtl
	
	cboDirector.value = 0
	cboOffice.value = 0
	cboProgramManager.value = 0
	cboSupervisor.value = 0
	cboReviewer.value = 0
	cboAuthBy.value = 0
	cboWorker.value = 0
	lstDirCopy.value = 0
	lstOffCopy.value = 0
	lstMgrCopy.value = 0
	lstSupCopy.value = 0
	lstRvwCopy.value = 0
	lstAuthCopy.value = 0
	lstWkrCopy.value = 0
	
	txtDirCopy.value = ""
	txtOffCopy.value = ""
	txtMgrCopy.value = ""
	txtSupCopy.value = ""
	txtRvwCopy.value = ""
	txtAuthCopy.value = ""
	txtWkrCopy.value = ""
	
	lstCaseAction.value = 0
	lstErrorDiscovery.value = 0
	lstCompliance.value = 0
	lstProgram.value = 0
	lstEligElement.value = 0
	
	txtCaseAction.value = ""
	txtErrorDiscovery.value = ""
	txtCompliance.value = ""
	txtProgram.value = ""
	txtEligElement.value = ""
	
    Set oCtl = Nothing
	If chkSubmitted.checked Then
		chkSubmitted.checked = False
	End If
	IF chkUnsubmitted.checked Then
		chkUnsubmitted.checked = False
	End If
	If chkDetail.checked Then
		chkDetail.checked = False
	End If
	txtCaseNumber.value = 0
	txtMinDays.value = 0
	lstHousehold.value = 0
	lstPartHours.value = 0
	lstResponse.value = 0
	txtStartDate.value = ""
	txtEndDate.value = ""

	For intI = 0 To lblReviewCount.innerText - 1
		If document.all("chkReviewType" & intI).Checked Then
			Call lblChkReviewType_onclick(intI)
		End If
	Next
	
	For intI = 0 To lblReviewClassCount.innerText - 1
		If document.all("chkReviewClass" & intI).Checked Then	
			Call lblChkReviewClass_onclick(intI)
		End If
	Next
End Sub
-->
</Script>