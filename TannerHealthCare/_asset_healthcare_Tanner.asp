<%@ EnableSessionState=False Language=VBScript %>
<% Option Explicit %>

<!--#INCLUDE FILE="../common/mc_all_cache.asp" -->
<!--#INCLUDE FILE="../common/mc_tabcontrol_Tanner.asp" -->

<%
Dim keyvalue,keyvalues,asclass,asdefault,lastaction,errorfield,errortabinfo,returnmessage,returnclass,newrecord,duprecord,mcmode,findtabtext,treefiltervalue,norecord,firstload,curtab
Dim ServerDate,addontheflymode,addontheflymodule,currentmodule

' Custom Variables for Assets
' -----------------------------------------------------------------------------
Dim AssetUpdate, IsIDChange, IsNameChange, IsTypeChange, IsClassificationChange, IsLocation, IsIconChange, IsParentChange
Dim OldFieldValue, EventIDs
Dim PMCycleStartBy, PMCycleStartByDesc

EventIDs = ""
IsIDChange = False
IsNameChange = False
IsTypeChange = False
IsClassificationChange = False
IsLocation = False
IsIconChange = False
IsParentChange = False
AssetUpdate = False

' -----------------------------------------------------------------------------

init
checkforjustdata
checkforaction
checkforsubmit
displayhtml
' ------------------------------------ Functions ---------------------------------------
Sub Init()

	If ErrorHandler Then
		On Error Resume Next
	End If

	GlobalInit

	If newrecord Then
		asclass = Trim(Request.QueryString("asclass"))
		asdefault = Trim(Request.QueryString("asdefault"))
	Else
		asclass = "1"
		asdefault = ""
	End If

	' This is NOT the SQL Server Date - rather the Web Server Date! (for now)
	ServerDate = DateTimeNullCheck(Now())

End Sub

Sub displayhtml

	' ModuleSpecific
	Dim windowonloadjs, headercmds
	windowonloadjs = ""
	headercmds = _
	"<script language=""javascript"" src=""../../javascript/normal/mc_tabcontrol.js""></script>" & _
	"<script language=""javascript"" src=""../../javascript/normal/mc_autocomplete.js""></script>" & _
	"<script type=""text/javascript"" src=""../../javascript/jquery/jquery-ui-1.11.2.min.js""></script>" & _
	"<link rel=""stylesheet"" type=""text/css"" href=""../../css/mc_tabstyle.css"">"
	'windowonloadjs = windowonloadjs + "top.fraTopic.document.body.style.backgroundColor = '#808080';" + nl
	'windowonloadjs = windowonloadjs + "top.fraTopic.document.body.style.backgroundImage = '';" + nl
	windowonloadjs = windowonloadjs + "top.fraTabbar.document.images.maxmin.style.marginRight='0';" + nl
	windowonloadjs = windowonloadjs + "top.fraTabbar.document.images.closemod.style.display='';" + nl
	windowonloadjs = windowonloadjs + "pagetabs_current = self.document.getElementById('pagetabs_31');" + nl
	windowonloadjs = windowonloadjs + "top.endprocess();" + nl
	windowonloadjs = windowonloadjs + "sched_init();" + nl
	windowonloadjs = windowonloadjs + "top.navcurrent(null,null,true);" + nl

	' Asset AutoComplete Fields
	windowonloadjs = windowonloadjs & "try {"
	windowonloadjs = windowonloadjs & "if (top.acom == true) {"
	windowonloadjs = windowonloadjs & "var txtParent_AC = new actb('AS','AS',document.mcform.txtParent);"
	windowonloadjs = windowonloadjs & "var txtType_AC = new actb('LOOKUP_ASSETTYPE','AS',document.mcform.txtType);"
	windowonloadjs = windowonloadjs & "var txtClassification_AC = new actb('CL','AS',document.mcform.txtClassification);"
	windowonloadjs = windowonloadjs & "var txtSystem_AC = new actb('LOOKUP_ASSETSYSTEM','AS',document.mcform.txtSystem);"
	windowonloadjs = windowonloadjs & "var txtOperator_AC = new actb('LA','AS',document.mcform.txtOperator);"
	windowonloadjs = windowonloadjs & "var txtDepartment_AC = new actb('DP','AS',document.mcform.txtDepartment);"
	windowonloadjs = windowonloadjs & "var txtTenant_AC = new actb('TN','AS',document.mcform.txtTenant);"
	windowonloadjs = windowonloadjs & "var txtAccount_AC = new actb('AC','AS',document.mcform.txtAccount);"
	windowonloadjs = windowonloadjs & "var txtRepairCenter_AC = new actb('RC','AS',document.mcform.txtRepairCenter);"
	windowonloadjs = windowonloadjs & "var txtShop_AC = new actb('SH','AS',document.mcform.txtShop);"
	windowonloadjs = windowonloadjs & "var txtPriority_AC = new actb('LOOKUP_WOPRIORITY','AS',document.mcform.txtPriority);"
	windowonloadjs = windowonloadjs & "var txtVendor_AC = new actb('CM_VENDOR','AS',document.mcform.txtVendor);"
	windowonloadjs = windowonloadjs & "var txtManufacturer_AC = new actb('CM_MANUFACTURER','AS',document.mcform.txtManufacturer);"
	windowonloadjs = windowonloadjs & "var txtContact_AC = new actb('LA_CONTACT','AS',document.mcform.txtContact);"
	windowonloadjs = windowonloadjs & "var txtOperator_AC = new actb('LA_CONTACT','AS',document.mcform.txtOperator);"
	windowonloadjs = windowonloadjs & "var txtTechnology_AC = new actb('LOOKUP_PDTECHNOLOGY','AS',document.mcform.txtTechnology);"
	windowonloadjs = windowonloadjs & "var txtMeter1Units_AC = new actb('LOOKUP_METERUNITS','AS',document.mcform.txtMeter1Units);"
	windowonloadjs = windowonloadjs & "var txtMeter2Units_AC = new actb('LOOKUP_METERUNITS','AS',document.mcform.txtMeter2Units);"
	windowonloadjs = windowonloadjs & "var txtPurchaseType_AC = new actb('LOOKUP_PURCHASETYPE','AS',document.mcform.txtPurchaseType);"
	windowonloadjs = windowonloadjs & "var txtZone_AC = new actb('ZN','AS',document.mcform.txtZone);"
	windowonloadjs = windowonloadjs & "var txtConstructionCode_AC = new actb('LOOKUP_CONSTRUCTIONCODE','AS',document.mcform.txtConstructionCode);"
	windowonloadjs = windowonloadjs & "var txtISOProtection_AC = new actb('LOOKUP_ISOPROTECTION','AS',document.mcform.txtISOProtection);"
	windowonloadjs = windowonloadjs & "var txtAutoSprinkler_AC = new actb('LOOKUP_AUTOSPRINKLER','AS',document.mcform.txtAutoSprinkler);"
	windowonloadjs = windowonloadjs & "var txtSmokeAlarm_AC = new actb('LOOKUP_SMOKEALARM','AS',document.mcform.txtSmokeAlarm);"
	windowonloadjs = windowonloadjs & "var txtHeatAlarm_AC = new actb('LOOKUP_HEATALARM','AS',document.mcform.txtHeatAlarm);"
	windowonloadjs = windowonloadjs & "var txtResponsibilityRepair_AC = new actb('CM','AS',document.mcform.txtResponsibilityRepair);"
	windowonloadjs = windowonloadjs & "var txtResponsibilityPM_AC = new actb('CM','AS',document.mcform.txtResponsibilityPM);"
	windowonloadjs = windowonloadjs & "var txtResponsibilitySafety_AC = new actb('CM','AS',document.mcform.txtResponsibilitySafety);"
	windowonloadjs = windowonloadjs & "var txtServiceRepair_AC = new actb('CM','AS',document.mcform.txtServiceRepair);"
	windowonloadjs = windowonloadjs & "var txtServicePM_AC = new actb('CM','AS',document.mcform.txtServicePM);"
	windowonloadjs = windowonloadjs & "var txtPMCycleStartBy_AC = new actb('LOOKUP_PMCYCLESTARTBY','AS',document.mcform.txtPMCycleStartBy);"
	windowonloadjs = windowonloadjs & "var txtModelLine_AC = new actb('LOOKUP_MODELLINE','AS',document.mcform.txtModelLine);"
	windowonloadjs = windowonloadjs & "var txtModelSeries_AC = new actb('LOOKUP_MODELSERIES','AS',document.mcform.txtModelSeries);"
	windowonloadjs = windowonloadjs & "var txtSystemPlatform_AC = new actb('LOOKUP_SYSTEMPLATFORM','AS',document.mcform.txtSystemPlatform);"
	windowonloadjs = windowonloadjs & "var txtClassIndustry_AC = new actb('LOOKUP_CLASSINDUSTRYHEALTHCARE','AS',document.mcform.txtClassIndustry);"
	windowonloadjs = windowonloadjs & "var txtAssessedBy_AC = new actb('LA_CONTACT','AS',document.mcform.txtAssessedBy);"
	windowonloadjs = windowonloadjs & "var txtRiskAssessmentGroup_AC = new actb('LOOKUP_RISKASSESSMENTGROUP','AS',document.mcform.txtRiskAssessmentGroup);"
	windowonloadjs = windowonloadjs & "var txtMaintainableTool_AC = new actb('TL','AS',document.mcform.txtMaintainableTool);"
	windowonloadjs = windowonloadjs & "var txtRotatingPart_AC = new actb('IN','AS',document.mcform.txtRotatingPart);"
	windowonloadjs = windowonloadjs & "}"
	windowonloadjs = windowonloadjs & "} catch(e) {};"

	Call domctop("",True,headercmds,True)

	FlushIt
	%>
	<div id="mcpage" name="asset" moduleid="AS" allownewrecords="Y">

		<% domcstartform %>
		<table border="0" cellpadding="0" width="100%" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#CCCCCC" cellspacing="0" height="98%">
		  <tr>
			<td class="mcbgcolor" valign="top">
			<div id="wcbox" class="wcboxstyle">                  		    
			<table id="mcpagetable" border="0" cellspacing="0" width="100%" cellpadding="2" height="100%" style="background-position: 95% 89%;background-repeat: no-repeat; background-attachment: fixed; background-image:;);">
			<tr>
				<td width="10" valign="top"></td>
				  <td valign="top">

					<!-- Splash Tab -->
					<div id="splash" name="tab71" STYLE="display:none;" mcfocus="NONE" watermark="images/as_watermark.jpg">
					<!--#INCLUDE FILE="asset_splash.htm" -->
					</div>

					<!-- Tab #1  -->
					<div id="as_details" name="tab71" dataloaded="Y" STYLE="display: none;" showheader="N" multipage="Y">
					<!--#INCLUDE FILE="asset_details60.htm" -->
					</div>

					<!-- Tab #2  -->
					<div id="as_risk" name="tab72" dataloaded="N" STYLE="display: none;" showheader="N">
					<!--#INCLUDE FILE="_asset_risk_Tanner.htm" --> 
					</div>					
					<!-- Tab #3  -->

					<div id="as_meter" name="tab73" dataloaded="N" STYLE="display: none;" showheader="N" watermark="images/meter_watermark.jpg">
					<!--#INCLUDE FILE="asset_meter.htm" -->
					</div>
					<% FlushIt %>

					<!-- Tab #3  -->
					<div id="as_specifications" name="tab74" dataloaded="N" STYLE="display:none;" showheader="N" mcfocus="NONE">
					<!--#INCLUDE FILE="asset_specifications.htm" -->
					</div>

					<!-- Tab #4  -->
					<div id="as_procedures" name="tab75" dataloaded="N" STYLE="display: none;" showheader="N" mcfocus="NONE" noaccessinedit="Y" watermark="images/pr_watermark.jpg">
					<!--#INCLUDE FILE="asset_procedures60.htm" -->
					</div>

					<!-- Tab #5  -->
					<div id="as_attach" name="tab76" dataloaded="N" STYLE="display: none;" showheader="N" mcfocus="NONE" showscroll="Y" noaccessinedit="Y" watermark="images/attach_watermark2.jpg">
					<!--#INCLUDE FILE="asset_attach.htm" -->
					</div>

					<!-- Tab #8  -->
					<div id="as_schedule" name="tabsched" dataloaded="N" STYLE="display: none;" showheader="N" mcfocus="NONE" showscroll="N" noaccessinedit="Y">
					<!--#INCLUDE FILE="asset_sched.htm" -->
					</div>

					<!-- Tab #6  -->
					<div id="as_history" name="tab77" dataloaded="N" STYLE="display: none;" showheader="N" showscroll="N" mcfocus="NONE" noaccessinedit="Y">
					<!--#INCLUDE FILE="asset_history.htm" -->
					</div>

					<!-- Tab #7 // CB2: 2.5  -->
					<div id="as_lease" name="tab78" dataloaded="N" STYLE="display: none;" showheader="N" showscroll="Y" mcfocus="NONE" noaccessinedit="Y">
					    <iframe allowtransparency="true" style="height:100%; width:100%;" name="leaseFrame" id="leaseFrame" frameborder="0" scrolling="no" src="../Leases/leases.asp?apk=-1"></iframe>
					</div>
					<% FlushIt %>
				  </td>
				  <td width="10" valign="top"></td>
			  </tr>
              
              </table>
              
              </div>

              <table id="footertable" style="visibility:hidden; margin-top:4px;" border="0" cellspacing="0" width="100%" cellpadding="0">                        
		      <tr>
		        <td width="10" valign="top"></td>
		        <td valign="bottom">
                  <div style="height:1px; background-color:#CCCCCC;"></div>
		        </td>
		        <td width="10" valign="top"></td>
		      </tr>
		      <tr>
		        <td width="10" valign="top"></td>
		        <td>
		        <table border="0" cellpadding="0" cellspacing="0" width="100%">
		          <tr>
		            <td width="35%" valign="top">

					<% Call OutputBottomTabs(3) %>

		            </td>
		            <td width="65%" valign="top" align="right" style="padding-top:3px;white-space:nowrap;">
                    <% Call ButtonFactory("G_ASSET","Asset.AssetPK","ASProfile") %>                        						
		            </td>
		          </tr>
		        </table>
		        </td>
		        <td width="10" valign="top"></td>
		      </tr>
		    </table>
		    </td>
		  </tr>
		</table>

		<script language="javascript">
			document.body.style.backgroundColor = '#808080';
			document.body.style.backgroundImage = '';
		</script>

		<div id="hiddenfields">
		</div>

		<div id="temphtml" style="display:none;">
		</div>

		</form>

	</div>

	<%
	domcbottom windowonloadjs
	'filldata false
	Response.Write jsstart
	Response.Write formjs
	Call Setup(False)
	Response.Write jsend
	%>
	</body>
	</html>
	<%

End Sub

Sub checkforsubmit()

	If ErrorHandler Then
		On Error Resume Next
	End If

	If Request.Form("asubmit") = "" Then
		Exit Sub
	End If

	Dim dok,derror,theaction,newjs,playsound

	theaction = LCase(Trim(Request.Form("asubmit")))
	lastaction = Trim(Request.Form("lastaction"))
	playsound = True

	Select Case Trim(UCase(theaction))

		Case "SAVE"

			Dim db,d,thing,themode,thechild,theindex,ra,pkandrv,rpk,rv,theprefix,thesuffix
			Dim spcwhere,amcwhere,apcwhere,aocwhere,alcwhere,arcwhere,atcwhere,nocwhere
			Dim spnew,amnew,apnew,aonew,alnew,arnew,atnew,nonew
			Dim spdwhere,amdwhere,apdwhere,aodwhere,aldwhere,ardwhere,atdwhere,nodwhere

			' <SERVER-SIDE VALIDATION STUFF HERE>

			aok = True
			errorfield = ""
			returnmessage = ""
			returnclass = "standardmessage"

			Call validate_master()

			' <LOOP THROUGH AND FORMAT FORM VARIABLES>

			Set d = CreateObject("Scripting.Dictionary")

			For each thing in Request.Form
				If Mid(thing,1,4) = "mcf_" Then

					ra = split(thing,"_")

					themode = UCase(ra(1))
					thechild = ra(2)
					theindex = ra(3)
					theprefix = UCase(Left(thechild,2))
					thesuffix = Mid(thechild,3,1)

					pkandrv = split(Trim(Request.Form( thing ).Item),"$")
					rpk = pkandrv(0)
					rv = pkandrv(1)

					'Response.Write thing + " -- " + themode + "--" + thechild + "--" + theindex + "<br>"
					Execute "Call validate_" & thechild & "(thing,theindex)"

					Select Case themode
						Case "C"
							Execute(theprefix & "cwhere = " & theprefix & "cwhere & """ & rpk & ",""")
							d.Add theprefix & rpk, theindex & "$" & rv
						Case "N"
							Execute(theprefix & "new = " & theprefix & "new & """ & thesuffix & theindex & "#!#""")
						Case "D"
							Execute(theprefix & "dwhere = " & theprefix & "dwhere & """ & rpk & ",""")
					End Select

				End If
			Next

			Set db = New ADOHelper

			If db.OpenClientConnection Then
				If db.OpenTransaction Then

					' <SAVE THE MAIN RECORD>

					Call db_master(db)

					' <SAVE THE CHILDREN RECORDS>

					Call db_child(db,d,"SP","assetspecification",spcwhere,spnew,spdwhere)
					Call db_child(db,d,"AM","assetpart",amcwhere,amnew,amdwhere)
					Call db_child(db,d,"AP","pmasset",apcwhere,apnew,apdwhere)
					Call db_child(db,d,"AO","assetcontract",aocwhere,aonew,aodwhere)
					Call db_child(db,d,"AL","assetlabor",alcwhere,alnew,aldwhere)
					Call db_child(db,d,"AR","assetrequester",arcwhere,arnew,ardwhere)
					Call db_child(db,d,"AT","assetdocument",atcwhere,atnew,atdwhere)
					Call db_child(db,d,"NO","assetnote",nocwhere,nonew,nodwhere)

					db.CloseTransaction

				End If
			End If
			dok = db.dok
			derror = db.derror
			Set d = Nothing

			' <RETURN TO CLIENT HERE>

			If dok Then

				'Call DoNotify(EventIDs,keyvalue,db)

				If addontheflymode = "Y" Then
					newjs = newjs + "	top.addonthefly_recupdated = true;" + nl
					If newrecord Then
						newjs = newjs & "	top.addonthefly_recaddedpk = '" & CStr(keyvalue) & "';"
					End If
				End If

				newjs = newjs + "	top.mcmode = 'EDIT';" + nl
				newjs = newjs + "	top.dirtydata = false;" + nl
				newjs = newjs + "	top.requiredfields('OFF');" + nl

				If Not newrecord Then
					If (IsIDChange and Not IsLocation) or (IsNameChange) or (IsIconChange) or (IsTypeChange) Then
						newjs = newjs + "   if (top.fraAssets) {" + nl
						If IsLocation Then
							newjs = newjs + "	top.fraAssets.UpdateTreeLI('" & keyvalue & "','" & JSEncode(Request.Form("txtAssetName")) & "','" & Trim(Request.Form("txtIcon")) & "',true);" + nl
						Else
							newjs = newjs + "	top.fraAssets.UpdateTreeLI('" & keyvalue & "','" & JSEncode(Request.Form("txtAssetName")) & " (" & JSEncode(Request.Form("txtAsset")) & ")','" & Trim(Request.Form("txtIcon")) & "',false);" + nl
						End If
						If (IsIDChange and Not IsLocation) or (IsNameChange) Then
							newjs = newjs + "	top.fraAssets.changetitles(top.fraAssets.document);" + nl
						End If
						newjs = newjs + "	}" + nl
					End If
				End If

				If IsParentChange Then
					' We need to set the correct action button here
					' because the next line starts a new process before
					' the save process actually gets done.
					newjs = newjs + "	top.showactions('aasrecbar');" + nl
					newjs = newjs + "	if (top.fraAssets) {top.fraAssets.movecomplete('save',null," & keyvalue & "," & Trim(Request.Form("txtParentPK").Item) & ");}" + nl
				End If

				returnmessage = "The Asset has been Saved"
				donocacheheader
				CommitSession
				dogenericheader
				Response.Write newjs

				If Not writerecorddata(db) Then
					db.CloseClientConnection
					Set db = Nothing
					Call dogenericendaction(True)
					dogenericfooter
					Response.End
				End If

				Call SetupFields("EDIT",True)
				Call EnableDisableFields("EDIT")

				' AUTO-REFRESH EXPLORER BY UN-COMMENTING LINES BELOW
				'======================================================================================================================================
				'If RefreshExplorerIsTurnedOn(db,currentmodule,newrecord) Then
				'	Response.Write("top.refreshcurrentexplorer(true);")+nl
				'End If
				'======================================================================================================================================
				db.CloseClientConnection

				Call dogenericendaction(True)
				Response.Write "	top.showtabinfo(top.currenttabinfo.id,false);" + nl
				If newrecord Then
				Response.Write "	top.navnewcomplete(top.currentmodule,'kv=" + CStr(keyvalue) + "');" + nl
				End If
				If db.warn Then
				Response.Write("   top.mcalert('info','Asset Message','The Asset has been saved successfully, but there were some warnings that you need to be aware of. The details of the warnings are described below.<br><br><u>Warning Details</u>:<br><br><span style=""color:royalblue;"">" & Replace(db.warntext,"'","\'") & "</span>','bg_okprint',700,370,'sounds/cancel.wav');") + nl
				Else
				Response.Write("	top.playsound('sounds/done.wav');")
				End If
				Response.Write "	top.dofocus();" + nl
				dogenericfooter
				Set db = Nothing
				Response.End

			Else
				If Not dok Then
					If db.isduplicate Then
						newjs = newjs + "   top.mcalert('warning','Asset Message','You have entered an ID that is already in use by another Asset or Location. Please change the ID and then click the SAVE button.','bg_okprint',700,240,'sounds/error.wav');" + nl
					Else
						newjs = newjs + "   top.mcalert('warning','Asset Message','There was a problem saving the Asset. The details of the problem are described below. You can try to SAVE the Asset again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.<br><br><u>Problem Details</u>:<br><br>" & Replace(derror,"'","\'") & "','bg_okprint',700,370,'sounds/error.wav');" + nl
					End If
					aok = False
				Else
					aok = True
				End If
				db.CloseClientConnection
				Set db = Nothing
				playsound = False
				returnmessage = ""
			End If

	Case Else

		returnmessage = "Error: Could not determine Action"
		returnclass = "errormessage"

	End Select

	Call DoResponse(lastaction,newjs,playsound,True,False,keyvalue)

End Sub

Sub checkforaction()

	If ErrorHandler Then
		On Error Resume Next
	End If

	Dim db,rs,derror,outarray,sql,dok,doaction,newjs,recinview,selrecs,criteria,ecriteria,showbox,actionwhere
	'@$CUSTOMISED
	Dim disassemble
	If Request("disassemble") = "1" Then
		disassemble = 1
	Else
		disassemble = 0
	End If

	doaction = UCase(Trim(Request("doaction")))
	If doaction = "" Then
		Exit Sub
	End If

	If Not Request.QueryString("doaction") = "" Then
		lastaction = Trim(Request.QueryString("lastaction"))
		recinview = Trim(UCase(Request.QueryString("recinview")))
		selrecs = Trim(Request.QueryString("sel"))
		criteria = Trim(Request.QueryString("findercrit"))
		ecriteria = Trim(Request.QueryString("finderecrit"))
		showbox = Trim(UCase(Request.QueryString("showbox")))
	Else
		lastaction = Trim(Request.Form("lastaction"))
		recinview = Trim(UCase(Request.Form("recinview")))
		selrecs = Trim(Request.Form("sel"))
		criteria = Trim(Request.Form("findercrit"))
		ecriteria = Trim(Request.Form("finderecrit"))
		showbox = Trim(UCase(Request.Form("showbox")))
	End If

	newjs = ""

	Set db = New ADOHelper

	Select Case doaction

		Case "DELETE"

				' <SERVER-SIDE VALIDATION STUFF HERE>

				aok = True
				errorfield = ""
				returnmessage = ""
				returnclass = "standardmessage"

				actionwhere = "WHERE (AssetPK IN (" & keyvalue & "))"

				If DemoMode() Then
					sql = "Select 'DemoTest' FROM Asset WITH (NOLOCK) " & _
					actionwhere & " AND Asset.DemoLaborPK = " & GetSession("UserPK")
					Call DemoNoActionMsg(db,sql)
				End If

				Call db.RunSP("MC_DeleteAsset",Array(Array("@AssetPK", adInteger, adParamInput, 4, keyvalue)),"")
				If db.isdeletecolref Then
					newjs = newjs + "   top.removemessage(); top.showactions('" & lastaction & "');top.loadmodeless('modules/common/mc_deleterecord.asp?t=ASSET&pk="&keyvalue&"',null,995,700,false,false);"+nl
					Call DoScript(newjs,False)
				Else
					Call dok_check(db,"Asset Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				End If

				'newjs = newjs + "		top.pendingmsg = 'The Asset has been Deleted';" + nl

				newjs = newjs + "		if (top.addontheflymode == true) {top.addonthefly_recupdated = true; top.addontheflyclose();} else {" + nl

				newjs = newjs + "		if (top.fraTocTop.currenttab.id == 'tab21') {actionUL = top.fraAssets.oActionNode.parentElement;" + nl
				newjs = newjs + "		if (top.fraAssets.oActionNode != null)" + nl
				newjs = newjs + "		{" + nl
				newjs = newjs + "			if (top.fraAssets.oActionNode.getAttribute('mckey') == '" & CStr(keyvalue) & "' && top.moduleforexplorer == 'AS') {top.fraAssets.oActionNode.removeNode(true); top.fraAssets.oActionNode=null;" + nl
				newjs = newjs + "		}" + nl
				newjs = newjs + "		if (actionUL && actionUL.getElementsByTagName(""li"").length == 0)" + nl
				newjs = newjs + "		{" + nl
				newjs = newjs + "			// make the parent node a bleaf because there are no more children" + nl
				newjs = newjs + "			actionUL.innerHTML = '';" + nl
				newjs = newjs + "			top.fraAssets.lihasnokids(actionUL);" + nl
				newjs = newjs + "		}}"

				newjs = newjs + "		top.fraAssets.oActionNode = null;" + nl

				newjs = newjs + "       } else {" + nl
				newjs = newjs + "       if (top.fraToc) {" + nl
				newjs = newjs + "               if (top.fraToc.currow != null) {" + nl
				newjs = newjs + "                   top.fraToc.currow.removeNode(true);" + nl
				newjs = newjs + "                   top.fraToc.currow = null;" + nl
				newjs = newjs + "               }" + nl
				newjs = newjs + "               top.fraToc.currow = null;" + nl
				newjs = newjs + "       }}" + nl

				If recinview = "YES" Then
					newjs = newjs + "		top.navadd('AS','kv=');" + nl
				End If

				newjs = newjs + "		}"

				'Call DoScript(newjs,False)

				returnmessage = "The Asset has been Deleted"

		Case "MOVE"

				' <SERVER-SIDE VALIDATION STUFF HERE>

				aok = True
				errorfield = ""
				returnmessage = ""
				returnclass = "standardmessage"

				actionwhere = "WHERE (AssetPK IN (" & Request.QueryString("FromPK").Item & "))"

				If DemoMode() Then
					sql = "Select 'DemoTest' FROM Asset WITH (NOLOCK) " & _
					actionwhere & " AND Asset.DemoLaborPK = " & GetSession("UserPK")
					Call DemoNoActionMsg(db,sql)
				End If

				'@$CUSTOMISED
				Call db.RunSP("MC_CheckCompoundAssetStockroomTransfer",Array(Array("@FromPK", adInteger, adParamInput, 4, Request.QueryString("FromPK").Item),Array("@ToPK", adInteger, adParamInput, 4, Request.QueryString("ToPK").Item),Array("@ReturnCode", adInteger, adParamOutPut, 4, "")),OutArray)
				'Response.Write OutArray(2)
				
				If disassemble = 0 AND OutArray(2) > 0 Then

					derror = "Unable to Move Assets that contain Child Assets.<br>Each Asset must be moved independently."
					'newjs = newjs + "   top.mcalert('warning','Asset Move Message','There was a problem moving the Asset. The details of the problem are described below.<br><br><u>Problem Details</u>:<br><br>" & Replace(derror,"'","\'") & "','bg_okprint',700,370,'sounds/error.wav');" + nl
					'newjs = newjs + "	top.fraAssets.oActionNode = null;" + nl
					returnmessage = ""

					newjs = newjs + "	if (top.addontheflymode == true) {top.addonthefly_recupdated = true;}" + nl	
					newjs = newjs + "	top.fraAssets.oActionNode = null;" + nl
					newjs = newjs + "   var theaction = top.startprocess();" + nl
					newjs = newjs + "   if (theaction == null || top.actioninprogress == false)" + nl
					newjs = newjs + "   { return; }" + nl
					newjs = newjs + "   if (top.mcquestion('Transfer Compound Asset?', 'Transferring an Asset that contains Child Assets into a Stockroom will result in the Asset being disassembled into individual Assets.<br><br> Are you sure you want to do this?') == true) {" + nl
					newjs = newjs + "      top.showmessage('Transferring...Please Wait', 'standardmessage');" + nl
					newjs = newjs + "      if (top.fraAssets) {" + nl

					newjs = newjs + "      if (top.fraAssets && top.fraAssets.eCurrentLI != null) {" + nl
					newjs = newjs + "         if ('INPUT' == top.fraAssets.eCurrentLI.children[1].tagName.toUpperCase()) {" + nl
					newjs = newjs + "            var whichDiv2 = top.fraAssets.eCurrentLI.children[3];" + nl
					newjs = newjs + "         }" + nl
					newjs = newjs + "         else {" + nl
					newjs = newjs + "            var whichDiv2 = top.fraAssets.eCurrentLI.children[2];" + nl
					newjs = newjs + "         }" + nl
					newjs = newjs + "      }" + nl
					newjs = newjs + "      else {" + nl
					newjs = newjs + "         var whichDiv2 = null;" + nl
					newjs = newjs + "      }" + nl

					newjs = newjs + "      var v = '';" + nl
					newjs = newjs + "      if (top.userpk) {" + nl
					newjs = newjs + "         v += '&txtRowVersionUserPK=' + top.userpk;" + nl
					newjs = newjs + "      }" + nl
					newjs = newjs + "      if (top.userinitials) {" + nl
					newjs = newjs + "         v += '&txtRowVersionInitials=' + top.userinitials;" + nl
					newjs = newjs + "      }" + nl

					newjs = newjs + "         top.fraAssets.oActionNode = whichDiv2.parentElement;" + nl
					newjs = newjs + "      }" + nl
					'newjs = newjs + "      alert(top.path + top.oAS.rooturl + top.thetop.oAS.actionpage + '?' + top.recordkey + '&recinview=yes&doaction=move&disassemble=1&lastaction=' + theaction + v + '&FromPK=" + Request.QueryString("FromPK").Item + "&ToPK=" + Request.QueryString("ToPK").Item + "');" + nl
					newjs = newjs + "      top.fraSubmit.location.replace(top.path + top.oAS.rooturl + top.oAS.actionpage + '?' + top.recordkey + '&recinview=yes&doaction=move&disassemble=1&lastaction=' + theaction + v + '&FromPK=" + Request.QueryString("FromPK").Item + "&ToPK=" + Request.QueryString("ToPK").Item + "');" + nl
					newjs = newjs + "   }" + nl
					newjs = newjs + "   else {" + nl
					newjs = newjs + "      top.showactions(theaction);" + nl
					newjs = newjs + "      top.endprocess();" + nl
					newjs = newjs + "   }" + nl

				ElseIf disassemble = 1 Then
					Call db.RunSP("MC_MoveCompoundAssetTree",Array(Array("@FromPK", adInteger, adParamInput, 4, Request.QueryString("FromPK").Item),Array("@ToPK", adInteger, adParamInput, 4, Request.QueryString("ToPK").Item),Array("@UserPK", adInteger, adParamInput, 4, Request.QueryString("txtRowVersionUserPK").Item),Array("@Initials", MC_ADVARCHAR, adParamInput, 5, Trim(Mid(Request.QueryString("txtRowVersionInitials").Item,1,5))),Array("@ErrorCode", adInteger, adParamOutPut, 4, "")),OutArray)
				Call dok_check(db,"Asset Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				Else
					Call db.RunSP("MC_MoveAssetTree",Array(Array("@FromPK", adInteger, adParamInput, 4, Request.QueryString("FromPK").Item),Array("@ToPK", adInteger, adParamInput, 4, Request.QueryString("ToPK").Item),Array("@UserPK", adInteger, adParamInput, 4, Request.QueryString("txtRowVersionUserPK").Item),Array("@Initials", MC_ADVARCHAR, adParamInput, 5, Trim(Mid(Request.QueryString("txtRowVersionInitials").Item,1,5))),Array("@ErrorCode", adInteger, adParamOutPut, 4, "")),OutArray)
					Call dok_check(db,"Asset Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				End If

				'Response.Write "FromPK: " & Request.QueryString("FromPK").Item & " ToPK: " & Request.QueryString("ToPK").Item & " Error: " & OutArray(4)
				'Response.End

				'@$CUSTOMISED
				If OutArray(4) = -1 Then
					derror = "The destination Asset / Location is an ancestor of the source Asset / Location. This is not permitted when moving an Asset. Please select an Asset / Location that does not belong to the source Asset / Location."
					newjs = newjs + "   top.mcalert('warning','Asset Move Message','There was a problem moving the Asset. The details of the problem are described below.<br><br><u>Problem Details</u>:<br><br>" & Replace(derror,"'","\'") & "','bg_okprint',700,370,'sounds/error.wav');" + nl
					newjs = newjs + "	top.fraAssets.oActionNode = null;" + nl
					returnmessage = ""
				ElseIf OutArray(4) = -2 Then
					derror = "Unable to Transfer Rotating Parts.<br>Rotating Parts can only be transferred via Work Order Cost functionality."
					newjs = newjs + "   top.mcalert('warning','Asset Move Message','There was a problem moving the Asset. The details of the problem are described below.<br><br><u>Problem Details</u>:<br><br>" & Replace(derror,"'","\'") & "','bg_okprint',700,370,'sounds/error.wav');" + nl
					newjs = newjs + "	top.fraAssets.oActionNode = null;" + nl
					returnmessage = ""
				Else
					'newjs = newjs + "		top.pendingmsg = 'The Asset has been Deleted';" + nl
					If UCase(Trim(Request.QueryString("isroot"))) = "Y" Then
						newjs = newjs + "		top.fraAssets.movecomplete('paste',true);" + nl
					Else
						newjs = newjs + "		top.fraAssets.movecomplete('paste',false);" + nl
					End If

					returnmessage = "The Asset / Location has been Moved"
				End If

		Case "CLONE"

				' <SERVER-SIDE VALIDATION STUFF HERE>

				Call SetScriptTimeoutTo(60)

				aok = True
				errorfield = ""
				returnmessage = ""
				returnclass = "standardmessage"
				
				Dim ClonePK
				Dim txtRowVersionUserPK,txtRowVersionInitials
				Dim copychildren,copyproc,copysa,copyattach,CopyLocationsOnly
				Dim idsearch,idreplace,namesearch,namereplace,txtserial
				Dim RepairCenterPK, RepairCenterID, RepairCenterName
				Dim NumberOfCopies

				txtRowVersionUserPK = Request.QueryString("txtRowVersionUserPK")
				txtRowVersionInitials = Request.QueryString("txtRowVersionInitials")

				If UCase(Trim(Request.QueryString("copyoptions"))) = "N" Then

					' Set the defaults since the user didn't do paste special
					' which gives them these options...

					idsearch = Null
					idreplace = Null
					namesearch = Null
					namereplace = Null
					txtserial = Null
					copychildren = 1
					copyproc = 1
					copysa = 1
					copyattach = 1
					repaircenterpk = Null
					repaircenterid = Null
					repaircentername = Null
					CopyLocationsOnly = 0
					NumberOfCopies = 1

				Else

					idsearch = Trim(Request.QueryString("idsearch"))
					idreplace = Trim(Request.QueryString("idreplace"))
					If idsearch = "" or idreplace = "" Then
						idsearch = Null
						idreplace = Null
					End If
					namesearch = Trim(Request.QueryString("namesearch"))
					namereplace = Trim(Request.QueryString("namereplace"))
					If namesearch = "" or namereplace = "" Then
						namesearch = Null
						namereplace = Null
					End If
					txtserial = Trim(Request.QueryString("txtserial"))
					If txtserial = "" Then
						txtserial = Null
					End If
					repaircenterpk = Trim(Request.QueryString("txtrepaircenterpk"))
					repaircenterid = Trim(Request.QueryString("txtrepaircenter"))
					repaircentername = Trim(Request.QueryString("txtrepaircenterdesch"))
					If repaircenterpk = "" Then
						repaircenterpk = Null
						repaircenterid = Null
						repaircentername = Null
					End If

					copychildren = Trim(UCase(Request.QueryString("copychildren")))
					If copychildren = "ON" Then
						copychildren = 1
					Else
						copychildren = 0
					End If
					copyproc = Trim(UCase(Request.QueryString("copyproc")))
					If copyproc = "ON" Then
						copyproc = 1
					Else
						copyproc = 0
					End If
					copysa = Trim(UCase(Request.QueryString("copysa")))
					If copysa = "ON" Then
						copysa = 1
					Else
						copysa = 0
					End If
					copyattach = Trim(UCase(Request.QueryString("copyattach")))
					If copyattach = "ON" Then
						copyattach = 1
					Else
						copyattach = 0
					End If
					CopyLocationsOnly = Trim(UCase(Request.QueryString("CopyLocationsOnly")))
					If CopyLocationsOnly = "ON" Then
						CopyLocationsOnly = 1
					Else
						CopyLocationsOnly = 0
					End If

					NumberOfCopies = Request.QueryString("NumberOfCopies")

				End If

                Dim z
                For z = 1 to NumberOfCopies
                    Call db.RunSP("MC_CloneAsset",Array(Array("@FromPK", adInteger, adParamInput, 4, CLng(Request.QueryString("frompk"))),Array("@ToPK", adInteger, adParamInput, 4, CLng(Trim(Request.QueryString("topk")))),Array("@UserPK", adInteger, adParamInput, 4, Request.QueryString("txtRowVersionUserPK").Item),Array("@Initials", MC_ADVARCHAR, adParamInput, 5, Trim(Mid(Request.QueryString("txtRowVersionInitials").Item,1,5))),Array("@CopyProc", adBoolean, adParamInput, 1, copyproc),Array("@CopySA", adBoolean, adParamInput, 1, copysa),Array("@CopyAttach", adBoolean, adParamInput, 1, copyattach),Array("@CopyChildren", adBoolean, adParamInput, 1, copychildren),Array("@idsearch", MC_ADVARCHAR, adParamInput, 100, idsearch),Array("@idreplace", MC_ADVARCHAR, adParamInput, 100, idreplace),Array("@namesearch", MC_ADVARCHAR, adParamInput, 150, namesearch),Array("@namereplace", MC_ADVARCHAR, adParamInput, 150, namereplace),Array("@txtserial", MC_ADVARCHAR, adParamInput, 50, txtserial),Array("@repaircenterpk", adInteger, adParamInput, 4, repaircenterpk),Array("@repaircenterid", MC_ADVARCHAR, adParamInput, 25, repaircenterid),Array("@repaircentername", MC_ADVARCHAR, adParamInput, 50, repaircentername),Array("@CopyLocationsOnly", adBoolean, adParamInput, 1, CopyLocationsOnly),Array("@newassetpk", adInteger, adParamOutPut, 4, "")),OutArray)
                Next
                ClonePK = OutArray(17)				

				If Not db.dok Then
					newjs = newjs + "   top.mcalert('warning','Asset Message','There was a problem cloning the Asset. The details of the problem are described below. You can try to SAVE the Asset again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.<br><br><u>Problem Details</u>:<br><br>" & Replace(db.derror,"'","\'") & "','bg_okprint',700,370,'sounds/error.wav');" + nl
					newjs = newjs + "	top.fraAssets.oActionNode = null;" + nl
					returnmessage = ""
				Else
					'newjs = newjs + "		top.pendingmsg = 'The Asset has been Deleted';" + nl
					newjs = newjs + "	    if (top.addontheflymode == true) {top.addonthefly_recupdated = true;}" + nl	
					newjs = newjs + "		// kill the current record" + nl
					newjs = newjs + "		if (top.fraAssets.oActionNode != null)" + nl
					newjs = newjs + "		{" + nl
					If UCase(Trim(Request.QueryString("isroot"))) = "Y" Then
						newjs = newjs + "			top.refreshcurrentexplorer();" + nl
					Else
						newjs = newjs + "			top.fraAssets.refreshtreenode(top.fraAssets.oActionNode);" + nl
					End If
					newjs = newjs + "		} else {" + nl

					newjs = newjs + "		top.recordkey = 'kv=" & CStr(ClonePK) & "';" + nl
					newjs = newjs + "       if (top.eCurrentLI) {top.eCurrentLI.setAttribute('mckey','"& CStr(ClonePK) &"');}" + nl
					newjs = newjs + "		var myrefreshli = top.fraAssets.findLI(top.fraTopic.document.mcform.txtParentPK.value);" + nl
					newjs = newjs + "		if (myrefreshli && myrefreshli.parentElement.id == 'ulRoot') { top.refreshcurrentexplorer(); } else { " + nl
					newjs = newjs + "		top.fraAssets.refreshtreenode(myrefreshli); }" + nl					
					newjs = newjs + "		top.navadd('AS','kv=" & CStr(ClonePK) & "');" + nl
		
					newjs = newjs + "		}" + nl

					newjs = newjs + "		top.fraAssets.oActionNode = null;" + nl
					returnmessage = "The Asset / Location has been Cloned"
				End If

				ResetScriptTimeOut

		Case Else

			returnmessage = "Error: Could not determine Action"
			returnclass = "errormessage"

	End Select

	db.CloseClientConnection
	Set db = Nothing

	Call DoResponse(lastaction,newjs,True,True,False,keyvalue)

End Sub

Sub Setup(justdata)

	If justdata Then
		If newrecord Then
			Response.Write "top.showtabinfo(top.firsttabinfo,false);" + nl
			Response.Write "top.changepage(myframe.document.getElementById('pagetabs_31'));" + nl
			Response.Write "try {myform.txtParent.focus();} catch(e) {};" + nl
			Response.Write("top.isdirty();")+nl
		Else
			Response.Write "	top.showtabinfo(top.currenttabinfo.id,false);" + nl
			Response.Write "	top.dofocus();" + nl
		End If
		'Response.Write "	mydoc.body.style.backgroundColor = '#808080';" + nl
		'Response.Write "	mydoc.body.style.backgroundImage = '';" + nl
	Else
		Response.Write "	top.setstarttabinfo();" + nl
		Response.Write nl
		Response.Write "function resizeit() {" & nl
		Response.Write("	myframe.document.getElementById('fraWO').style.height = top.fraTopic.innerHeight - 140 + 'px';") + nl
		Response.Write("	myframe.document.getElementById('fraPJ').style.height = top.fraTopic.innerHeight - 140 + 'px';") + nl
		Response.Write("	myframe.document.getElementById('fraPD').style.height = top.fraTopic.innerHeight - 140 + 'px';") + nl
		Response.Write("	myframe.document.getElementById('fraWOGroup').style.height = top.fraTopic.innerHeight - 140 + 'px';") + nl
		Response.Write("	myframe.document.getElementById('fraWODownTime').style.height = top.fraTopic.innerHeight - 140 + 'px';") + nl
		Response.Write("    myframe.document.getElementById('fraPhoto').style.height = top.fraTopic.innerHeight - 140 + 'px';") + nl
		Response.Write("	myframe.document.getElementById('leaseFrame').style.height = top.fraTopic.innerHeight - 105 + 'px';") + nl
		Response.Write("	myframe.document.getElementById('contractsFrame').style.height = top.fraTopic.innerHeight - 105 + 'px';") + nl
		'Response.Write("	myframe.document.getElementById('fileFrame').style.height = '200px';") + nl
		Response.Write "}" & nl
	End If

	Response.Write("")+nl
End Sub

Function SetupFields(themode,justdata)
	Response.Write("") + nl

	Response.Write("// SetupFields Fields")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl

	Dim db,rs

	Select Case themode

		Case "NEW"

			' Standard Settings
			'------------------------------------------------------------------------------------
			Response.Write("top.recordkey = 'kv=NEW'")+nl
			Response.Write("top.moduleaction = 'asave';")+nl
			Response.Write("top.fraTabbar.document.getElementById('header_keytext').innerText='New Asset';")+nl
			Response.Write("top.fraTabbar.document.getElementById('header_keytext2').innerText='';")+nl							
			Response.Write("top.walk(top.fraTabbar.document.getElementById('header_keydiv'), null, top, false, true);")+nl
			Response.Write("top.clearallfields();")+nl
			Response.Write("top.updatestatusbar('',null,'');")+nl
			Response.Write("top.removebuttongrpTopic();")

			ClearTables

			' UI Settings
			'------------------------------------------------------------------------------------
			Response.Write("	myform.imgRotatingPart.disabled = false;")
			Response.Write("	myform.txtRotatingPart.readOnly = false;")
			Response.Write("	myform.txtRotatingPart.style.backgroundColor = '#ffffff';")
	
			Response.Write("	myform.imgMaintainableTool.disabled = false;")
			Response.Write("	myform.txtMaintainableTool.readOnly = false;")
			Response.Write("	myform.txtMaintainableTool.style.backgroundColor = '#ffffff';")

			' Defaults
			'------------------------------------------------------------------------------------

			' Set the Asset Level High so that if they come from an AssetLevel setting they don't
			' have access to - then this will allow them to add a new record.
			Response.Write("	if (myform.txtAssetLevel) {myform.txtAssetLevel.value = '1000';}")+nl

			Response.Write("	myform.txtIsUp.checked = true;")+nl
			Response.Write("	myform.txtRequesterCanView.checked = true;")+nl
			Response.Write("    myframe.document.getElementById('assetmessagebox').innerText = '';")+nl
			Response.Write("    myframe.document.getElementById('assetmessagebox').style.display = 'none';")+nl
			Response.Write("	top.setpk('txtRepairCenterPK','');")+nl
			Response.Write("	myform.txtRepairCenter.value = '';")+nl
			Response.Write("	top.setdesc(myframe.document.getElementById('txtRepairCenterDesc'),'');")+nl
			Response.Write("	myform.txtType.value = 'L';")+nl
			Response.Write("	top.setdesc(myframe.document.getElementById('txtTypeDesc'),'Location') ;")+nl
			Response.Write("	top.setassetdivs();") + nl

			Response.Write("	if (myform.txtMeter1RollDownMethod) {")+nl
			Response.Write("	myform.txtMeter1RollDownMethod[0].checked = false;")+nl
			Response.Write("	myform.txtMeter1RollDownMethod[1].checked = false;")+nl
			Response.Write("}")+nl

			Response.Write("	if (myform.txtMeter2RollDownMethod) {")+nl
			Response.Write("	myform.txtMeter2RollDownMethod[0].checked = false;")+nl
			Response.Write("	myform.txtMeter2RollDownMethod[1].checked = false;")+nl
			Response.Write("}")+nl

			Response.Write("	myform.txtAssetName.disabled = false;")+nl

			Response.Write("	myform.txtModel.disabled = false;")+nl
			Response.Write("	myform.txtModel.style.border = '1px solid #c0c0c0';")+nl
			Response.Write("	myform.txtModelNumber.disabled = false;")+nl
			Response.Write("	myform.txtModelNumber.style.border = '1px solid #c0c0c0';")+nl
			Response.Write("	myform.txtModelNumberMFG.disabled = false;")+nl
			Response.Write("	myform.txtModelNumberMFG.style.border = '1px solid #c0c0c0';")+nl
			Response.Write("	myform.txtModelLine.disabled = false;")+nl
			Response.Write("	myform.txtModelLine.style.border = '1px solid #c0c0c0';")+nl
			Response.Write("	myform.txtModelSeries.disabled = false;")+nl
			Response.Write("	myform.txtModelSeries.style.border = '1px solid #c0c0c0';")+nl
			Response.Write("	myform.txtSystemPlatform.disabled = false;")+nl
			Response.Write("	myform.txtSystemPlatform.style.border = '1px solid #c0c0c0';")+nl
			Response.Write("	myform.txtManufacturer.disabled = false;")+nl
			Response.Write("	myform.txtType.disabled = false;")+nl

			Call OutputUDFLabels(Null,"Asset")

			Response.Write("	myform.txtClassIndustry.value = 'F-G';")+nl
			Response.Write("	top.setdesc(myframe.document.getElementById('txtClassIndustryDesc'),'Facility - General');")+nl

			'CB AUTOCALC FIX 1/10/2013
			'============================================================================================================================
			Response.Write("	if (myform.txtRiskLevelAutoCalc) {myform.txtRiskLevelAutoCalc.checked = true;}")+nl
			Response.Write("	if (myform.txtPMRequiredAutoCalc) {myform.txtPMRequiredAutoCalc.checked = true;}")+nl
			'============================================================================================================================

			Response.Write("	myform.txtRiskAssessmentGroup.value = 'F';")+nl
			Response.Write("	top.setdesc(myframe.document.getElementById('txtRiskAssessmentGroupDesc'),'Facilities');")+nl
			Response.Write("	myframe.setriskfactors('F');")+nl
			Response.Write("	myframe.SetPMRequired(myform.txtPMRequired);")+nl
			

			Response.Write("	if (top.addontheflymode == true && top.dialogArguments != null && top.dialogArguments.setdefaults)") + nl
			Response.Write("	{") + nl
			Response.Write("		if (top.dialogArguments.AssetPK != null)") + nl
			Response.Write("		{") + nl
			Response.Write("			top.endprocess();") + nl
			Response.Write("			top.dovalid('AS','txtParent','AS','N',null,null,top.dialogArguments.AssetPK);")+nl
			Response.Write("		}") + nl
			Response.Write("	}") + nl

			If Not Request.QueryString("PAssetPK") = "" Then
				Response.Write("			top.endprocess();") + nl
				Response.Write("			top.dovalid('AS','txtParent','AS','N',null,null," & Request.QueryString("PAssetPK") & ");")+nl
			End If

		Case "EDIT"

			' Standard Settings
			'------------------------------------------------------------------------------------
			Response.Write("top.recordkey = 'kv=" + CStr(keyvalue) + "'")+nl
			Response.Write("top.moduleaction = 'asave';")+nl
			Response.Write("top.showactions('aasrecbar');")+nl
			Response.Write("top.showbuttongrpTopic('bg_normal');")+nl

	End Select

	' Reset the Classification Count to 0
	Response.Write("	myform.txtClassificationCount.value = '';")+nl

End Function

Function EnableDisableFields(themode)
	Response.Write("")+nl

	Response.Write("// Enable/Disable Fields")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl

	Select Case themode

		Case "EDIT"
			'Response.Write("myform.txtPmNo.disabled = true;")+nl
			'Response.Write("myform.txtPmNo.className = 'disabled';")+nl

		Case "NEW"
			Response.Write("	myframe.document.getElementById('parentfields').style.display = '';") + nl
			Response.Write("	myframe.document.getElementById('parentfieldsspacer').style.display = 'none';") + nl
			'Response.Write("myform.txtPmNo.disabled = true;")+nl
			'Response.Write("myform.txtPmNo.className = 'disabled';")+nl

	End Select

	Response.Write("")+nl

End Function

Function ClearTables()
	Response.Write("")+nl

	Response.Write("myframe.document.getElementById('as_details').setAttribute('dataloaded','Y');") + nl
	Response.Write("myframe.document.getElementById('as_meter').setAttribute('dataloaded','Y');") + nl
	Response.Write(nl)
	Response.Write("function loadtab71()")+nl
	Response.Write("{")+nl
	Response.Write(nl)

	' CUSTOMIZED
	'--------------------------------------------------------------------------------------------------------------
	Response.Write("	top.displayicon(mydoc.images.asseticon,'');")+nl
	Response.Write("	top.displayphoto(mydoc.images.assetphoto,'');")+nl
	Response.Write("	top.displayicon(mydoc.images.passeticon,'images/icons/blankicon_g.gif');")+nl
	'--------------------------------------------------------------------------------------------------------------

	Response.Write(nl)
	Response.Write("}")+nl
	Response.Write("loadtab71();")+nl

	Response.Write(nl)
	Response.Write("myframe.document.getElementById('as_specifications').setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab74()")+nl
	Response.Write("{")+nl
	Response.Write(nl)

	Response.Write("// Clear Specifications")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("top.cleartable(myframe.document.getElementById('osp1'));")+nl

	Response.Write("// Clear Materials")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("top.cleartable(myframe.document.getElementById('oam1'));")+nl

	Response.Write("// Clear Labor")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("top.cleartable(myframe.document.getElementById('oal1'));")+nl

	Response.Write("// Clear Contracts")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("top.cleartable(myframe.document.getElementById('oao1'));")+nl
	Response.Write("myframe.document.getElementById('contractsFrame').style.display = 'none';") + nl

	Response.Write("// Clear Occupants / Requesters")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("top.cleartable(myframe.document.getElementById('oar1'));")+nl

	Response.Write(nl)
	Response.Write("}")+nl

	Response.Write(nl)
	Response.Write("myframe.document.getElementById('as_procedures').setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab75()")+nl
	Response.Write("{")+nl
	Response.Write(nl)

	Response.Write("// Clear Procedures")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("top.cleartable(myframe.document.getElementById('oap1'));")+nl

	Response.Write(nl)
	Response.Write("}")+nl

	Response.Write(nl)
	Response.Write("myframe.document.getElementById('as_attach').setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab76()")+nl
	Response.Write("{")+nl
	Response.Write(nl)

	Response.Write("// Clear Documents")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("top.cleartable(myframe.document.getElementById('oat1'));")+nl

	Response.Write(nl)
	Response.Write("// Clear Notes Rows")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("top.cleartable(myframe.document.getElementById('ono1'));")+nl

	Response.Write("// Clear Images")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("   mydoc.getElementById('fraPhoto').style.display = 'none';") + nl

	Response.Write("// Clear Misc Attachments")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	'Response.Write("myframe.document.getElementById('fileFrame').style.display = 'none';") + nl

	Response.Write(nl)
	Response.Write("// Clear Rules Rows")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("top.cleartable(myframe.document.getElementById('orl1'));")+nl

	Response.Write(nl)
	Response.Write("}")+nl

	Response.Write(nl)
	Response.Write("myframe.document.getElementById('as_history').setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab77()")+nl
	Response.Write("{")+nl
	Response.Write(nl)


	Response.Write(nl)
	Response.Write("}")+nl

	
	Response.Write("myframe.document.getElementById('as_lease').setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab78()")+nl
	Response.Write("{")+nl
	Response.Write(nl)

	Response.Write("// Clear Leases")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("myframe.document.getElementById('leaseFrame').style.display = 'none';") + nl

	Response.Write(nl)
	Response.Write("}")+nl

End Function

Function writerecorddata(db)

	Dim killdbonend,rs,outarray,dok,derror,newjs,playsound
	Dim rs_AssetName

	writerecorddata = True

	If IsNull(db) or (Not IsObject(db)) Then
		Set db = New ADOHelper
		killdbonend = True
	Else
		killdbonend = False
	End If

	Dim prefvalue, prefdesc, prefpk, PM_AssetsInserviceOnly

	If GetDefaultPreference(db,False,"PM_AssetsInserviceOnly",prefvalue, prefdesc, prefpk) Then
		If UCase(prefvalue) = "YES" Then
			PM_AssetsInserviceOnly = True
		Else
			PM_AssetsInserviceOnly = False
		End If
	End If

	Set rs = db.RunSPReturnMultiRS("MC_GetAsset",Array(Array("@AssetPK", adInteger, adParamInput, 4, keyvalue),Array("@UserPK", adInteger, adParamInput, 4, GetSession("UserPK"))),outarray)
	Call dok_check(db,"Asset Message","There was a problem retrieving the selected Asset. The details of the problem are described below. You can try to retrieve the Asset again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")

	If rs.eof Then
		db.dok = False
		db.derror = "The selected Asset could not be found. Please refresh your Asset list (to ensure another user has not deleted this Asset) and try again."
		' End Process here (and it will actually end the process again...but no big deal)
		' We need to end it here because the navadd will fail if we don't because we are
		' actually inside a process at this point in the game.
		Response.Write("top.endprocess();") + nl
		Response.Write("top.navadd('AS','kv=');") + nl
		Response.Write("top.mcalert('warning','Asset Not Found','" & db.derror & "','bg_okprint',700,235,'sounds/error.wav');") + nl
		Call CloseObj(rs)
		If killdbonend Then
			db.CloseClientConnection
			Set db = Nothing
		End If
		writerecorddata = False
		Exit Function
	End If

	' Output All Fields Below

	rs_AssetName = JSEncode(RS("AssetName"))

	Response.Write("")+nl

	Response.Write("	// Clear all hidden fields")+nl
	Response.Write("	// -------------------------------------------------------------------------")+nl
	Response.Write("	top.clearallhiddenfields();")+nl

	Response.Write("	// Set the Primary Key and Row Version")+nl
	Response.Write("	// -------------------------------------------------------------------------")+nl
	Response.Write("	top.sethidden('kv','" & JSEncode(RS("AssetPK")) & "');")+nl
	Response.Write("	top.sethidden('txtRowVersionDate','" & NullCheck(RS("RowVersionDate")) & "');")+nl
	Response.Write("")+nl

	Response.Write("	// Update the Header Bar")+nl
	Response.Write("	// -------------------------------------------------------------------------")+nl
	Dim aorl
	'If RS("IsLocation") Then
	'	aorl = ""
	'Else
	'	aorl = JSEncode(RS("ClassificationName"))
	'	If Not aorl = "" Then
	'		aorl = aorl & ": "
	'	End If
	'End If
	'Response.Write("	top.fraTabbar.document.getElementById('header_keytext').innerHTML='" & aorl)
	Response.Write("	top.fraTabbar.document.getElementById('header_keytext').innerHTML=''")+nl
	Response.Write("	top.fraTabbar.document.getElementById('header_keytext2').innerHTML=' ")
	Response.Write(JSEncode(RS("AssetName")))
	Response.Write("&nbsp;&nbsp;(")
	Response.Write(JSEncode(RS("AssetID")))
	Response.Write(")")
	'Response.Write(JSEncode(RS("AssetPK")))
	Response.Write("';")+nl
	Response.Write("top.walk(top.fraTabbar.document.getElementById('header_keydiv'), null, top, false, true);")+nl
	Response.Write("")+nl

	Response.Write("	// Update the Status Bar")+nl
	Response.Write("	// -------------------------------------------------------------------------")+nl
	Response.Write("	top.updatestatusbar('" & JSEncode(RS("AssetName")) & "','','" & JSEncode(RS("TypeDesc")) & "');")+nl
	Response.Write("")+nl

	Response.Write("	// Set the Tabs that MUST be loaded prior to Saving")+nl
	Response.Write("	// -------------------------------------------------------------------------")+nl
	Response.Write("	top.checktabsbeforesave = 'top.checktabdataloaded(\'as_details\');top.checktabdataloaded(\'as_meter\');top.checktabdataloaded(\'as_risk\');';")+nl
	Response.Write("")+nl

	Response.Write("	// Output all fields on Tabs that might not be loaded prior to Saving ")+nl
	Response.Write("	// -------------------------------------------------------------------------")+nl
	Call OutputDefaults(rs,1)

	Response.Write("	myframe.document.getElementById('as_details').setAttribute('dataloaded','N');") + nl
	Response.Write("")+nl
	Response.Write("function loadtab71()")+nl
	Response.Write("{")+nl
	Response.Write("")+nl

		Response.Write("	// Write Asset Record Field Data")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl

		Call OutputDefaults(rs,2)

	Response.Write("}")+nl
	Response.Write(nl)

	Response.Write("	myframe.document.getElementById('as_risk').setAttribute('dataloaded','N');") + nl
	Response.Write("")+nl
	Response.Write("function loadtab72()")+nl
	Response.Write("{")+nl
	Response.Write("")+nl

		Response.Write("	// Write Additional Info Field Data")+nl
		Response.Write("	// Write Risk Field Data")+nl

		Call OutputDefaults(rs,3)					

	Response.Write("}")+nl	
	Response.Write(nl)	
	

	Response.Write("	myframe.document.getElementById('as_meter').setAttribute('dataloaded','N');") + nl
	Response.Write("")+nl
	Response.Write("function loadtab73()")+nl
	Response.Write("{")+nl
	Response.Write("")+nl

		Response.Write("	// Write Additional Info Field Data")+nl
		Response.Write("	// Write Meter Field Data")+nl

		Call OutputDefaults(rs,4)

	Response.Write("}")+nl
	Response.Write(nl)

	Response.Write("	myframe.document.getElementById('as_specifications').setAttribute('dataloaded','N');") + nl
	Response.Write("")+nl
	Response.Write("function loadtab74()")+nl
	Response.Write("{")+nl
	Response.Write("")+nl

		Set rs = rs.NextRecordset

		Response.Write(nl)
		Response.Write("	// Write Specification Data")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	top.cleartable(myframe.document.getElementById('osp1'));")+nl
		Dim LineTemplate,SpecHiOK,SpecLowOK, TrackHistory, ValueOutOfRangeWO
		Dim HeaderTemplate,LastCategory,RowHeaderName
		LastCategory = "#!@#"
		If rs.recordcount > 25 Then
			If Not NullCheck(rs("CategoryName")) = "!NONE!" Then
				HeaderTemplate = 6
				Response.Write("	top.hiderowonstart = 'Y';")+nl
			End If
		Else
			HeaderTemplate = 5
			Response.Write("	top.hiderowonstart = null;")+nl
		End If
		Do Until rs.eof
			If Not LastCategory = NullCheck(rs("CategoryName")) Then
				LastCategory = NullCheck(rs("CategoryName"))
				If Not NullCheck(rs("CategoryName")) = "!NONE!" Then
					If NullCheck(rs("CategoryName")) = "" Then
						RowHeaderName = "Uncategorized"
					Else
						RowHeaderName = JSEncode(rs("CategoryName"))
					End If
					Response.Write("	if (top.tablesection_showhide) {top.builddatarow(myframe.document.getElementById('osp1body')," & HeaderTemplate & ",null,'','','CA',false,'',null,null,'" & RowHeaderName & "');}")+nl
				End If
			End If
			SpecHiOK = False
			SpecLowOK = False
			If RS("TrackHistory") Then
				TrackHistory = "<img src=""../../images/taskchecked.gif"">"
			Else
				TrackHistory = "<img src=""../../images/taskline.gif"">"
			End If
			If RS("ValueOutOfRangeWO") Then
				ValueOutOfRangeWO = "<img src=""../../images/taskchecked.gif"">"
			Else
				ValueOutOfRangeWO = "<img src=""../../images/taskline.gif"">"
			End If
			LineTemplate = 2
			If Not NullCheck(rs("ValueHi")) = "" and Not NullCheck(rs("ValueNumeric")) = "" Then
				'If IsNumeric(rs("ValueNumeric")) and IsNumeric(rs("ValueHi")) Then
					If CDbl(rs("ValueNumeric")) > CDbl(rs("ValueHi")) Then
						LineTemplate = 4
					Else
						SpecHiOK = True
					End If
				'End If
			End If
			If Not NullCheck(rs("ValueLow")) = "" and Not NullCheck(rs("ValueNumeric")) = "" Then
				'If IsNumeric(rs("ValueNumeric")) and IsNumeric(rs("ValueLow")) Then
					If CDbl(rs("ValueNumeric")) < CDbl(rs("ValueLow")) Then
						LineTemplate = 4
					Else
						SpecLowOK = True
					End If
				'End If
			End If
			If (SpecLowOK or SpecHiOK) and Not LineTemplate = 4 Then
				LineTemplate = 3
			End If
			Response.Write("	top.builddatarow(myframe.document.getElementById('osp1body')," & LineTemplate & ",null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','" & RS("SpecificationPK") & "','SP',false,'',null,null,'" & JSEncode(RS("SpecificationName")) & "','" & JSEncode(RS("ValueText")) & "','" & DateNullCheck(RS("ValueDate")) & "','" & NullCheck(RS("ValueNumeric")) & "','" & NullCheck(RS("ValueLow")) & "','" & NullCheck(RS("ValueHi")) & "','" & NullCheck(RS("ValueOptimal")) & "','" & ValueOutOfRangeWO & "','" & JSEncode(RS("ProcedureID")) & "<mcbr> / " & JSEncode(RS("ProcedureID")) & "','" & TrackHistory & "');")+nl
			If BitNullCheck(rs("UseLookupTable")) and Not NullCheck(rs("LookupTable")) = "" Then
				' Add Drop-Down Image
				Response.Write("    myframe.document.getElementById('osp1body').rows[myframe.document.getElementById('osp1body').rows.length-1].cells[2].setAttribute('mclookuptable','" & JSEncode(rs("LookupTable")) & "');")+nl
			End If
			rs.MoveNext
		Loop
		Response.Write("	top.hiderowonstart = null;")+nl

		Set rs = rs.NextRecordset

		Response.Write(nl)
		Response.Write("	// Write Material Data")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	top.cleartable(myframe.document.getElementById('oam1'));")+nl
		LastCategory = "#!@#"
		If rs.recordcount > 25 Then
			If Not NullCheck(rs("CategoryName")) = "!NONE!" Then
				HeaderTemplate = 4
				Response.Write("	top.hiderowonstart = 'Y';")+nl
			End If
		Else
			HeaderTemplate = 3
			Response.Write("	top.hiderowonstart = null;")+nl
		End If
		Do Until rs.eof
			If Not LastCategory = NullCheck(rs("CategoryName")) Then
				LastCategory = NullCheck(rs("CategoryName"))
				If Not NullCheck(rs("CategoryName")) = "!NONE!" Then
					If NullCheck(rs("CategoryName")) = "" Then
						RowHeaderName = "Uncategorized"
					Else
						RowHeaderName = JSEncode(rs("CategoryName"))
					End If
					Response.Write("	if (top.tablesection_showhide) {top.builddatarow(myframe.document.getElementById('oam1body')," & HeaderTemplate & ",null,'','','CA',false,'',null,null,'" & RowHeaderName & "');}")+nl
				End If
			End If
			Response.Write("	top.builddatarow(myframe.document.getElementById('oam1body'),2,null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','" & RS("PartPK") & "','IN',false,'" & JSEncode(RS("Photo")) & "',null,null,'" & JSEncode(RS("PartID")) & "','" & JSEncode(RS("PartName")) & "','" & NullCheck(RS("QTY")) & "','" & JSEncode(RS("Comments")) & "');")+nl
			rs.MoveNext
		Loop
		Response.Write("	top.hiderowonstart = null;")+nl

		Set rs = rs.NextRecordset

		Response.Write(nl)
		Response.Write("	// Write Labor / Contact Data")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	top.cleartable(myframe.document.getElementById('oal1'));")+nl
		Dim AutoAssign, BackupResource
		LastCategory = "#!@#"
		If rs.recordcount > 25 Then
			If Not NullCheck(rs("CategoryName")) = "!NONE!" Then
				HeaderTemplate = 4
				Response.Write("	top.hiderowonstart = 'Y';")+nl
			End If
		Else
			HeaderTemplate = 3
			Response.Write("	top.hiderowonstart = null;")+nl
		End If
		Dim jobtitle
		Do Until rs.eof
			If Not LastCategory = NullCheck(rs("CategoryName")) Then
				LastCategory = NullCheck(rs("CategoryName"))
				If Not NullCheck(rs("CategoryName")) = "!NONE!" Then
					If NullCheck(rs("CategoryName")) = "" Then
						RowHeaderName = "Uncategorized"
					Else
						RowHeaderName = JSEncode(rs("CategoryName"))
					End If
					Response.Write("	if (top.tablesection_showhide) {top.builddatarow(myframe.document.getElementById('oal1body')," & HeaderTemplate & ",null,'','','CA',false,'',null,null,'" & RowHeaderName & "');}")+nl
				End If
			End If
			If RS("AutoAssign") Then
				AutoAssign = "<img src=""../../images/taskchecked.gif"">"
			Else
				AutoAssign = "<img src=""../../images/taskline.gif"">"
			End If
			If RS("BackupResource") Then
				BackupResource = "<img src=""../../images/taskchecked.gif"">"
			Else
				BackupResource = "<img src=""../../images/taskline.gif"">"
			End If
			If Not JSEncode(RS("JobTitle")) = "" Then
				jobtitle = " - " & JSEncode(RS("JobTitle"))
			Else
				jobtitle = ""
			End If
			Response.Write("	top.builddatarow(myframe.document.getElementById('oal1body'),2,null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','" & RS("LaborPK") & "','" & RS("ModuleID") & "',false,'" & JSEncode(RS("Photo")) & "',null,null,'" & JSEncode(RS("LaborName")) & "','" & JSEncode(RS("LaborTypeDesc")) & jobtitle & "','" & JSEncode(RS("JobTitle")) & "','" & JSEncode(RS("PhoneWork")) & "','" & JSEncode(RS("PhoneHome")) & "','" & JSEncode(RS("PhoneMobile")) & "','" & JSEncode(RS("Pager")) & "','" & JSEncode(RS("Fax")) & "','" & JSEncode(RS("Email")) & "','" & JSEncode(RS("RepairCenterID")) & "','" & NullCheck(RS("Priority")) & "','" & AutoAssign & "','" & BackupResource & "');")+nl
			rs.MoveNext
		Loop
		Response.Write("	top.hiderowonstart = null;")+nl

		Set rs = rs.NextRecordset

		Response.Write(nl)
		Response.Write("	// Write Contract Data")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	top.cleartable(myframe.document.getElementById('oao1'));")+nl
		Do Until rs.eof
			Response.Write("	top.builddatarow(myframe.document.getElementById('oao1body'),2,null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','" & RS("CompanyPK") & "','CM',false,'" & JSEncode(RS("Photo")) & "',null,null,'" & JSEncode(RS("CompanyName")) & "','" & JSEncode(RS("Phone")) & "','" & DateNullCheck(RS("PeriodStart")) & "','" & DateNullCheck(RS("PeriodEnd")) & "','" & JSEncode(RS("VendorContractNum")) & "','" & JSEncode(RS("ContractSummary")) & "');")+nl
			rs.MoveNext
		Loop

		Response.Write("	var apk = top.recordkey.substring(3, top.recordkey.length);") + nl
		Response.Write("	var caContent = myframe.document.getElementById('contractsFrame');") + nl
		Response.Write("	caContent.src = '../Contracts/Contracts.asp?apk=' + apk;") + nl
		Response.Write("    myframe.resizeit();")+nl
		Response.Write("	caContent.style.display = '';") + nl

		Set rs = rs.NextRecordset

		Response.Write(nl)
		Response.Write("	// Write Occupant / Tenant Data")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	top.cleartable(myframe.document.getElementById('oar1'));")+nl
		Dim IsLease,Active,IsPrimary
		Do Until rs.eof
			If RS("IsLease") Then
				IsLease = "<img src=""../../images/taskchecked.gif"">"
			Else
				IsLease = "<img src=""../../images/taskline.gif"">"
			End If
			If RS("Active") Then
				Active = "<img src=""../../images/taskchecked.gif"">"
			Else
				Active = "<img src=""../../images/taskline.gif"">"
			End If
			If RS("IsPrimary") Then
				IsPrimary = "<img src=""../../images/taskchecked.gif"">"
			Else
				IsPrimary = "<img src=""../../images/taskline.gif"">"
			End If
			Response.Write("	top.builddatarow(myframe.document.getElementById('oar1body'),2,null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','" & RS("LaborPK") & "','RQ',false,'" & JSEncode(RS("Photo")) & "',null,null,'" & JSEncode(RS("LaborName")) & "','" & JSEncode(RS("PhoneWork")) & "','" & JSEncode(RS("PhoneHome")) & "','" & JSEncode(RS("PhoneMobile")) & "','" & JSEncode(RS("Pager")) & "','" & JSEncode(RS("Fax")) & "','" & JSEncode(RS("Email")) & "','" & DateNullCheck(RS("StartDate")) & "','" & DateNullCheck(RS("EndDate")) & "','" & DateNullCheck(RS("ExpireDate")) & "','" & NullCheck(RS("Percentage")) & "','" & IsLease & "','" & IsPrimary & "','" & Active & "');")+nl
			rs.MoveNext
		Loop

	Response.Write("}")+nl
	Response.Write(nl)

	Response.Write("	myframe.document.getElementById('as_procedures').setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab75()")+nl
	Response.Write("{")+nl

		Set rs = rs.NextRecordset

		Response.Write(nl)
		Response.Write("	// Write Procedure Schedule Data")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	top.cleartable(myframe.document.getElementById('oap1'));")+nl
		Dim nextscheduled, Frequency, scheduledisabled, waitingonwomsg, nogenmsg
		nextscheduled=""
		Do Until rs.eof
			nogenmsg=False
			waitingonwomsg = ""
			If BitNullCheck(RS("ScheduleDisabled")) Then
				scheduledisabled = "<img src=""../../images/taskline.gif"">"
			Else
				scheduledisabled = "<img src=""../../images/taskchecked.gif"">"
			End If
			If (Not NullCheck(rs("WOPK")) = "") Then
				If (BitNullCheck(rs("IsOpenWO"))) Then
					waitingonwomsg = waitingonwomsg & "<br/><font color=""red"">Waiting for <span style=""cursor:pointer;"" onclick=""event.cancelBubble=true;top.doviewedit(\'WO\',\'" & rs("WOPK") & "\',top.fraTopic);""><u>WO #" & NullCheck(rs("WOID")) & "</u><span> targeted<br/>for " & DateNullCheck(rs("TargetDate")) & " to be closed</font>"
					nogenmsg=True
				Else
					If NullCheck(rs("WOID")) = "" Then
						waitingonwomsg = waitingonwomsg & "<br/><font color=""green"">Last Work Order: Deleted</font>"
					Else
						waitingonwomsg = waitingonwomsg & "<br/><font color=""green"">Last Work Order: <span style=""cursor:pointer;"" onclick=""event.cancelBubble=true;top.doviewedit(\'WO\',\'" & rs("WOPK") & "\',top.fraTopic);""><u>WO #" & NullCheck(rs("WOID")) & "</u><br/>targeted for " & DateNullCheck(rs("TargetDate")) & "</span></font>"
					End If
				End If
			End If

			If Not NullCheck(rs("PMCycleStartBy")) = "" and Not NullCheck(rs("PMCycleStartBy")) = "PM" Then
			Select Case UCase(Trim(rs("PMCycleStartBy")))
				Case "AC"
					waitingonwomsg = waitingonwomsg & "<br/><font color=""royalblue""><b>Next Scheduled Date is based on the PM Cycle <br/>Start Date for the Account tied to this Asset.</b></font>"
				Case "AS"
					waitingonwomsg = waitingonwomsg & "<br/><font color=""royalblue""><b>Next Scheduled Date is based on the PM Cycle <br/>Start Date for this Asset.</b></font>"
				Case "CL"
					waitingonwomsg = waitingonwomsg & "<br/><font color=""royalblue""><b>Next Scheduled Date is based on the PM Cycle <br/>Start Date for the Classification tied to this Asset.</b></font>"
				Case "CLP"
					waitingonwomsg = waitingonwomsg & "<br/><font color=""royalblue""><b>Next Scheduled Date is based on the PM Cycle <br/>Start Date for the Classificatoin Parent tied to this Asset.</b></font>"
				Case "CU"
					waitingonwomsg = waitingonwomsg & "<br/><font color=""royalblue""><b>Next Scheduled Date is based on the PM Cycle <br/>Start Date for the Customer tied to this Asset.</b></font>"
				Case "DP"
					waitingonwomsg = waitingonwomsg & "<br/><font color=""royalblue""><b>Next Scheduled Date is based on the PM Cycle <br/>Start Date for the Department tied to this Asset.</b></font>"
				Case "RC"
					waitingonwomsg = waitingonwomsg & "<br/><font color=""royalblue""><b>Next Scheduled Date is based on the PM Cycle <br/>Start Date for the Repair Center tied to this Asset.</b></font>"
				Case "SH"
					waitingonwomsg = waitingonwomsg & "<br/><font color=""royalblue""><b>Next Scheduled Date is based on the PM Cycle <br/>Start Date for the Shop tied to this Asset.</b></font>"
				Case "ZN"
					waitingonwomsg = waitingonwomsg & "<br/><font color=""royalblue""><b>Next Scheduled Date is based on the PM Cycle <br/>Start Date for the Zone tied to this Asset.</b></font>"
			End Select
			End If
			If Not BitNullCheck(rs("IsUp")) Then
				waitingonwomsg = waitingonwomsg & "<br/><font color=""red"">Asset is not In-Service</font>"
				If PM_AssetsInserviceOnly Then
					nogenmsg=True
				End If
			End If
			If Not BitNullCheck(rs("PMActive")) Then
				waitingonwomsg = waitingonwomsg & "<br/><font color=""red"">The PM Schedule is Disabled (for all Assets)</font>"
				nogenmsg=True
			Else
				If BitNullCheck(rs("ScheduleDisabled")) Then
					waitingonwomsg = waitingonwomsg & "<br/><font color=""red"">PM is not Enabled for this Asset</font>"
					nogenmsg=True
				End If
			End If
			If nogenmsg Then
				waitingonwomsg = waitingonwomsg & "<br/><font color=""red""><b>PM will not generate for this Asset</b></font>"
			End If
			Frequency = NullCheck(RS("Frequency"))
			If RS("PMEnded") Then
				nextscheduled = "<IMG src=""../../images/warn.gif"" style=""margin-right:4px;"" border=0>PM Ended"
			Else
				If Not Frequency = "METER" Then
					nextscheduled = DateNullCheck(RS("NextScheduledDate"))
				Else
					If Not NullCheck(RS("Meter1NextInterval")) = "0" and _
					   Not NullCheck(RS("Meter2NextInterval")) = "0" Then
						nextscheduled = CLng(RS("Meter1NextInterval")) & " " & JSEncode(RS("Meter1UnitsDesc")) & " / " & CLng(RS("Meter2NextInterval")) & " " & JSEncode(RS("Meter2UnitsDesc"))
					ElseIf Not NullCheck(RS("Meter1NextInterval")) = "0" Then
						nextscheduled = CLng(RS("Meter1NextInterval")) & " " & JSEncode(RS("Meter1UnitsDesc"))
					ElseIf Not NullCheck(RS("Meter2NextInterval")) = "0" Then
						nextscheduled = CLng(RS("Meter2NextInterval")) & " " & JSEncode(RS("Meter2UnitsDesc"))
					End If
				End If
			End If
			Response.Write("	top.builddatarow(myframe.document.getElementById('oap1body'),2,null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','" & RS("PMPK") & "','PM',false,'',null,null,'" & JSEncode(RS("PMID")) & "<br>" & JSEncode(RS("PMName")) & waitingonwomsg & "','" & JSEncode(RS("FrequencyDesc")) & "','" & NullCheck(RS("RouteOrder")) & "','" & DateNullCheck(RS("LastGeneratedDate")) & "','" & DateNullCheck(RS("LastCompletedDate")) & "','" & NullCheck(RS("Meter1ReadingLastInterval")) & "','" & NullCheck(RS("Meter2ReadingLastInterval")) & "','" & NullCheck(RS("TimesCounter")) & "','" & nextscheduled & "','" & NullCheck(RS("PMCounter")) & "','" & JSEncode(RS("ProcedureID")) & "<br>" & JSEncode(RS("ProcedureName")) & "','" & JSEncode(RS("RepairCenterID")) & "<mcbr> / " & JSEncode(RS("RepairCenterName")) & "','" & JSEncode(RS("StockRoomID")) & "<mcbr> / " & JSEncode(RS("StockRoomName")) & "','" & JSEncode(RS("ToolRoomID")) & "<mcbr> / " & JSEncode(RS("ToolRoomName")) & "','" & JSEncode(RS("ShopID")) & "<mcbr> / " & JSEncode(RS("ShopName")) & "','" & JSEncode(RS("ShiftID")) & "<mcbr> / " & JSEncode(RS("ShiftName")) & "','" & JSEncode(RS("SupervisorID")) & "<mcbr> / " & JSEncode(RS("SupervisorName")) & "','" & JSEncode(RS("AccountID")) & "<mcbr> / " & JSEncode(RS("AccountName")) & "','" & JSEncode(RS("DepartmentID")) & "<mcbr> / " & JSEncode(RS("DepartmentName")) & "','" & JSEncode(RS("TenantID")) & "<mcbr> / " & JSEncode(RS("TenantName")) & "','" & JSEncode(RS("ProjectID")) & "<mcbr> / " & JSEncode(RS("ProjectName")) & "','" & scheduledisabled & "');")+nl
			rs.MoveNext
		Loop

		Set rs = rs.NextRecordset

		Response.Write(nl)
		Response.Write("	// Write Task Data")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	top.cleartable(myframe.document.getElementById('ota1'));")+nl
		Dim Fail, Complete
		Do Until rs.eof
			If BitNullCheck(RS("Fail")) Then
				Fail = "<img src=""../../images/taskfailed.gif"">"
			Else
				Fail = "<img src=""../../images/taskunchecked.gif"">"
			End If
			If BitNullCheck(RS("TaskComplete")) Then
				Complete = "<img src=""../../images/taskchecked.gif"">"
			Else
				Complete = "<img src=""../../images/taskline.gif"">"
			End If
			Response.Write("	top.builddatarow(myframe.document.getElementById('ota1body'),2,null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','" & RS("WOPK") & "','WO',false,'',null,null,'" & JSEncode(RS("TaskAction")) & "<br/><span style=""color:gray;"">" & JSEncode(RS("Comments")) & "</span>','" & DateNullCheck(RS("LastCompleted")) & "','" & JSEncode(RS("WOID")) & "','" & DateNullCheck(RS("Complete")) & "','" & NullCheck(RS("Rate")) & "','" & NullCheck(RS("Measurement")) & "','" & JSEncode(RS("ToolID")) & "<mcbr> / " & JSEncode(RS("ToolID")) & "','" & JSEncode(RS("Initials")) & "','" + Fail + "','" + Complete + "');")+nl
			rs.MoveNext
		Loop

	Response.Write("}")+nl

	Response.Write("	myframe.document.getElementById('as_attach').setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab76()")+nl
	Response.Write("{")+nl

		Response.Write(nl)
		Response.Write("	// Write Images Data")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	if (top.fraTopic.TabControl2_currentTab == null || top.fraTopic.TabControl2_currentTab.id == 'TabControl2_t3') { ")+nl
		Response.Write(nl)
		Response.Write("	top.dosearchASImages();}") + nl

		Response.Write("	if (top.fraTopic.TabControl2_currentTab == null || top.fraTopic.TabControl2_currentTab.id == 'TabControl2_t4') { ")+nl
		Response.Write(nl)
		Response.Write("	top.dosearchASImages(null,null,true);}") + nl

		Response.Write("	mydoc.getElementById('fraPhoto').style.display = '';") + nl
		Response.Write("    myframe.resizeit();")+nl

		Set rs = rs.NextRecordset

		Response.Write(nl)
		Response.Write("	// Build Attachments Rows")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	top.cleartable(myframe.document.getElementById('oat1'));")+nl

		Call OutputAttachments(rs)

		Set rs = rs.NextRecordset

		Response.Write(nl)
		Response.Write("	// Build Note Rows")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	top.cleartable(myframe.document.getElementById('ono1'));")+nl
		Dim NoteCustom1, NoteCustom2, NoteTemplate
		Do Until rs.eof
			NoteTemplate = NullCheck(rs("NoteTemplate"))
			If NoteTemplate = "" Then
				NoteTemplate = 2
			End If
			If RS("Custom1") Then
				NoteCustom1 = "<img src=""../../images/taskchecked.gif"">"
			Else
				NoteCustom1 = "<img src=""../../images/taskline.gif"">"
			End If
			If RS("Custom2") Then
				NoteCustom2 = "<img src=""../../images/taskchecked.gif"">"
			Else
				NoteCustom2 = "<img src=""../../images/taskline.gif"">"
			End If
			If NullCheck(RS("PK")) = "-1" Then
				Response.Write("	top.builddatarow(myframe.document.getElementById('ono1body')," & NoteTemplate & ",null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','','',false,'',null,null,'" & DateNullCheckAT(RS("NoteDate")) & "','" & TimeNullCheckAT(RS("NoteDate")) & "','" & JSEncode(RS("Initials")) & "','" & JSEncode(RS("Note")) & "','" & NoteCustom1 & "','" & NoteCustom2 & "');")+nl
			Else
				Response.Write("	top.builddatarow(myframe.document.getElementById('ono1body')," & NoteTemplate & ",null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','','',false,'',null,null,'" & DateNullCheckAT(RS("NoteDate")) & "','" & TimeNullCheckAT(RS("NoteDate")) & "','" & JSEncode(RS("Initials")) & "','" & JSEncode(RS("Note")) & "','" & NoteCustom1 & "','" & NoteCustom2 & "');")+nl
			End If
			rs.MoveNext
		Loop

		Response.Write(nl)
		Response.Write("	// Write Misc Attachments")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		'Response.Write("	myframe.fileFrame.location.replace(top.path+'modules/attachments/mc_genericFilesMain.asp?pid="&keyvalue&"&table=assets');")+nl					
		Response.Write("    myframe.resizeit();")+nl
		If Not newrecord Then
		'Response.Write("	myframe.document.getElementById('fileFrame').style.display = '';") + nl
		End If

		Response.Write("outputrules();") + nl

		Response.Write(nl)

	Response.Write("}")+nl

	Response.Write("	myframe.document.getElementById('as_history').setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("	myframe.flag_history_t1 = false;") + nl
	Response.Write("	myframe.flag_history_t2 = false;") + nl
	Response.Write("	myframe.flag_history_t3 = false;") + nl
	Response.Write("	myframe.flag_history_t4 = false;") + nl
	Response.Write("	myframe.flag_history_t5 = false;") + nl
	Response.Write(nl)
	Response.Write("function loadtab77()")+nl
	Response.Write("{")+nl

		Response.Write("	myframe.TabControl1_changeTabsAction();")+nl

	Response.Write("}")+nl

	Response.Write("	myframe.document.getElementById('as_lease').setAttribute('dataloaded','N');") + nl
	Response.Write("function loadtab78()")+nl
	Response.Write("{")+nl

		Response.Write(nl)
		Response.Write("	// Write Lease Data")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("    var apk = top.recordkey.substring(3, top.recordkey.length);") + nl
		Response.Write("	var gaContent = myframe.document.getElementById('leaseFrame');") + nl
		Response.Write("	gaContent.src = '../Leases/leases.asp?apk=' + apk;") + nl
		Response.Write("    myframe.resizeit();")+nl
		Response.Write("	gaContent.style.display = '';") + nl

	Response.Write("}")+nl

	Set rs = rs.NextRecordset
	Call OutputUDFLabels(rs,"")

	Set rs = rs.NextRecordset
	Call OutputFavorites(rs)

	Set rs = rs.NextRecordset	
	Call OutputRules(rs)

	Call CloseObj(rs)

	If killdbonend Then
		db.CloseClientConnection
		Set db = Nothing
	End If

End Function

Sub OutputDefaults(rs,section)

	Dim LockFieldsinAssetModule, LockNameinAssetModule
	If BitNullCheck(RS("LockFieldsinAssetModule")) Then
		LockFieldsinAssetModule = True
	Else
		LockFieldsinAssetModule = False
	End If
	If BitNullCheck(RS("LockNameinAssetModule")) Then
		LockNameinAssetModule = True
	Else
		LockNameinAssetModule = False
	End If

	If CLng(section) = 1 or CLng(section) = 5 Then

		Response.Write("	top.setpk('txtAssetPK','" & RS("AssetPK") & "');")+nl

		Response.Write("	if (myform.txtFormName) {myform.txtFormName.value = '" & JSEncode(RS("FormName")) & "';};")+nl

		If NullCheck(RS("PMCycleStartBy")) = "" Then
			PMCycleStartBy = "PM"
			PMCycleStartByDesc = "PM Settings"
		Else
			PMCycleStartBy = JSEncode(RS("PMCycleStartBy"))
			PMCycleStartByDesc = JSEncode(RS("PMCycleStartByDesc"))
		End If

		Response.Write("	if (myform.txtPMCycleStartBy) {myform.txtPMCycleStartBy.value = '" & PMCycleStartBy & "';}")+nl
		Response.Write("	if (myform.txtPMCycleStartBy) {top.setdesc(myframe.document.getElementById('txtPMCycleStartByDesc'),'" & PMCycleStartByDesc & "');}")+nl
		Response.Write("	if (myform.txtPMCycleStartDate) {myform.txtPMCycleStartDate.value = '" & DateNullCheck(RS("PMCycleStartDate")) & "';}")+nl

		Response.Write("	if (myframe.enabledisablecyclestartdate) {myframe.enabledisablecyclestartdate();};")+nl

		Response.Write("	if (myform.txtAssetLevel) {myform.txtAssetLevel.value = '" & JSEncode(RS("AssetLevel")) & "';}")+nl

	End If
	If CLng(section) = 2 or CLng(section) = 5 Then

		Response.Write("	// Write Asset Record Field Data")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	myform.txtAsset.value = '" & JSEncode(RS("AssetID")) & "';")+nl

		If LockNameinAssetModule Then
			Response.Write("	myform.txtAssetName.disabled = true;")+nl
		Else
			Response.Write("	myform.txtAssetName.disabled = false;")+nl
		End If

		If LockFieldsinAssetModule Then
			Response.Write("	myform.txtModel.disabled = true;")+nl
			Response.Write("	myform.txtModel.style.border = 0;")+nl
			Response.Write("	myform.txtModelNumber.disabled = true;")+nl
			Response.Write("	myform.txtModelNumber.style.border = 0;")+nl
			Response.Write("	myform.txtModelNumberMFG.disabled = true;")+nl
			Response.Write("	myform.txtModelNumberMFG.style.border = 0;")+nl
			Response.Write("	myform.txtModelLine.disabled = true;")+nl
			Response.Write("	myform.txtModelLine.style.border = 0;")+nl
			Response.Write("	myform.txtModelSeries.disabled = true;")+nl
			Response.Write("	myform.txtModelSeries.style.border = 0;")+nl
			Response.Write("	myform.txtSystemPlatform.disabled = true;")+nl
			Response.Write("	myform.txtSystemPlatform.style.border = 0;")+nl
			Response.Write("	myform.txtManufacturer.disabled = true;")+nl
			Response.Write("	myform.txtType.disabled = true;")+nl
		Else
			Response.Write("	myform.txtModel.disabled = false;")+nl
			Response.Write("	myform.txtModel.style.border = '1px solid #c0c0c0';")+nl
			Response.Write("	myform.txtModelNumber.disabled = false;")+nl
			Response.Write("	myform.txtModelNumber.style.border = '1px solid #c0c0c0';")+nl
			Response.Write("	myform.txtModelNumberMFG.disabled = false;")+nl
			Response.Write("	myform.txtModelNumberMFG.style.border = '1px solid #c0c0c0';")+nl
			Response.Write("	myform.txtModelLine.disabled = false;")+nl
			Response.Write("	myform.txtModelLine.style.border = '1px solid #c0c0c0';")+nl
			Response.Write("	myform.txtModelSeries.disabled = false;")+nl
			Response.Write("	myform.txtModelSeries.style.border = '1px solid #c0c0c0';")+nl
			Response.Write("	myform.txtSystemPlatform.disabled = false;")+nl
			Response.Write("	myform.txtSystemPlatform.style.border = '1px solid #c0c0c0';")+nl
			Response.Write("	myform.txtManufacturer.disabled = false;")+nl
			Response.Write("	myform.txtType.disabled = false;")+nl
		End If

		' CUSTOMIZED
		'--------------------------------------------------------------------------------------------------------------
		Response.Write("	myform.txtAssetName.value = '" & JSEncode(RS("AssetName")) & "';")+nl
		Response.Write("	top.setpk('txtParentPK','" & RS("ParentPK") & "');")+nl
		Response.Write("	myform.txtParent.value = '" & JSEncode(RS("ParentID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtParentDesc'),'" & JSEncode(RS("ParentName")) & "');")+nl
		'Response.Write("	myform.txtParentAssetLevel.value = '" & JSEncode(RS("AssetLevel")) & "';")+nl
		If NullCheck(RS("ParentPK")) = "" Then
			Response.Write("	myframe.document.getElementById('parentfields').style.display = 'none';") + nl
			Response.Write("	myframe.document.getElementById('parentfieldsspacer').style.display = '';") + nl
			Response.Write("	top.setmultiobj(mydoc.images.buttonaction_delete,'style.display','none');") + nl
		Else
			Response.Write("	myframe.document.getElementById('parentfields').style.display = '';") + nl
			Response.Write("	myframe.document.getElementById('parentfieldsspacer').style.display = 'none';") + nl
			Response.Write("	top.setmultiobj(mydoc.images.buttonaction_delete,'style.display','');") + nl
		End If
		Response.Write("	top.setassetdivs('" & NullCheck(RS("Type")) & "');") + nl
		'--------------------------------------------------------------------------------------------------------------

		'If RS("IsLocation") = 0 Then
		'Response.Write("	myform.txtIsLocation.checked = false;")+nl
		'Else
		'Response.Write("	myform.txtIsLocation.checked = true;")+nl
		'End If

        'CB2->Update for expired contracts.
        dim iB, lRs,iRs,hasExpiredItems
        Set iB = New ADOHelper
        Set iRs = iB.RunSQLReturnRS("select top 1 ct.contractpk from assetcontract AC WITH (NOLOCK) join [contract] CT WITH (NOLOCK) on (ct.contractpk = ac.contractpk ) where ct.active=1 AND ac.assetpk=" & RS("AssetPK") & " ", "")
            If Not iRs.EOF Then
                Set lRs = iB.RunSQLReturnRS("SELECT ConName = CT.ContractName, ConNumber = CT.VendorContractNum FROM AssetContract AC WITH ( NOLOCK ) JOIN [Contract] CT WITH (NOLOCK) ON ( AC.ContractPK = CT.ContractPK ) WHERE AC.AssetPK = " & RS("AssetPK") & " AND CT.PeriodEnd < GetDate() AND CT.Active=1 ", "")
                    If Not lRs.EOF Then
                        hasExpiredItems = ""
                        Do While Not lRs.EOF
                            If Len(lRs("ConNumber")) > 30 Then
                                hasExpiredItems = hasExpiredItems & "Contract: " & Left(lRs("ConNumber"),27) & "...<br>"
                            Else
                                hasExpiredItems = hasExpiredItems & "Contract: " & lRs("ConNumber") & "<br>"
                            End IF
                            lRs.MoveNext
                        Loop
            			Response.Write("    myframe.document.getElementById('assetmessagebox').style.display = '';")+nl
		                Response.Write("    myframe.document.getElementById('assetmessagebox').style.color = 'red';")+nl
		                Response.Write("    myframe.document.getElementById('assetmessagebox').style.backgroundColor = '#FFF1F1';")+nl
		                Response.Write("    myframe.document.getElementById('assetmessagebox').innerHTML = 'Service Contract(s) Expired<br><span style=""font-weight:normal; color:#363636;"">" & Replace(hasExpiredItems,"'","''") & "</span>';")+nl
                    Else
            			Response.Write("    myframe.document.getElementById('assetmessagebox').style.display = '';")+nl
		                Response.Write("    myframe.document.getElementById('assetmessagebox').style.color = 'green';")+nl
		                Response.Write("    myframe.document.getElementById('assetmessagebox').style.backgroundColor = '#E7FFE7';")+nl
		                Response.Write("    myframe.document.getElementById('assetmessagebox').innerText = 'Covered Under Service Contract';")+nl
                    End If
                Set lRs = Nothing
            Else
       			Response.Write("    myframe.document.getElementById('assetmessagebox').style.display = 'none';")+nl
		        Response.Write("    myframe.document.getElementById('assetmessagebox').innerText = '';")+nl
            End If
        Set iRs = Nothing

		'If RS("isActiveContracts") Then
		'    Response.Write("    myframe.document.getElementById('assetmessagebox').style.color = 'green';")+nl
		'    Response.Write("    myframe.document.getElementById('assetmessagebox').style.backgroundColor = '#E7FFE7';")+nl
		'    Response.Write("    myframe.document.getElementById('assetmessagebox').innerText = 'Covered Under Service Contract';")+nl
		'ElseIf RS("isContracts") Then
		'    Response.Write("    myframe.document.getElementById('assetmessagebox').style.color = 'red';")+nl
		'    Response.Write("    myframe.document.getElementById('assetmessagebox').style.backgroundColor = '#FFF1F1';")+nl
		'    Response.Write("    myframe.document.getElementById('assetmessagebox').innerText = 'Service Contract Expired';")+nl
		'Else
		'    Response.Write("    myframe.document.getElementById('assetmessagebox').innerText = '';")+nl
		'End If

		If RS("RequesterCanView") = 0 Then
		Response.Write("	myform.txtRequesterCanView.checked = false;")+nl
		Else
		Response.Write("	myform.txtRequesterCanView.checked = true;")+nl
		End If
		If RS("IsUp") = 0 Then
		Response.Write("	myform.txtIsUp.checked = false; myform.txtIsUp.setAttribute('mcorigvalue',myform.txtIsUp.checked);")+nl
		Else
		Response.Write("	myform.txtIsUp.checked = true; myform.txtIsUp.setAttribute('mcorigvalue',myform.txtIsUp.checked);")+nl
		End If
		Response.Write("	myform.txtIsUp.disabled = false; ")+nl

		Response.Write("	myform.txtStatus.value = '" & JSEncode(RS("Status")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtStatusDesc'),'" & JSEncode(RS("StatusDesc")) & "');")+nl

		Response.Write("	myform.txtType.value = '" & JSEncode(RS("Type")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtTypeDesc'),'" & JSEncode(RS("TypeDesc")) & "');")+nl
		Response.Write("	myform.txtModel.value = '" & JSEncode(RS("Model")) & "';")+nl

		Response.Write("	myform.txtModelNumber.value = '" & JSEncode(RS("ModelNumber")) & "';")+nl
		Response.Write("	myform.txtModelNumberMFG.value = '" & JSEncode(RS("ModelNumberMFG")) & "';")+nl
		Response.Write("	myform.txtModelLine.value = '" & JSEncode(RS("ModelLine")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtModelLineDesc'),'" & JSEncode(RS("ModelLineDesc")) & "');")+nl
		Response.Write("	myform.txtModelSeries.value = '" & JSEncode(RS("ModelSeries")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtModelSeriesDesc'),'" & JSEncode(RS("ModelSeriesDesc")) & "');")+nl
		Response.Write("	myform.txtSystemPlatform.value = '" & JSEncode(RS("SystemPlatform")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtSystemPlatformDesc'),'" & JSEncode(RS("SystemPlatformDesc")) & "');")+nl

		Response.Write("	myform.txtSerial.value = '" & JSEncode(RS("Serial")) & "';")+nl

		Response.Write("	top.setpk('txtClassificationPK','" & RS("ClassificationPK") & "');")+nl
		Response.Write("	myform.txtClassification.value = '" & JSEncode(RS("ClassificationID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtClassificationDesc'),'" & JSEncode(RS("ClassificationName")) & "');")+nl

		Response.Write("	top.setpk('txtRotatingPartPK','" & RS("RotatingPartPK") & "');")+nl
		Response.Write("	myform.txtRotatingPart.value = '" & JSEncode(RS("RotatingPartID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtRotatingPartDesc'),'" & JSEncode(RS("RotatingPartName")) & "');")+nl

		'@$CUSTOMISED
		Response.Write("	top.setpk('txtMaintainableToolPK','" & RS("MaintainableToolPK") & "');")+nl
		Response.Write("	myform.txtMaintainableTool.value = '" & JSEncode(RS("MaintainableToolID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtMaintainableToolDesc'),'" & JSEncode(RS("MaintainableToolName")) & "');")+nl
		
		IF(NOT RS("RotatingPartPK") = "") THEN
			Response.Write("	myform.imgMaintainableTool.disabled = true;")
			Response.Write("	myform.txtMaintainableTool.readOnly = true;")
			Response.Write("	myform.txtMaintainableTool.style.backgroundColor = '#f0f0f0';")
		ELSE
			Response.Write("	myform.imgMaintainableTool.disabled = false;")
			Response.Write("	myform.txtMaintainableTool.readOnly = false;")
			Response.Write("	myform.txtMaintainableTool.style.backgroundColor = '#ffffff';")
		END IF
		
		IF(NOT RS("MaintainableToolPK") = "") THEN
			Response.Write("	myform.imgRotatingPart.disabled = true;")
			Response.Write("	myform.txtRotatingPart.readOnly = true;")
			Response.Write("	myform.txtRotatingPart.style.backgroundColor = '#f0f0f0';")
		ELSE
			Response.Write("	myform.imgRotatingPart.disabled = false;")
			Response.Write("	myform.txtRotatingPart.readOnly = false;")
			Response.Write("	myform.txtRotatingPart.style.backgroundColor = '#ffffff';")
		END IF
		
		'@$END

		Response.Write("	myform.txtSystem.value = '" & JSEncode(RS("System")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtSystemDesc'),'" & JSEncode(RS("SystemDesc")) & "');")+nl
		Response.Write("	myform.txtPriority.value = '" & JSEncode(RS("Priority")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtPriorityDesc'),'" & JSEncode(RS("PriorityDesc")) & "');")+nl
		Response.Write("	myform.txtVicinity.value = '" & JSEncode(RS("Vicinity")) & "';")+nl
		Response.Write("	top.setpk('txtAccountPK','" & RS("AccountPK") & "');")+nl
		Response.Write("	myform.txtAccount.value = '" & JSEncode(RS("AccountID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtAccountDesc'),'" & JSEncode(RS("AccountName")) & "');")+nl
		Response.Write("	top.setpk('txtDepartmentPK','" & RS("DepartmentPK") & "');")+nl
		Response.Write("	myform.txtDepartment.value = '" & JSEncode(RS("DepartmentID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtDepartmentDesc'),'" & JSEncode(RS("DepartmentName")) & "');")+nl
		Response.Write("	top.setpk('txtTenantPK','" & RS("TenantPK") & "');")+nl
		Response.Write("	myform.txtTenant.value = '" & JSEncode(RS("TenantID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtTenantDesc'),'" & JSEncode(RS("TenantName")) & "');")+nl
		Response.Write("	top.setpk('txtRepairCenterPK','" & RS("RepairCenterPK") & "');")+nl
		Response.Write("	myform.txtRepairCenter.value = '" & JSEncode(RS("RepairCenterID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtRepairCenterDesc'),'" & JSEncode(RS("RepairCenterName")) & "');")+nl
		Response.Write("	top.setpk('txtShopPK','" & NullCheck(RS("ShopPK")) & "') ;")+nl
		Response.Write("	myform.txtShop.value = '" & JSEncode(RS("ShopID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtShopDesc'),'" & JSEncode(RS("ShopName")) & "') ;")+nl
		Response.Write("	top.setpk('txtVendorPK','" & RS("VendorPK") & "');")+nl
		Response.Write("	myform.txtVendor.value = '" & JSEncode(RS("VendorID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtVendorDesc'),'" & JSEncode(RS("VendorName")) & "');")+nl
		Response.Write("	top.setpk('txtManufacturerPK','" & RS("ManufacturerPK") & "');")+nl
		Response.Write("	myform.txtManufacturer.value = '" & JSEncode(RS("ManufacturerID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtManufacturerDesc'),'" & JSEncode(RS("ManufacturerName")) & "');")+nl
		Response.Write("	myform.txtWarrantyExpire.value = '" & DateNullCheck(RS("WarrantyExpire")) & "';")+nl

		Response.Write("	if (myform.txtPurchaseType) {myform.txtPurchaseType.value = '" & JSEncode(RS("PurchaseType")) & "';}")+nl
		Response.Write("	if (myform.txtPurchaseType) {top.setdesc(myframe.document.getElementById('txtPurchaseTypeDesc'),'" & JSEncode(RS("PurchaseTypeDesc")) & "');}")+nl

		Response.Write("	myform.txtPurchasedDate.value = '" & DateNullCheck(RS("PurchasedDate")) & "';")+nl
		Response.Write("	myform.txtPurchaseOrder.value = '" & JSEncode(RS("PurchaseOrder")) & "';")+nl
		Response.Write("	myform.txtPurchaseCost.value = '" & RS("PurchaseCost") & "';")+nl
		Response.Write("	myform.txtInstallDate.value = '" & DateNullCheck(RS("InstallDate")) & "';")+nl
		Response.Write("	myform.txtReplaceDate.value = '" & DateNullCheck(RS("ReplaceDate")) & "';")+nl
		Response.Write("	myform.txtReplacementCost.value = '" & RS("ReplacementCost") & "';")+nl
		Response.Write("	myform.txtDisposalDate.value = '" & DateNullCheck(RS("DisposalDate")) & "';")+nl

		Response.Write("	myform.txtInsuranceCarrier.value = '" & JSEncode(RS("InsuranceCarrier")) & "';")+nl
		Response.Write("	myform.txtInsurancePolicy.value = '" & JSEncode(RS("InsurancePolicy")) & "';")+nl
		Response.Write("	myform.txtLeaseNumber.value = '" & JSEncode(RS("LeaseNumber")) & "';")+nl
		Response.Write("	myform.txtRegistrationDate.value = '" & DateNullCheck(RS("RegistrationDate")) & "';")+nl
		Response.Write("	myform.txtTechnology.value = '" & JSEncode(RS("Technology")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtTechnologyDesc'),'" & JSEncode(RS("TechnologyDesc")) & "') ;")+nl
		Response.Write("	myform.txtAddress.value = '" & JSEncode(RS("Address")) & "';")+nl
		Response.Write("	myform.txtCity.value = '" & JSEncode(RS("City")) & "';")+nl
		Response.Write("	myform.txtState.value = '" & JSEncode(RS("State")) & "';")+nl
		Response.Write("	myform.txtZip.value = '" & JSEncode(RS("Zip")) & "';")+nl
		Response.Write("	myform.txtCountry.value = '" & JSEncode(RS("Country")) & "';")+nl

		If RS("IsPredictive") = 0 Then
		Response.Write("	myform.txtIsPredictive.checked = false;")+nl
		Else
		Response.Write("	myform.txtIsPredictive.checked = true;")+nl
		End If

		' DrawingUpdatesNeeded is nullable so check if true rather than false
		If RS("DrawingUpdatesNeeded") Then
		Response.Write("	myform.txtDrawingUpdatesNeeded.checked = true;")+nl
		Else
		Response.Write("	myform.txtDrawingUpdatesNeeded.checked = false;")+nl
		End If

		If RS("DisplayMapOnWO") = 0 Then
		Response.Write("	myform.txtDisplayMapOnWO.checked = false;")+nl
		Else
		Response.Write("	myform.txtDisplayMapOnWO.checked = true;")+nl
		End If

		' CUSTOMIZED
		'--------------------------------------------------------------------------------------------------------------
		Response.Write("	myform.txtInstructions.value = unescape('" & JSEncode(RS("Instructions")) & "');")+nl
		'--------------------------------------------------------------------------------------------------------------

		If RS("InstructionsToWO") = 0 Then
		Response.Write("	myform.txtInstructionsToWO.checked = false;")+nl
		Else
		Response.Write("	myform.txtInstructionsToWO.checked = true;")+nl
		End If

		Response.Write("	myform.txtLicenseNumber.value = '" & JSEncode(RS("LicenseNumber")) & "';")+nl

		Response.Write("	top.setpk('txtOperatorPK','" & RS("OperatorPK") & "');")+nl
		Response.Write("	myform.txtOperator.value = '" & JSEncode(RS("OperatorID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtOperatorDesc'),'" & JSEncode(RS("OperatorName")) & "');")+nl

		Response.Write("	top.setpk('txtContactPK','" & RS("ContactPK") & "');")+nl
		Response.Write("	myform.txtContact.value = '" & JSEncode(RS("ContactID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtContactDesc'),'" & JSEncode(RS("ContactName")) & "');")+nl

		Response.Write("	myform.txtIcon.value = '" & JSEncode(RS("Icon")) & "';")+nl
		Response.Write("	myform.txtPhoto.value = '" & JSEncode(RS("Photo")) & "';")+nl

		'Response.Write("	myform.txtUser_Guid.value = '" & JSEncode(RS("User_Guid")) & "';")+nl

		' CUSTOMIZED
		'--------------------------------------------------------------------------------------------------------------
		Response.Write("	top.displayphoto(mydoc.images.assetphoto,'" & JSEncode(RS("Photo")) & "');")+nl
		Response.Write("	top.displayicon(mydoc.images.asseticon,'" & JSEncode(RS("Icon")) & "');")+nl
		If Not RS("ParentIsLocation") = vbNull Then
			Response.Write("	top.displayicon(mydoc.images.passeticon,'" & JSEncode(RS("ParentIcon")) & "'," & LCase(JSEncode(RS("ParentIsLocation"))) & ");")+nl
		End If
		'--------------------------------------------------------------------------------------------------------------

		Response.Write("	top.setpk('txtZonePK','" & RS("ZonePK") & "');")+nl
		Response.Write("	myform.txtZone.value = '" & JSEncode(RS("ZoneID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtZoneDesc'),'" & JSEncode(RS("ZoneName")) & "');")+nl
		Response.Write("	myform.txtCounty.value = '" & JSEncode(RS("County")) & "';")+nl
		Response.Write("	myform.txtYearBuilt.value = '" & DateNullCheck(RS("YearBuilt")) & "';")+nl
		Response.Write("	myform.txtMajorRenovations.value = '" & JSEncode(RS("MajorRenovations")) & "';")+nl
		Response.Write("	myform.txtSquareFootage.value = '" & RS("SquareFootage") & "';")+nl
		Response.Write("	myform.txtConstructionCode.value = '" & JSEncode(RS("ConstructionCode")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtConstructionCodeDesc'),'" & JSEncode(RS("ConstructionCodeDesc")) & "');")+nl
		Response.Write("	myform.txtNumberOfStories.value = '" & RS("NumberOfStories") & "';")+nl
		Response.Write("	myform.txtISOProtection.value = '" & JSEncode(RS("ISOProtection")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtISOProtectionDesc'),'" & JSEncode(RS("ISOProtectionDesc")) & "');")+nl
		Response.Write("	myform.txtAutoSprinkler.value = '" & JSEncode(RS("AutoSprinkler")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtAutoSprinklerDesc'),'" & JSEncode(RS("AutoSprinklerDesc")) & "');")+nl
		Response.Write("	myform.txtSmokeAlarm.value = '" & JSEncode(RS("SmokeAlarm")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtSmokeAlarmDesc'),'" & JSEncode(RS("SmokeAlarmDesc")) & "');")+nl
		Response.Write("	myform.txtHeatAlarm.value = '" & JSEncode(RS("HeatAlarm")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtHeatAlarmDesc'),'" & JSEncode(RS("HeatAlarmDesc")) & "');")+nl
		Response.Write("	myform.txtFloodZone.value = '" & JSEncode(RS("FloodZone")) & "';")+nl
		Response.Write("	myform.txtQuakeZone.value = '" & JSEncode(RS("QuakeZone")) & "';")+nl
		Response.Write("	myform.txtExt100Feet.value = '" & JSEncode(RS("Ext100Feet")) & "';")+nl
		Response.Write("	myform.txtOperatingUnits.value = '" & RS("OperatingUnits") & "';")+nl
		Response.Write("	myform.txtEstimatedValue.value = '" & RS("EstimatedValue") & "';")+nl
		Response.Write("	top.setpk('txtResponsibilityRepairPK','" & RS("ResponsibilityRepairPK") & "');")+nl
		Response.Write("	myform.txtResponsibilityRepair.value = '" & JSEncode(RS("ResponsibilityRepairID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtResponsibilityRepairDesc'),'" & JSEncode(RS("ResponsibilityRepairName")) & "');")+nl
		Response.Write("	top.setpk('txtResponsibilityPMPK','" & RS("ResponsibilityPMPK") & "');")+nl
		Response.Write("	myform.txtResponsibilityPM.value = '" & JSEncode(RS("ResponsibilityPMID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtResponsibilityPMDesc'),'" & JSEncode(RS("ResponsibilityPMName")) & "');")+nl
		Response.Write("	top.setpk('txtResponsibilitySafetyPK','" & RS("ResponsibilitySafetyPK") & "');")+nl
		Response.Write("	myform.txtResponsibilitySafety.value = '" & JSEncode(RS("ResponsibilitySafetyID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtResponsibilitySafetyDesc'),'" & JSEncode(RS("ResponsibilitySafetyName")) & "');")+nl
		Response.Write("	top.setpk('txtServiceRepairPK','" & RS("ServiceRepairPK") & "');")+nl
		Response.Write("	myform.txtServiceRepair.value = '" & JSEncode(RS("ServiceRepairID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtServiceRepairDesc'),'" & JSEncode(RS("ServiceRepairName")) & "');")+nl
		Response.Write("	top.setpk('txtServicePMPK','" & RS("ServicePMPK") & "');")+nl
		Response.Write("	myform.txtServicePM.value = '" & JSEncode(RS("ServicePMID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtServicePMDesc'),'" & JSEncode(RS("ServicePMName")) & "');")+nl

		Call OutputUDFData(rs)

	End If


	'Risk Fields
	If CInt(section) = 3 or CInt(section) = 5 Then

		If NullCheck(RS("ClassIndustry")) = "" Then
			Response.Write("	myform.txtClassIndustry.value = 'F-G';")+nl
			Response.Write("	top.setdesc(myframe.document.getElementById('txtClassIndustryDesc'),'Facility - General');")+nl		
		Else
			Response.Write("	myform.txtClassIndustry.value = '" & JSEncode(RS("ClassIndustry")) & "';")+nl
			Response.Write("	top.setdesc(myframe.document.getElementById('txtClassIndustryDesc'),'" & JSEncode(RS("ClassIndustryDesc")) & "');")+nl
		End If
		If RS("RiskAssessmentRequired") Then
		Response.Write("	myform.txtRiskAssessmentRequired.checked = true;")+nl
		Else
		Response.Write("	myform.txtRiskAssessmentRequired.checked = false;")+nl
		End If
		Response.Write("	top.setpk('txtAssessedByPK','" & RS("AssessedByPK") & "');")+nl
		Response.Write("	myform.txtAssessedBy.value = '" & JSEncode(RS("AssessedByID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtAssessedByDesc'),'" & JSEncode(RS("AssessedByName")) & "');")+nl
		Response.Write("	myform.txtLastAssessed.value = '" & DateNullCheck(RS("LastAssessed")) & "';")+nl

		Dim RiskAssessmentGroup, RiskAssessmentGroupDesc
		If NullCheck(RS("RiskAssessmentGroup")) = "" Then
			RiskAssessmentGroup = "F"
			RiskAssessmentGroupDesc = "Facilities"
		Else
			RiskAssessmentGroup = NullCheck(RS("RiskAssessmentGroup"))
			RiskAssessmentGroupDesc = NullCheck(RS("RiskAssessmentGroupDesc"))		
		End If

		Response.Write("	myform.txtRiskAssessmentGroup.value = '" & JSEncode(RiskAssessmentGroup) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtRiskAssessmentGroupDesc'),'" & JSEncode(RiskAssessmentGroupDesc) & "');")+nl

		If RiskAssessmentGroup = "CE" Then
			Response.Write("	myform.AGCE1.value = '" & JSEncode(RS("RiskFactor1")) & "';")+nl
			'Response.Write("	myform.AGCE1ScoreH.value = '" & RS("RiskFactor1Score") & "';")+nl
			Response.Write("	myform.AGCE2.value = '" & JSEncode(RS("RiskFactor2")) & "';")+nl
			'Response.Write("	myform.AGCE2ScoreH.value = '" & RS("RiskFactor2Score") & "';")+nl
			Response.Write("	myform.AGCE3.value = '" & JSEncode(RS("RiskFactor3")) & "';")+nl
			'Response.Write("	myform.AGCE3ScoreH.value = '" & RS("RiskFactor3Score") & "';")+nl
			Response.Write("	myform.AGCE4.value = '" & JSEncode(RS("RiskFactor4")) & "';")+nl
			'Response.Write("	myform.AGCE4ScoreH.value = '" & RS("RiskFactor4Score") & "';")+nl
			Response.Write("	myform.AGCE5.value = '" & JSEncode(RS("RiskFactor5")) & "';")+nl
			'Response.Write("	myform.AGCE5ScoreH.value = '" & RS("RiskFactor5Score") & "';")+nl
			'Response.Write("	myform.AGCERiskScoreH.value = '" & RS("RiskScore") & "';")+nl
		ElseIf RiskAssessmentGroup = "F" Then
			Response.Write("	myform.AGF1.value = '" & JSEncode(RS("RiskFactor1")) & "';")+nl
			'Response.Write("	myform.AGF1ScoreH.value = '" & RS("RiskFactor1Score") & "';")+nl
			Response.Write("	myform.AGF2.value = '" & JSEncode(RS("RiskFactor2")) & "';")+nl
			'Response.Write("	myform.AGF2ScoreH.value = '" & RS("RiskFactor2Score") & "';")+nl
			Response.Write("	myform.AGF3.value = '" & JSEncode(RS("RiskFactor3")) & "';")+nl
			'Response.Write("	myform.AGF3ScoreH.value = '" & RS("RiskFactor3Score") & "';")+nl
			Response.Write("	myform.AGF4.value = '" & JSEncode(RS("RiskFactor4")) & "';")+nl
			'Response.Write("	myform.AGF4ScoreH.value = '" & RS("RiskFactor4Score") & "';")+nl
			If NullCheck(RS("RiskFactor5")) = "" Then
				Response.Write("	myform.AGF5.value = 'N/A';")+nl
				'Response.Write("	myform.AGF5ScoreH.value = '0';")+nl		
			Else			
				Response.Write("	myform.AGF5.value = '" & JSEncode(RS("RiskFactor5")) & "';")+nl
				'Response.Write("	myform.AGF5ScoreH.value = '" & RS("RiskFactor5Score") & "';")+nl		
			End If
			'Response.Write("	myform.AGFRiskScoreH.value = '" & RS("RiskScore") & "';")+nl
		End If
		
		'CB AUTOCALC FIX 1/10/2013
		'============================================================================================================================
		If CheckIfFieldExists(RS,"RiskLevelAutoCalc") Then
			If BitNullCheckTrue(RS("RiskLevelAutoCalc")) Then
				Response.Write("	if (myform.txtRiskLevelAutoCalc) {myform.txtRiskLevelAutoCalc.checked = true;}")+nl
			Else
				Response.Write("	if (myform.txtRiskLevelAutoCalc) {myform.txtRiskLevelAutoCalc.checked = false;}")+nl
			End If
			If BitNullCheckTrue(RS("PMRequiredAutoCalc")) Then
				Response.Write("	if (myform.txtPMRequiredAutoCalc) {myform.txtPMRequiredAutoCalc.checked = true;}")+nl
			Else
				Response.Write("	if (myform.txtPMRequiredAutoCalc) {myform.txtPMRequiredAutoCalc.checked = false;}")+nl
			End If
		End If
		'============================================================================================================================

		''''''''''''''''''''''''''''''''''''''' Start of changes by Remi
        Response.Write("	myform.RemiRisk1.value = '" & RS(JSEncode("UDFChar11")) & "';")+nl
        Response.Write("	myform.RemiRisk6.value = '" & RS(JSEncode("UDFChar12")) & "';")+nl
		Response.Write("	myform.RemiRisk2.value = '" & RS(JSEncode("UDFChar13")) & "';")+nl
        Response.Write("	myform.RemiRisk7.value = '" & RS(JSEncode("UDFChar14")) & "';")+nl
        Response.Write("	myform.RemiRisk3.value = '" & RS(JSEncode("UDFChar15")) & "';")+nl
        Response.Write("	myform.RemiRisk9.value = '" & RS(JSEncode("UDFChar16")) & "';")+nl
        Response.Write("	myform.RemiRisk10.value = '" & RS(JSEncode("UDFChar18")) & "';")+nl
        Response.Write("	myform.RemiRisk11.value = '" & RS(JSEncode("UDFChar19")) & "';")+nl
        Response.Write("	myform.RemiRisk12.value = '" & RS(JSEncode("UDFChar20")) & "';")+nl
        Response.Write("	myform.RemiRisk16.value = '" & RS(JSEncode("UDFChar31")) & "';")+nl
        Response.Write("	myform.RemiRisk13.value = '" & RS(JSEncode("UDFChar32")) & "';")+nl
        Response.Write("	myform.RemiRisk15.value = '" & RS(JSEncode("UDFChar33")) & "';")+nl
        Response.Write("	myform.RemiRisk17.value = '" & RS(JSEncode("UDFChar34")) & "';")+nl
        Response.Write("	myform.RemiRisk19.value = '" & RS(JSEncode("UDFChar35")) & "';")+nl
        Response.Write("	myform.RemiRisk20.value = '" & RS(JSEncode("UDFChar36")) & "';")+nl
        Response.Write("	myform.RemiRisk21.value = '" & RS(JSEncode("UDFChar37")) & "';")+nl  
        Response.Write("	myform.txtSetBy1.value = '" & RS(JSEncode("UDFChar38")) & "';")+nl
        Response.Write("	myform.txtSetBy2.value = '" & RS(JSEncode("UDFChar42")) & "';")+nl
        Response.Write("	myform.txtDate.value = '" & RS(JSEncode("UDFChar44")) & "';")+nl
        Response.Write("	myform.txtDate2.value = '" & RS(JSEncode("UDFChar48")) & "';")+nl
        Response.Write("	myform.txtDate3.value = '" & RS(JSEncode("UDFChar45")) & "';")+nl
        Response.Write("	myform.txtSetBy3.value = '" & RS(JSEncode("UDFChar46")) & "';")+nl
        Response.Write("	top.setdesc(myframe.document.getElementById('RemiRisk1Desc'),'" & JSENCODE(RS("UDFChar21")) & "');")+nl
        Response.Write("	top.setdesc(myframe.document.getElementById('RemiRisk6Desc'),'" & JSENCODE(RS("UDFChar22")) & "');")+nl
        Response.Write("	top.setdesc(myframe.document.getElementById('RemiRisk2Desc'),'" & JSENCODE(RS("UDFChar23")) & "');")+nl
        Response.Write("	top.setdesc(myframe.document.getElementById('RemiRisk7Desc'),'" & JSENCODE(RS("UDFChar24")) & "');")+nl
        Response.Write("	top.setdesc(myframe.document.getElementById('RemiRisk3Desc'),'" & JSENCODE(RS("UDFChar25")) & "');")+nl
        Response.Write("	top.setdesc(myframe.document.getElementById('RemiRisk9Desc'),'" & JSENCODE(RS("UDFChar26")) & "');")+nl
        Response.Write("	top.setdesc(myframe.document.getElementById('RemiRisk10Desc'),'" & JSENCODE(RS("UDFChar28")) & "');")+nl
        Response.Write("	top.setdesc(myframe.document.getElementById('RemiRisk11Desc'),'" & JSENCODE(RS("UDFChar29")) & "');")+nl
        Response.Write("	top.setdesc(myframe.document.getElementById('RemiRisk12Desc'),'" & JSENCODE(RS("UDFChar30")) & "');")+nl
        Response.Write("	top.setdesc(myframe.document.getElementById('RemiRisk16Desc'),'" & JSENCODE(RS("UDFChar41")) & "');")+nl
        Response.Write("	top.setdesc(myframe.document.getElementById('txtSetBy1Desc'),'" & JSENCODE(RS("UDFChar39")) & "');")+nl
        Response.Write("	top.setdesc(myframe.document.getElementById('txtSetBy2Desc'),'" & JSENCODE(RS("UDFChar43")) & "');")+nl
        Response.Write("	top.setdesc(myframe.document.getElementById('txtSetBy3Desc'),'" & JSENCODE(RS("UDFChar47")) & "');")+nl
        If RS("UDFChar17") Then
		Response.Write("	myform.txtMaintenanceHistory.checked = true;")+nl
        Else 
        Response.Write("	myform.txtMaintenanceHistory.checked = false;")+nl
        End If
        If RS("UDFChar27") Then
		Response.Write("	myform.txtAssistantDirector.checked = true;")+nl
        Else 
        Response.Write("	myform.txtAssistantDirector.checked = false;")+nl
        End If

        dim riskAH, riskRS         
		Set riskAH = New ADOHelper
		Set riskRS = riskAH.RunSQLReturnRS("select Comments, SpecificationName, Description from Specification WITH ( NOLOCK ) where Description='Risk_Tab';","")
						Do While Not riskRs.EOF
							
                            Select case JSEncode(riskRs("SpecificationName"))
                            case "Seriousness"
								Response.Write("	top.fraTopic.document.getElementById('lblRisk1_RiskText').getElementsByTagName('font')[0].innerText= '" & JSEncode(riskRs("Comments")) & "';")+nl
							
							case "Consequences"
								Response.Write("    top.fraTopic.document.getElementById('lblRisk2_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
							case "Score"
								Response.Write("    top.fraTopic.document.getElementById('lblRisk3_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
							case "Review" 
								Response.Write("    top.fraTopic.document.getElementById('lblRisk4_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
                                Response.Write("    top.fraTopic.document.getElementById('lblRisk18_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
                                Response.Write("    top.fraTopic.document.getElementById('lblRisk21_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
							case "Hospital"
								Response.Write("    top.fraTopic.document.getElementById('lblRisk6_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
							case "Equipment"
								Response.Write("    top.fraTopic.document.getElementById('lblRisk7_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
							case "Effectiveness" 
								Response.Write("    top.fraTopic.document.getElementById('lblRisk8_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
							case "Maintenance"
								Response.Write("    top.fraTopic.document.getElementById('lblRisk9_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
							case "Requirements"
								Response.Write("    top.fraTopic.document.getElementById('lblRisk10_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
							case "Modified"
								Response.Write("    top.fraTopic.document.getElementById('lblRisk11_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
							case "Timeliness"
								Response.Write("    top.fraTopic.document.getElementById('lblRisk12_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
							case "Failure"
								Response.Write("    top.fraTopic.document.getElementById('lblRisk13_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
							case "Evaluation" 
								Response.Write("    top.fraTopic.document.getElementById('lblRisk14_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
							case "Malfunction" 
								Response.Write("    top.fraTopic.document.getElementById('lblRisk15_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
							case "Remove" 
								Response.Write("    top.fraTopic.document.getElementById('lblRisk16_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
							case "Degraded"
								Response.Write("    top.fraTopic.document.getElementById('lblRisk17_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
							case "Why" 
								Response.Write("    top.fraTopic.document.getElementById('lblRisk19_RiskText').getElementsByTagName('font')[0].innerText='" & JSEncode(riskRs("Comments")) & "';")+nl
                            case else 
                                Response.Write("Error")
						    End Select
                        riskRS.MoveNext
					Loop
        Response.Write("        top.fraTopic.document.getElementById('lblRisk5_RiskText').getElementsByTagName('font')[0].innerText='Date';")+nl
		Set riskRS = Nothing
        'Response.Write("	top.setdesc(myframe.document.getElementById('txtsetby2'),'" & JSEncode(RS("txtSetBy2")) & "');")+nl
		'Response.Write("	myform.txtRiskLevel.value = '" & JSEncode(RS("RiskLevel")) & "';")+nl

		''''''''''''''''''''''''''''' End of changes by Remi


        If RS("PMRequired") Then
		Response.Write("	myform.txtPMRequired.checked = true;")+nl
		Else
		Response.Write("	myform.txtPMRequired.checked = false;")+nl
		End If
		Response.Write("	myframe.SetPMRequired(myform.txtPMRequired);")+nl
		If RS("PlanForImprovement") Then
		Response.Write("	myform.txtPlanForImprovement.checked = true;")+nl
		Else
		Response.Write("	myform.txtPlanForImprovement.checked = false;")+nl
		End If
		If RS("HIPPARelated") Then
		Response.Write("	myform.txtHIPPARelated.checked = true;")+nl
		Else
		Response.Write("	myform.txtHIPPARelated.checked = false;")+nl
		End If
		If RS("StatementOfConditions") Then
		Response.Write("	myform.txtStatementOfConditions.checked = true;")+nl
		Else
		Response.Write("	myform.txtStatementOfConditions.checked = false;")+nl
		End If
		If RS("StatementOfConditionsCompliant") Then
		Response.Write("	myform.txtStatementOfConditionsCompliant.checked = true;")+nl
		Else
		Response.Write("	myform.txtStatementOfConditionsCompliant.checked = false;")+nl
		End If
		
		Response.Write("	myframe.setriskfactors('" & JSEncode(RiskAssessmentGroup) & "');")+nl

	End If

	If CInt(section) = 4 or CInt(section) = 5 Then

		If Not RS("IsMeter") Then
		Response.Write("	myform.txtIsMeter.checked = false;")+nl
		Else
		Response.Write("	myform.txtIsMeter.checked = true;")+nl
		End If

		If IsNull(RS("Meter1RollDown")) or Not RS("Meter1RollDown") Then
		Response.Write("	myform.txtMeter1RollDown.checked = false;")+nl
		Response.Write("	if (myform.txtMeter1RollDownMethod) {")+nl
		Response.Write("	myform.txtMeter1RollDownMethod[0].checked = false;")+nl
		Response.Write("	myform.txtMeter1RollDownMethod[1].checked = false;")+nl
		Response.Write("}")+nl
		Else
		Response.Write("	myform.txtMeter1RollDown.checked = true;")+nl
		Response.Write("	if (myform.txtMeter1RollDownMethod) {")+nl
		If NullCheck(RS("Meter1RollDownMethod")) = "" or _
		   NullCheck(RS("Meter1RollDownMethod")) = "S" Then
		Response.Write("	myform.txtMeter1RollDownMethod[0].checked = true;")+nl
		Else
		Response.Write("	myform.txtMeter1RollDownMethod[1].checked = true;")+nl
		End If
		Response.Write("}")+nl
		End If

		If IsNull(RS("Meter2RollDown")) or Not RS("Meter2RollDown") Then
		Response.Write("	myform.txtMeter2RollDown.checked = false;")+nl
		Response.Write("	if (myform.txtMeter2RollDownMethod) {")+nl
		Response.Write("	myform.txtMeter2RollDownMethod[0].checked = false;")+nl
		Response.Write("	myform.txtMeter2RollDownMethod[1].checked = false;")+nl
		Response.Write("}")+nl
		Else
		Response.Write("	myform.txtMeter2RollDown.checked = true;")+nl
		Response.Write("	if (myform.txtMeter2RollDownMethod) {")+nl
		If NullCheck(RS("Meter2RollDownMethod")) = "" or _
		   NullCheck(RS("Meter2RollDownMethod")) = "S" Then
		Response.Write("	myform.txtMeter2RollDownMethod[0].checked = true;")+nl
		Else
		Response.Write("	myform.txtMeter2RollDownMethod[1].checked = true;")+nl
		End If
		Response.Write("}")+nl
		End If

		If IsNull(RS("Meter1TrackHistory")) or Not RS("Meter1TrackHistory") Then
		Response.Write("	myform.txtMeter1TrackHistory.checked = false;")+nl
		Else
		Response.Write("	myform.txtMeter1TrackHistory.checked = true;")+nl
		End If

		If IsNull(RS("Meter2TrackHistory")) or Not RS("Meter2TrackHistory") Then
		Response.Write("	myform.txtMeter2TrackHistory.checked = false;")+nl
		Else
		Response.Write("	myform.txtMeter2TrackHistory.checked = true;")+nl
		End If

		Response.Write("	myform.txtMeter1Reading.value = '" & RS("Meter1Reading") & "';")+nl

		Response.Write("	if (myform.txtMeter1ReadingLife) {myform.txtMeter1ReadingLife.value = '" & RS("Meter1ReadingLife") & "' };")+nl

		Response.Write("	myform.txtMeter1Units.value = '" & JSEncode(RS("Meter1Units")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtMeter1UnitsDesc'),'" & JSEncode(RS("Meter1UnitsDesc")) & "');")+nl
		Response.Write("	myform.txtMeter2Reading.value = '" & RS("Meter2Reading") & "';")+nl

		Response.Write("	if (myform.txtMeter2ReadingLife) {myform.txtMeter2ReadingLife.value = '" & RS("Meter2ReadingLife") & "' };")+nl

		Response.Write("	myform.txtMeter2Units.value = '" & JSEncode(RS("Meter2Units")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtMeter2UnitsDesc'),'" & JSEncode(RS("Meter2UnitsDesc")) & "');")+nl

	End If
   
End Sub

Function validate_master()

	'If Request.Form("txtSupervisor") = "" Then
	'	aok = False
	'	errorfield = "top.frames['fraTopic'].document.forms['mcform'].txtSupervisor"
	'	errortabinfo = "wo_details"
	'	returnmessage = "Supervisor can NOT be left blank (Server Side Validation)"
	'	returnclass = "errormessage"
	'	Call DoResponse(lastaction,"",True,True,False,keyvalue)
	'End If

	'If NewRecord and Request.Form("ItemID") Then
	'    Dim rs,sql
	'    sql = "SELECT ASVALID = CASE WHEN EXISTS (SELECT 1 FROM Asset WITH (NOLOCK) WHERE AssetPK = '" & ParentPK & "' AND RotatingLocationPK Is Not Null) THEN 1 ELSE 0 END "
	'    Set rs = db.RunSQLReturnRS(sql)
	'End If

End Function

Function validate_sp1(thing,theindex)

End Function

Function validate_no1(thing,theindex)

End Function

Function validate_am1(thing,theindex)

End Function

Function validate_ap1(thing,theindex)

End Function

Function validate_ao1(thing,theindex)

End Function

Function validate_al1(thing,theindex)

End Function

Function validate_ar1(thing,theindex)

End Function

Function validate_at1(thing,theindex)

End Function

Sub db_master(db)

	Dim rs,OutArray

	If ErrorHandler Then
		On Error Resume Next
	End If

	If Not db.dok Then
		Exit Sub
	End If

	'Check for Duplicate Asset ID
	If newrecord or duprecord Then
		Set rs = db.RunSQLReturnRS("SELECT AssetPK FROM Asset WITH (NOLOCK) WHERE AssetID = '" & Trim(Request.Form("txtAsset")) & "' ","")
	Else
		Set rs = db.RunSQLReturnRS("SELECT AssetPK FROM Asset WITH (NOLOCK) WHERE AssetID = '" & Trim(Request.Form("txtAsset")) & "' AND NOT AssetPK = " & NullCheck(keyvalue),"")
	End If
	If Not db.dok Then
		Exit Sub
	End If
	If Not rs.Eof Then
		db.dok = False
		db.isduplicate = True
		Exit Sub
	End If

	Set rs = db.RunSqlReturnRS("Select GETDATE() AS ServerDate","")
	If Not db.dok Then
		Exit Sub
	End If
	ServerDate = rs("ServerDate")

	If newrecord or duprecord Then
		Set rs = db.RunSQLReturnRS_RW("SELECT TOP 0 * FROM Asset","")
		If Not db.dok Then
			Exit Sub
		End If
		rs.AddNew
	Else
		Set rs = db.RunSQLReturnRS_RW("SELECT * FROM Asset WHERE AssetPK=" & NullCheck(keyvalue),"")
		If Not db.dok Then
			Exit Sub
		End If
	End If

	If rs.eof Then
		' Looks like another user deleted the record first
		db.dok = False
		db.derror = "Another user has deleted this Asset while you were working with it. If you feel this is not the case, you can try again, otherwise please click the CANCEL button to cancel out of this Asset."
		Exit Sub
	End If

	Call DemoNoEditMsg(db,rs)

	Dim EditConflict,ConflictIsStatus,ConflictIsAuthStatus,ConflictIsAssign,rvd,rva

	rvd = Trim(rs("RowVersionDate"))
	rva = Trim(rs("RowVersionAction"))
	If Not rvd = NullCheck(Request.Form("txtRowVersionDate")) Then
		' Looks like another user made a change and this data is 'stale'.
		' Let's see what was changed and try not to bother the user
		EditConflict = True
		If DateDiff("n",NullCheck(Request.Form("txtRowVersionDate")),rvd) < 15 Then
			If False Then
				'ConflictIsStatus = True
			Else
				db.dok = False
				db.derror = "Another user has made modifications to this Asset while you were working with it. Since you could potentially overwrite changes the other user made, you will need to Cancel your changes and start again. Please click the CANCEL button to cancel your changes."
				Exit Sub
			End If
		Else
			db.dok = False
			db.derror = "Another user has made modifications to this Asset while you were working with it. Since you could potentially overwrite changes the other user made, you will need to Cancel your changes and start again. Please click the CANCEL button to cancel your changes."
			Exit Sub
		End If
	End If

	'Check for Duplicate Serial #
	Dim rsSerial
	If Not Trim(Request.Form("txtSerial")) = "" _
	   And Not Trim(NullCheck(rs("Serial"))) = Trim(Request.Form("txtSerial")) Then
		If newrecord or duprecord Then
			Set rsSerial = db.RunSQLReturnRS("SELECT AssetPK FROM Asset WITH (NOLOCK) WHERE Serial = '" & Trim(Request.Form("txtSerial")) & "' ","")
		Else
			Set rsSerial = db.RunSQLReturnRS("SELECT AssetPK FROM Asset WITH (NOLOCK) WHERE Serial = '" & Trim(Request.Form("txtSerial")) & "' AND NOT AssetPK = " & NullCheck(keyvalue),"")
		End If
		If Not db.dok Then
			Exit Sub
		End If
		If Not rsSerial.Eof Then
			db.warn = True
			db.warntext = Trim(db.warntext + Trim(" The Serial # that was entered is already in use by another Asset."))
		End If
		Call CloseObj(rsSerial)
	End If

	' CUSTOMIZED
	'--------------------------------------------------------------------------------------------------------------
	If Not Trim(NullCheck(rs("AssetID"))) = Trim(Request.Form("txtAsset")) Then
		IsIDChange = True
		' If the Asset Name changed - we must rebuild the Tree Records
		AssetUpdate = True
	End If
	rs("AssetID") = Trim(Mid(Request.Form("txtAsset").Item,1,100))	' Nullable: No Type: nvarchar

	If InStr(rs("AssetID"),"-CLONE") > 0 Then
		rs("IsClone") = 1
	Else
		rs("IsClone") = 0
	End If

	If Not IsEmpty(Request.Form("txtAssetName").Item) Then
		If Not Trim(NullCheck(rs("AssetName"))) = Trim(Request.Form("txtAssetName")) Then
			IsNameChange = True
			' If the Asset Name changed - we must rebuild the Tree Records
			AssetUpdate = True
		End If

		rs("AssetName") = Trim(Mid(Request.Form("txtAssetName").Item,1,150))	' Nullable: YES Type: nvarchar
	End If
	'--------------------------------------------------------------------------------------------------------------

	If Not IsEmpty(Request.Form("txtFormName").Item) Then
		rs("FormName") = Trim(Mid(Request.Form("txtFormName").Item,1,50))	' Nullable: YES Type: nvarchar
	End If

	rs("RequesterCanView") = Not Request.Form("txtRequesterCanView").Item = ""	' Nullable: No Type: bit

	OldFieldValue = rs("IsUp")
	rs("IsUp") = Not Request.Form("txtIsUp").Item = ""	' Nullable: No Type: bit

	' CUSTOMIZED
	'--------------------------------------------------------------------------------------------------------------
	If Not IsEmpty(Request.Form("txtType").Item) Then
		If Not NullCheck(rs("Type")) = Trim(Request.Form("txtType")) Then
			IsTypeChange = True
		End If
		rs("Type") = Trim(Mid(Request.Form("txtType").Item,1,25))	' Nullable: YES Type: nvarchar
		If Trim(Mid(Request.Form("txtType").Item,1,25)) = "A" or _
		   Trim(Mid(Request.Form("txtType").Item,1,25)) = "AL" Then
			rs("IsLocation") = False ' Nullable: No Type: bit
		Else
			rs("IsLocation") = True	' Nullable: No Type: bit
		End If
		If Trim(Mid(Request.Form("txtType").Item,1,25)) = "AL" Then
			rs("IsLinear") = True ' Nullable: No Type: bit
		Else
			rs("IsLinear") = False	' Nullable: No Type: bit
		End If
	End If
	'--------------------------------------------------------------------------------------------------------------
	IsLocation = rs("IsLocation")
	If Not IsEmpty(Request.Form("txtTypeDescH").Item) Then
		rs("TypeDesc") = Trim(Mid(Request.Form("txtTypeDescH").Item,1,50))	' Nullable: YES Type: nvarchar
	End If
	If Not IsEmpty(Request.Form("txtModel").Item) Then
		If Len(Request.Form("txtModel").Item) > 0 Then
			rs("Model") = Trim(Mid(Request.Form("txtModel").Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("Model") = Null
		End If
	End If

	If Not IsEmpty(Request.Form("txtModelNumber").Item) Then
		If Len(Request.Form("txtModelNumber").Item) > 0 Then
			rs("ModelNumber") = Trim(Mid(Request.Form("txtModelNumber").Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("ModelNumber") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtModelNumberMFG").Item) Then
		If Len(Request.Form("txtModelNumberMFG").Item) > 0 Then
			rs("ModelNumberMFG") = Trim(Mid(Request.Form("txtModelNumberMFG").Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("ModelNumberMFG") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtModelLine").Item) Then
		If Len(Request.Form("txtModelLine").Item) > 0 Then
			rs("ModelLine") = Trim(Mid(Request.Form("txtModelLine").Item,1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtModelLineDescH").Item) Then
				rs("ModelLineDesc") = Trim(Mid(Request.Form("txtModelLineDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("ModelLine") = Null
			rs("ModelLineDesc") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtModelSeries").Item) Then
		If Len(Request.Form("txtModelSeries").Item) > 0 Then
			rs("ModelSeries") = Trim(Mid(Request.Form("txtModelSeries").Item,1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtModelSeriesDescH").Item) Then
				rs("ModelSeriesDesc") = Trim(Mid(Request.Form("txtModelSeriesDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("ModelSeries") = Null
			rs("ModelSeriesDesc") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtSystemPlatform").Item) Then
		If Len(Request.Form("txtSystemPlatform").Item) > 0 Then
			rs("SystemPlatform") = Trim(Mid(Request.Form("txtSystemPlatform").Item,1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtSystemPlatformDescH").Item) Then
				rs("SystemPlatformDesc") = Trim(Mid(Request.Form("txtSystemPlatformDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("SystemPlatform") = Null
			rs("SystemPlatformDesc") = Null
		End If
	End If

	If Not IsEmpty(Request.Form("txtSerial").Item) Then
		If Len(Request.Form("txtSerial").Item) > 0 Then
			rs("Serial") = Trim(Mid(Request.Form("txtSerial").Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("Serial") = Null
		End If
	End If
	' CUSTOMIZED
	'---------------------------------------------------------------------------------------------------------------
	If Not Trim(NullCheck(rs("ClassificationID"))) = Trim(Request.Form("txtClassification")) Then
		IsClassificationChange = True
		' If the Asset Name changed - we must rebuild the Tree Records
		AssetUpdate = True
	End If
	'---------------------------------------------------------------------------------------------------------------
	If Len(Trim(Request.Form("txtClassificationPK").Item)) > 0 Then
		rs("ClassificationPK") = Request.Form("txtClassificationPK").Item	' Nullable: No Type: int
		If Not IsEmpty(Request.Form("txtClassification").Item) Then
			rs("ClassificationID") = Trim(Mid(Request.Form("txtClassification").Item,1,25))	' Nullable: YES Type: nvarchar
		End If
		If Not IsEmpty(Request.Form("txtClassificationDescH").Item) Then
			rs("ClassificationName") = Trim(Mid(Request.Form("txtClassificationDescH").Item,1,50))	' Nullable: YES Type: nvarchar
		End If
	Else
		rs("ClassificationPK") = Null
		rs("ClassificationID") = Null
		rs("ClassificationName") = Null
	End If
	If Not IsEmpty(Request.Form("txtSystem").Item) Then
		If Len(Request.Form("txtSystem").Item) > 0 Then
			rs("System") = Trim(Mid(Request.Form("txtSystem").Item,1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtSystemDescH").Item) Then
				rs("SystemDesc") = Trim(Mid(Request.Form("txtSystemDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("System") = Null
			rs("SystemDesc") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtPriority").Item) Then
		If Len(Request.Form("txtPriority").Item) > 0 Then
			rs("Priority") = Trim(Mid(Request.Form("txtPriority").Item,1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtPriorityDescH").Item) Then
				rs("PriorityDesc") = Trim(Mid(Request.Form("txtPriorityDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("Priority") = Null
			rs("PriorityDesc") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtStatus").Item) Then
		If Len(Request.Form("txtStatus").Item) > 0 Then
			rs("Status") = Trim(Mid(Request.Form("txtStatus").Item,1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtStatusDescH").Item) Then
				rs("StatusDesc") = Trim(Mid(Request.Form("txtStatusDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("Status") = Null
			rs("StatusDesc") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtVicinity").Item) Then
		If Len(Request.Form("txtVicinity").Item) > 0 Then
			rs("Vicinity") = Trim(Mid(Request.Form("txtVicinity").Item,1,100))	' Nullable: YES Type: nvarchar
		Else
			rs("Vicinity") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtAccountPK").Item) Then
		If Len(Trim(Request.Form("txtAccountPK").Item)) > 0 Then
			rs("AccountPK") = Request.Form("txtAccountPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtAccount").Item) Then
				rs("AccountID") = Trim(Mid(Request.Form("txtAccount").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtAccountDescH").Item) Then
				rs("AccountName") = Trim(Mid(Request.Form("txtAccountDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("AccountPK") = Null
			rs("AccountID") = Null
			rs("AccountName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtTenantPK").Item) Then
		If Len(Trim(Request.Form("txtTenantPK").Item)) > 0 Then
			rs("TenantPK") = Request.Form("txtTenantPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtTenant").Item) Then
				rs("TenantID") = Trim(Mid(Request.Form("txtTenant").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtTenantDescH").Item) Then
				rs("TenantName") = Trim(Mid(Request.Form("txtTenantDescH").Item,1,100))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("TenantPK") = Null
			rs("TenantID") = Null
			rs("TenantName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtDepartmentPK").Item) Then
		If Len(Trim(Request.Form("txtDepartmentPK").Item)) > 0 Then
			rs("DepartmentPK") = Request.Form("txtDepartmentPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtDepartment").Item) Then
				rs("DepartmentID") = Trim(Mid(Request.Form("txtDepartment").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtDepartmentDescH").Item) Then
				rs("DepartmentName") = Trim(Mid(Request.Form("txtDepartmentDescH").Item,1,100))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("DepartmentPK") = Null
			rs("DepartmentID") = Null
			rs("DepartmentName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtRotatingPartPK").Item) Then
		If Len(Trim(Request.Form("txtRotatingPartPK").Item)) > 0 Then
			rs("RotatingPartPK") = Request.Form("txtRotatingPartPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtRotatingPart").Item) Then
				rs("RotatingPartID") = Trim(Mid(Request.Form("txtRotatingPart").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtRotatingPartDescH").Item) Then
				rs("RotatingPartName") = Trim(Mid(Request.Form("txtRotatingPartDescH").Item,1,500))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("RotatingPartPK") = Null
			rs("RotatingPartID") = Null
			rs("RotatingPartName") = Null
		End If
	End If

	'@$CUSTOMISED - Enhanced Tool Module
	If Not IsEmpty(Request.Form("txtMaintainableToolPK").Item) Then
		If Len(Trim(Request.Form("txtMaintainableToolPK").Item)) > 0 Then
			rs("MaintainableToolPK") = Request.Form("txtMaintainableToolPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtMaintainableTool").Item) Then
				rs("MaintainableToolID") = Trim(Mid(Request.Form("txtMaintainableTool").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtMaintainableToolDescH").Item) Then
				rs("MaintainableToolName") = Trim(Mid(Request.Form("txtMaintainableToolDescH").Item,1,100))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("MaintainableToolPK") = Null
			rs("MaintainableToolID") = Null
			rs("MaintainableToolName") = Null
		End If
	End If
	'@$END

	' CUSTOMIZED
	'--------------------------------------------------------------------------------------------------------------
	If Not IsEmpty(Request.Form("txtRepairCenterPK").Item) Then

		If Len(Trim(Request.Form("txtRepairCenterPK").Item)) > 0 Then

			' Check to see if the Repair Center changed
			' If it did - then we need to plug this in the AssetMove table

			If (Not NewRecord) And (Not Trim(rs("RepairCenterPK")) = Trim(Request.Form("txtRepairCenterPK").Item)) Then
				Call db.RunSP("MC_MoveAssetRC",Array(Array("@AssetPK", adInteger, adParamInput, 4, keyvalue),Array("@FromPK", adInteger, adParamInput, 4, rs("RepairCenterPK")),Array("@ToPK", adInteger, adParamInput, 4, Trim(Request.Form("txtRepairCenterPK").Item)),Array("@UserPK", adInteger, adParamInput, 4, Request.Form("txtRowVersionUserPK").Item),Array("@Initials", MC_ADVARCHAR, adParamInput, 5, Trim(Mid(Request.Form("txtRowVersionInitials").Item,1,5)))),"")
			End If

			rs("RepairCenterPK") = Request.Form("txtRepairCenterPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtRepairCenter").Item) Then
				rs("RepairCenterID") = Trim(Mid(Request.Form("txtRepairCenter").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtRepairCenterDescH").Item) Then
				rs("RepairCenterName") = Trim(Mid(Request.Form("txtRepairCenterDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("RepairCenterPK") = Null
			rs("RepairCenterID") = Null
			rs("RepairCenterName") = Null
		End If
	End If
	'--------------------------------------------------------------------------------------------------------------

	If Not IsEmpty(Request.Form("txtShopPK").Item) Then
		If Len(Trim(Request.Form("txtShopPK").Item)) > 0 Then
			rs("ShopPK") = Request.Form("txtShopPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtShop").Item) Then
				rs("ShopID") = Trim(Mid(Request.Form("txtShop").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtShopDescH").Item) Then
				rs("ShopName") = Trim(Mid(Request.Form("txtShopDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("ShopPK") = Null
			rs("ShopID") = Null
			rs("ShopName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtVendorPK").Item) Then
		If Len(Trim(Request.Form("txtVendorPK").Item)) > 0 Then
			rs("VendorPK") = Request.Form("txtVendorPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtVendor").Item) Then
				rs("VendorID") = Trim(Mid(Request.Form("txtVendor").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtVendorDescH").Item) Then
				rs("VendorName") = Trim(Mid(Request.Form("txtVendorDescH").Item,1,100))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("VendorPK") = Null
			rs("VendorID") = Null
			rs("VendorName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtManufacturerPK").Item) Then
		If Len(Trim(Request.Form("txtManufacturerPK").Item)) > 0 Then
			rs("ManufacturerPK") = Request.Form("txtManufacturerPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtManufacturer").Item) Then
				rs("ManufacturerID") = Trim(Mid(Request.Form("txtManufacturer").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtManufacturerDescH").Item) Then
				rs("ManufacturerName") = Trim(Mid(Request.Form("txtManufacturerDescH").Item,1,100))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("ManufacturerPK") = Null
			rs("ManufacturerID") = Null
			rs("ManufacturerName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtWarrantyExpire").Item) Then
		If Not Request.Form("txtWarrantyExpire").Item = "" Then
			rs("WarrantyExpire") = SQLdatetimeADO(Request.Form("txtWarrantyExpire").Item)	' Nullable: YES Type: datetime
		Else
			rs("WarrantyExpire") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtAddress").Item) Then
		If Len(Request.Form("txtAddress").Item) > 0 Then
			rs("Address") = Trim(Mid(Request.Form("txtAddress").Item,1,80))	' Nullable: YES Type: nvarchar
		Else
			rs("Address") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtCity").Item) Then
		If Len(Request.Form("txtCity").Item) > 0 Then
			rs("City") = Trim(Mid(Request.Form("txtCity").Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("City") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtState").Item) Then
		If Len(Request.Form("txtState").Item) > 0 Then
			rs("State") = Trim(Mid(Request.Form("txtState").Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("State") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtZip").Item) Then
		If Len(Request.Form("txtZip").Item) > 0 Then
			rs("Zip") = Trim(Mid(Request.Form("txtZip").Item,1,15))	' Nullable: YES Type: nvarchar
		Else
			rs("Zip") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtCountry").Item) Then
		If Len(Request.Form("txtCountry").Item) > 0 Then
			rs("Country") = Trim(Mid(Request.Form("txtCountry").Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("Country") = Null
		End If
	End If
	rs("DisplayMapOnWO") = Not Request.Form("txtDisplayMapOnWO").Item = ""	' Nullable: No Type: bit

	If Not IsEmpty(Request.Form("txtPMCycleStartBy").Item) Then
		If Len(Request.Form("txtPMCycleStartBy").Item) > 0 Then
			rs("PMCycleStartBy") = Trim(Mid(Request.Form("txtPMCycleStartBy").Item,1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtPMCycleStartByDescH").Item) Then
				rs("PMCycleStartByDesc") = Trim(Mid(Request.Form("txtPMCycleStartByDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If

			If Trim(UCase(rs("PMCycleStartBy"))) = "AS" Then
				If Not IsEmpty(Request.Form("txtPMCycleStartDate").Item) Then
					If Not Request.Form("txtPMCycleStartDate").Item = "" Then
						rs("PMCycleStartDate") = SQLdatetimeADO(Request.Form("txtPMCycleStartDate").Item)	' Nullable: YES Type: datetime
					End If
				End If
			End If

		Else
			rs("PMCycleStartBy") = Null
			rs("PMCycleStartByDesc") = Null
		End If
	End If

	If Not IsEmpty(Request.Form("txtPurchaseType").Item) Then
		If Len(Request.Form("txtPurchaseType").Item) > 0 Then
			rs("PurchaseType") = Trim(Mid(Request.Form("txtPurchaseType").Item,1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtPurchaseTypeDescH").Item) Then
				rs("PurchaseTypeDesc") = Trim(Mid(Request.Form("txtPurchaseTypeDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("PurchaseType") = Null
			rs("PurchaseTypeDesc") = Null
		End If
	End If

	If Not IsEmpty(Request.Form("txtPurchasedDate").Item) Then
		If Not Request.Form("txtPurchasedDate").Item = "" Then
			rs("PurchasedDate") = SQLdatetimeADO(Request.Form("txtPurchasedDate").Item)	' Nullable: YES Type: datetime
		Else
			rs("PurchasedDate") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtPurchaseOrder").Item) Then
		If Len(Request.Form("txtPurchaseOrder").Item) > 0 Then
			rs("PurchaseOrder") = Trim(Mid(Request.Form("txtPurchaseOrder").Item,1,25))	' Nullable: YES Type: nvarchar
		Else
			rs("PurchaseOrder") = Null
		End If
	End If
	If Len(Trim(Request.Form("txtPurchaseCost").Item)) > 0 Then
		rs("PurchaseCost") = FixInternationalNumber(Request.Form("txtPurchaseCost").Item)	' Nullable: No Type: money
	End If
	If Not IsEmpty(Request.Form("txtInstallDate").Item) Then
		If Not Request.Form("txtInstallDate").Item = "" Then
			rs("InstallDate") = SQLdatetimeADO(Request.Form("txtInstallDate").Item)	' Nullable: YES Type: datetime
		Else
			rs("InstallDate") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtReplaceDate").Item) Then
		If Not Request.Form("txtReplaceDate").Item = "" Then
			rs("ReplaceDate") = SQLdatetimeADO(Request.Form("txtReplaceDate").Item)	' Nullable: YES Type: datetime
		Else
			rs("ReplaceDate") = Null
		End If
	End If
	If Len(Trim(Request.Form("txtReplacementCost").Item)) > 0 Then
		rs("ReplacementCost") = FixInternationalNumber(Request.Form("txtReplacementCost").Item)	' Nullable: No Type: money
	End If
	If Not IsEmpty(Request.Form("txtDisposalDate").Item) Then
		If Not Request.Form("txtDisposalDate").Item = "" Then
			rs("DisposalDate") = SQLdatetimeADO(Request.Form("txtDisposalDate").Item)	' Nullable: YES Type: datetime
		Else
			rs("DisposalDate") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtLicenseNumber").Item) Then
		If Len(Request.Form("txtLicenseNumber").Item) > 0 Then
			rs("LicenseNumber") = Trim(Mid(Request.Form("txtLicenseNumber").Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("LicenseNumber") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtContactPK").Item) Then
		If Len(Trim(Request.Form("txtContactPK").Item)) > 0 Then
			rs("ContactPK") = Request.Form("txtContactPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtContact").Item) Then
				rs("ContactID") = Trim(Mid(Request.Form("txtContact").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtContactDescH").Item) Then
				rs("ContactName") = Trim(Mid(Request.Form("txtContactDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("ContactPK") = Null
			rs("ContactID") = Null
			rs("ContactName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtOperatorPK").Item) Then
		If Len(Trim(Request.Form("txtOperatorPK").Item)) > 0 Then
			rs("OperatorPK") = Request.Form("txtOperatorPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtOperator").Item) Then
				rs("OperatorID") = Trim(Mid(Request.Form("txtOperator").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtOperatorDescH").Item) Then
				rs("OperatorName") = Trim(Mid(Request.Form("txtOperatorDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("OperatorPK") = Null
			rs("OperatorID") = Null
			rs("OperatorName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtInsuranceCarrier").Item) Then
		If Len(Request.Form("txtInsuranceCarrier").Item) > 0 Then
			rs("InsuranceCarrier") = Trim(Mid(Request.Form("txtInsuranceCarrier").Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("InsuranceCarrier") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtInsurancePolicy").Item) Then
		If Len(Request.Form("txtInsurancePolicy").Item) > 0 Then
			rs("InsurancePolicy") = Trim(Mid(Request.Form("txtInsurancePolicy").Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("InsurancePolicy") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtLeaseNumber").Item) Then
		If Len(Request.Form("txtLeaseNumber").Item) > 0 Then
			rs("LeaseNumber") = Trim(Mid(Request.Form("txtLeaseNumber").Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("LeaseNumber") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtRegistrationDate").Item) Then
		If Not Request.Form("txtRegistrationDate").Item = "" Then
			rs("RegistrationDate") = SQLdatetimeADO(Request.Form("txtRegistrationDate").Item)	' Nullable: YES Type: datetime
		Else
			rs("RegistrationDate") = Null
		End If
	End If

	If Not IsEmpty(Request.Form("txtTechnology")) Then
		If Len(Request.Form("txtTechnology").Item) > 0 Then
			rs("Technology") = Trim(Mid(Request.Form("txtTechnology"),1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtTechnologyDescH")) Then
				rs("TechnologyDesc") = Trim(Mid(Request.Form("txtTechnologyDescH"),1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("Technology") = Null
			rs("TechnologyDesc") = Null
		End If
	End If

	rs("IsPredictive") = Not Request.Form("txtIsPredictive").Item = ""	' Nullable: No Type: bit
	rs("DrawingUpdatesNeeded") = Not Request.Form("txtDrawingUpdatesNeeded").Item = ""	' Nullable: No Type: bit

	rs("IsMeter") = Not Request.Form("txtIsMeter").Item = ""	' Nullable: No Type: bit

	rs("Meter1RollDown") = Not Request.Form("txtMeter1RollDown").Item = ""	' Nullable: Yes Type: bit
	rs("Meter2RollDown") = Not Request.Form("txtMeter2RollDown").Item = ""	' Nullable: Yes Type: bit

	If Not IsEmpty(Request.Form("txtMeter1RollDownMethod").Item) Then
		If Len(Request.Form("txtMeter1RollDownMethod").Item) > 0 Then
			rs("Meter1RollDownMethod") = Request.Form("txtMeter1RollDownMethod").Item  ' Nullable: Yes Type: char(1)
		End If
	End If

	If Not IsEmpty(Request.Form("txtMeter2RollDownMethod").Item) Then
		If Len(Request.Form("txtMeter2RollDownMethod").Item) > 0 Then
			rs("Meter2RollDownMethod") = Request.Form("txtMeter2RollDownMethod").Item  ' Nullable: Yes Type: char(1)
		End If
	End If

	rs("Meter1TrackHistory") = Not Request.Form("txtMeter1TrackHistory").Item = ""	' Nullable: No Type: bit
	rs("Meter2TrackHistory") = Not Request.Form("txtMeter2TrackHistory").Item = ""	' Nullable: No Type: bit

	If Len(Trim(Request.Form("txtMeter1Reading").Item)) > 0 Then
		rs("Meter1Reading") = FixInternationalNumber(Request.Form("txtMeter1Reading").Item)	' Nullable: No Type: real
	End If

	If Len(Trim(Request.Form("txtMeter1ReadingLife").Item)) > 0 Then
		rs("Meter1ReadingLife") = FixInternationalNumber(Request.Form("txtMeter1ReadingLife").Item)	' Nullable: No Type: real
	End If

	If Not IsEmpty(Request.Form("txtMeter1Units").Item) Then
		If Len(Request.Form("txtMeter1Units").Item) > 0 Then
			rs("Meter1Units") = Trim(Mid(Request.Form("txtMeter1Units").Item,1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtMeter1UnitsDescH").Item) Then
				rs("Meter1UnitsDesc") = Trim(Mid(Request.Form("txtMeter1UnitsDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("Meter1Units") = Null
			rs("Meter1UnitsDesc") = Null
		End If
	End If
	If Len(Trim(Request.Form("txtMeter2Reading").Item)) > 0 Then
		rs("Meter2Reading") = FixInternationalNumber(Request.Form("txtMeter2Reading").Item)	' Nullable: No Type: real
	End If

	If Len(Trim(Request.Form("txtMeter2ReadingLife").Item)) > 0 Then
		rs("Meter2ReadingLife") = FixInternationalNumber(Request.Form("txtMeter2ReadingLife").Item)	' Nullable: No Type: real
	End If

	If Not IsEmpty(Request.Form("txtMeter2Units").Item) Then
		If Len(Request.Form("txtMeter2Units").Item) > 0 Then
			rs("Meter2Units") = Trim(Mid(Request.Form("txtMeter2Units").Item,1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtMeter2UnitsDescH").Item) Then
				rs("Meter2UnitsDesc") = Trim(Mid(Request.Form("txtMeter2UnitsDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("Meter2Units") = Null
			rs("Meter2UnitsDesc") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtLastMaintained").Item) Then
		If Not Request.Form("txtLastMaintained").Item = "" Then
			rs("LastMaintained") = SQLdatetimeADO(Request.Form("txtLastMaintained").Item)	' Nullable: YES Type: datetime
		Else
			rs("LastMaintained") = Null
		End If
	End If
	rs("InstructionsToWO") = Not Request.Form("txtInstructionsToWO").Item = ""	' Nullable: No Type: bit

	' CUSTOMIZED
	'--------------------------------------------------------------------------------------------------------------
	If Not IsEmpty(Request.Form("txtInstructions").Item) Then
		If Len(Request.Form("txtInstructions").Item) > 0 Then
			rs("Instructions") = Trim(Mid(Replace(Request.Form("txtInstructions").Item,Chr(13)+Chr(10),"%0D%0A"),1,2000))	' Nullable: YES Type: nvarchar
		Else
			rs("Instructions") = Null
		End If
	End If
	'--------------------------------------------------------------------------------------------------------------

	If Not IsEmpty(Request.Form("txtIcon").Item) Then
		If Len(Request.Form("txtIcon").Item) > 0 Then
			' We add the ! because "" = Everything!
			If Not Trim(Request.Form("txtIcon").Item)&"!" = rs("Icon")&"!" Then
				IsIconChange = True
			End If
			rs("Icon") = Trim(Mid(Request.Form("txtIcon").Item,1,200))	' Nullable: YES Type: nvarchar
		Else
			rs("Icon") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtPhoto").Item) Then
		If Len(Request.Form("txtPhoto").Item) > 0 Then
			rs("Photo") = Trim(Mid(Request.Form("txtPhoto").Item,1,200))	' Nullable: YES Type: nvarchar
		Else
			rs("Photo") = Null
		End If
	End If
' Risk Fields
	'-----------------------------------------------------------------------------------------------------------------

	If Not IsEmpty(Request.Form("txtClassIndustry")) Then
		rs("ClassIndustry") = Trim(Mid(Request.Form("txtClassIndustry"),1,25))	' Nullable: YES Type: varchar
	End If
	If Not IsEmpty(Request.Form("txtClassIndustryDescH")) Then
		rs("ClassIndustryDesc") = Trim(Mid(Request.Form("txtClassIndustryDescH"),1,50))	' Nullable: YES Type: varchar
	End If
	rs("RiskAssessmentRequired") = Not Request.Form("txtRiskAssessmentRequired") = ""	' Nullable: YES Type: bit
	If Not IsEmpty(Request.Form("txtAssessedByPK").Item) Then
		If Len(Trim(Request.Form("txtAssessedByPK").Item)) > 0 Then
			rs("AssessedByPK") = Request.Form("txtAssessedByPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtAssessedBy").Item) Then
				rs("AssessedByID") = Trim(Mid(Request.Form("txtAssessedBy").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtAssessedByDescH").Item) Then
				rs("AssessedByName") = Trim(Mid(Request.Form("txtAssessedByDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("AssessedByPK") = Null
			rs("AssessedByID") = Null
			rs("AssessedByName") = Null
		End If
	End If	
	If Not IsEmpty(Request.Form("txtLastAssessed")) Then
		If Not Request.Form("txtLastAssessed") = "" Then
			rs("LastAssessed") = SQLdatetimeADO(Request.Form("txtLastAssessed"))	' Nullable: YES Type: datetime
		End If
	End If
	If Not IsEmpty(Request.Form("txtRiskAssessmentGroup")) Then
	
		rs("RiskAssessmentGroup") = Trim(Mid(Request.Form("txtRiskAssessmentGroup"),1,25))	' Nullable: YES Type: varchar
		
		If Not IsEmpty(Request.Form("txtRiskAssessmentGroupDescH")) Then
			rs("RiskAssessmentGroupDesc") = Trim(Mid(Request.Form("txtRiskAssessmentGroupDescH"),1,50))	' Nullable: YES Type: varchar
		End If

		If Request.Form("txtRiskAssessmentGroup") = "CE" Then
				
			If Not IsEmpty(Request.Form("AGCE1")) Then
				rs("RiskFactor1") = Trim(Mid(Request.Form("AGCE1"),1,100))	' Nullable: YES Type: varchar
			End If
			If Not IsEmpty(Request.Form("AGCE1ScoreH")) Then
				If IsNumeric(Request.Form("AGCE1ScoreH")) and Len(Trim(Request.Form("AGCE1ScoreH"))) > 0 Then
					rs("RiskFactor1Score") = Request.Form("AGCE1ScoreH")	' Nullable: YES Type: smallint
				End If
			End If
			If Not IsEmpty(Request.Form("AGCE2")) Then
				rs("RiskFactor2") = Trim(Mid(Request.Form("AGCE2"),1,100))	' Nullable: YES Type: varchar
			End If
			If Not IsEmpty(Request.Form("AGCE2ScoreH")) Then
				If IsNumeric(Request.Form("AGCE2ScoreH")) and Len(Trim(Request.Form("AGCE2ScoreH"))) > 0 Then
					rs("RiskFactor2Score") = Request.Form("AGCE2ScoreH")	' Nullable: YES Type: smallint
				End If
			End If
			If Not IsEmpty(Request.Form("AGCE3")) Then
				rs("RiskFactor3") = Trim(Mid(Request.Form("AGCE3"),1,100))	' Nullable: YES Type: varchar
			End If
			If Not IsEmpty(Request.Form("AGCE3ScoreH")) Then
				If IsNumeric(Request.Form("AGCE3ScoreH")) and Len(Trim(Request.Form("AGCE3ScoreH"))) > 0 Then
					rs("RiskFactor3Score") = Request.Form("AGCE3ScoreH")	' Nullable: YES Type: smallint
				End If
			End If
			If Not IsEmpty(Request.Form("AGCE4")) Then
				rs("RiskFactor4") = Trim(Mid(Request.Form("AGCE4"),1,100))	' Nullable: YES Type: varchar
			End If
			If Not IsEmpty(Request.Form("AGCE4ScoreH")) Then
				If IsNumeric(Request.Form("AGCE4ScoreH")) and Len(Trim(Request.Form("AGCE4ScoreH"))) > 0 Then
					rs("RiskFactor4Score") = Request.Form("AGCE4ScoreH")	' Nullable: YES Type: smallint
				End If
			End If
			If Not IsEmpty(Request.Form("AGCE5")) Then
				rs("RiskFactor5") = Trim(Mid(Request.Form("AGCE5"),1,100))	' Nullable: YES Type: varchar
			End If
			If Not IsEmpty(Request.Form("AGCE5ScoreH")) Then
				If IsNumeric(Request.Form("AGCE5ScoreH")) and Len(Trim(Request.Form("AGCE5ScoreH"))) > 0 Then
					rs("RiskFactor5Score") = Request.Form("AGCE5ScoreH")	' Nullable: YES Type: smallint
				End If
			End If
			If Not IsEmpty(Request.Form("AGCERiskScoreH")) Then
				If IsNumeric(Request.Form("AGCERiskScoreH")) and Len(Trim(Request.Form("AGCERiskScoreH"))) > 0 Then
					rs("RiskScore") = Request.Form("AGCERiskScoreH")	' Nullable: YES Type: decimal
				End If
			End If		
		
		ElseIf Request.Form("txtRiskAssessmentGroup") = "F" Then

			If Not IsEmpty(Request.Form("AGF1")) Then
				rs("RiskFactor1") = Trim(Mid(Request.Form("AGF1"),1,100))	' Nullable: YES Type: varchar
			End If
			If Not IsEmpty(Request.Form("AGF1ScoreH")) Then
				If IsNumeric(Request.Form("AGF1ScoreH")) and Len(Trim(Request.Form("AGF1ScoreH"))) > 0 Then
					rs("RiskFactor1Score") = Request.Form("AGF1ScoreH")	' Nullable: YES Type: smallint
				End If
			End If
			If Not IsEmpty(Request.Form("AGF2")) Then
				rs("RiskFactor2") = Trim(Mid(Request.Form("AGF2"),1,100))	' Nullable: YES Type: varchar
			End If
			If Not IsEmpty(Request.Form("AGF2ScoreH")) Then
				If IsNumeric(Request.Form("AGF2ScoreH")) and Len(Trim(Request.Form("AGF2ScoreH"))) > 0 Then
					rs("RiskFactor2Score") = Request.Form("AGF2ScoreH")	' Nullable: YES Type: smallint
				End If
			End If
			If Not IsEmpty(Request.Form("AGF3")) Then
				rs("RiskFactor3") = Trim(Mid(Request.Form("AGF3"),1,100))	' Nullable: YES Type: varchar
			End If
			If Not IsEmpty(Request.Form("AGF3ScoreH")) Then
				If IsNumeric(Request.Form("AGF3ScoreH")) and Len(Trim(Request.Form("AGF3ScoreH"))) > 0 Then
					rs("RiskFactor3Score") = Request.Form("AGF3ScoreH")	' Nullable: YES Type: smallint
				End If
			End If
			If Not IsEmpty(Request.Form("AGF4")) Then
				rs("RiskFactor4") = Trim(Mid(Request.Form("AGF4"),1,100))	' Nullable: YES Type: varchar
			End If
			If Not IsEmpty(Request.Form("AGF4ScoreH")) Then
				If IsNumeric(Request.Form("AGF4ScoreH")) and Len(Trim(Request.Form("AGF4ScoreH"))) > 0 Then
					rs("RiskFactor4Score") = Request.Form("AGF4ScoreH")	' Nullable: YES Type: smallint
				End If
			End If
			If Not IsEmpty(Request.Form("AGF5")) Then
				rs("RiskFactor5") = Trim(Mid(Request.Form("AGF5"),1,100))	' Nullable: YES Type: varchar
			End If
			If Not IsEmpty(Request.Form("AGF5ScoreH")) Then
				If IsNumeric(Request.Form("AGF5ScoreH")) and Len(Trim(Request.Form("AGF5ScoreH"))) > 0 Then
					rs("RiskFactor5Score") = Request.Form("AGF5ScoreH")	' Nullable: YES Type: smallint
				End If
			End If
			If Not IsEmpty(Request.Form("AGFRiskScoreH")) Then
				If IsNumeric(Request.Form("AGFRiskScoreH")) and Len(Trim(Request.Form("AGFRiskScoreH"))) > 0 Then
					rs("RiskScore") = Request.Form("AGFRiskScoreH")	' Nullable: YES Type: decimal
				End If
			End If		
		
		End If
		
	End If

    ' Saving new values  Remi
    '============================================================================================================================
        If Not IsEmpty(Request.Form("RemiRisk1")) Then
			rs("UDFChar11") = Trim(Mid(Request.Form("RemiRisk1"),1,50))	' Nullable: YES Type: varchar
            rs("UDFChar21") = Trim(Mid(Request.Form("RemiRisk1DescH").Item,1,50)) 
        End If
        If Not IsEmpty(Request.Form("RemiRisk6")) Then
			rs("UDFChar12") = Trim(Mid(Request.Form("RemiRisk6"),1,50))	' Nullable: YES Type: varchar
			rs("UDFChar22") = Trim(Mid(Request.Form("RemiRisk6DescH").Item,1,50))	' Nullable: YES Type: varchar
		End If
        If Not IsEmpty(Request.Form("RemiRisk2")) Then
			rs("UDFChar13") = Trim(Mid(Request.Form("RemiRisk2"),1,50))	' Nullable: YES Type: varchar
            rs("UDFChar23") = Trim(Mid(Request.Form("RemiRisk2DescH").Item,1,50))	' Nullable: YES Type: varchar
		End If
        If Not IsEmpty(Request.Form("RemiRisk7")) Then
			rs("UDFChar14") = Trim(Mid(Request.Form("RemiRisk7"),1,50))	' Nullable: YES Type: varchar
            rs("UDFChar24") = Trim(Mid(Request.Form("RemiRisk7DescH").Item,1,50))	' Nullable: YES Type: varchar
		End If
        If Not IsEmpty(Request.Form("RemiRisk3")) Then
			rs("UDFChar15") = Trim(Mid(Request.Form("RemiRisk3"),1,50))	' Nullable: YES Type: varchar
            rs("UDFChar25") = Trim(Mid(Request.Form("RemiRisk3DescH").Item,1,50))	' Nullable: YES Type: varchar
		End If
        If Not IsEmpty(Request.Form("RemiRisk9")) Then
			rs("UDFChar16") = Trim(Mid(Request.Form("RemiRisk9"),1,50))	' Nullable: YES Type: varchar
            rs("UDFChar26") = Trim(Mid(Request.Form("RemiRisk9DescH").Item,1,50))	' Nullable: YES Type: varchar
        End If
        If Not IsEmpty(Request.Form("RemiRisk10")) Then
			rs("UDFChar18") = Trim(Mid(Request.Form("RemiRisk10"),1,50))	' Nullable: YES Type: varchar
            rs("UDFChar28") = Trim(Mid(Request.Form("RemiRisk10DescH").Item,1,50))	' Nullable: YES Type: varchar
        End If
        If Not IsEmpty(Request.Form("RemiRisk11")) Then
			rs("UDFChar19") = Trim(Mid(Request.Form("RemiRisk11"),1,50))	' Nullable: YES Type: varchar
            rs("UDFChar29") = Trim(Mid(Request.Form("RemiRisk11DescH").Item,1,50))	' Nullable: YES Type: varchar
        End If
        If Not IsEmpty(Request.Form("RemiRisk12")) Then
			rs("UDFChar20") = Trim(Mid(Request.Form("RemiRisk12"),1,50))	' Nullable: YES Type: varchar
            rs("UDFChar30") = Trim(Mid(Request.Form("RemiRisk12DescH").Item,1,50))	' Nullable: YES Type: varchar
        End If
        If Not IsEmpty(Request.Form("RemiRisk16")) Then
			rs("UDFChar31") = Trim(Mid(Request.Form("RemiRisk16"),1,50))	' Nullable: YES Type: varchar
            rs("UDFChar41") = Trim(Mid(Request.Form("RemiRisk16DescH").Item,1,50))	' Nullable: YES Type: varchar
        End If
        If Not IsEmpty(Request.Form("RemiRisk13")) Then
			rs("UDFChar32") = Trim(Mid(Request.Form("RemiRisk13"),1,50))	' Nullable: YES Type: varchar
        End If
        If Not IsEmpty(Request.Form("RemiRisk15")) Then
			rs("UDFChar33") = Trim(Mid(Request.Form("RemiRisk15"),1,50))	' Nullable: YES Type: varchar
        End If
        If Not IsEmpty(Request.Form("RemiRisk17")) Then
			rs("UDFChar34") = Trim(Mid(Request.Form("RemiRisk17"),1,50))	' Nullable: YES Type: varchar
        End If
        If Not IsEmpty(Request.Form("RemiRisk19")) Then
			rs("UDFChar35") = Trim(Mid(Request.Form("RemiRisk19"),1,50))	' Nullable: YES Type: varchar
        End If
        If Not IsEmpty(Request.Form("RemiRisk20")) Then
			rs("UDFChar36") = Trim(Mid(Request.Form("RemiRisk20"),1,50))	' Nullable: YES Type: varchar
        End If
        If Not IsEmpty(Request.Form("RemiRisk21")) Then
			rs("UDFChar37") = Trim(Mid(Request.Form("RemiRisk21"),1,50))	' Nullable: YES Type: varchar
        End If
        If Not IsEmpty(Request.Form("txtSetBy1")) Then
			rs("UDFChar38") = Trim(Mid(Request.Form("txtSetBy1"),1,50))	' Nullable: YES Type: varchar
            rs("UDFChar39") = Trim(Mid(Request.Form("txtSetBy1DescH").Item,1,50))	' Nullable: YES Type: varchar
        End If
        If Not IsEmpty(Request.Form("txtSetBy2")) Then
			rs("UDFChar42") = Trim(Mid(Request.Form("txtSetBy2"),1,50))	' Nullable: YES Type: varchar
            rs("UDFChar43") = Trim(Mid(Request.Form("txtSetBy2DescH").Item,1,50))	' Nullable: YES Type: varchar
        End If
        If Not IsEmpty(Request.Form("txtDate")) Then
			rs("UDFChar44") = Trim(Mid(Request.Form("txtDate"),1,50))	' Nullable: YES Type: varchar
        End If
        If Not IsEmpty(Request.Form("txtDate2")) Then
			rs("UDFChar48") = Trim(Mid(Request.Form("txtDate2"),1,50))	' Nullable: YES Type: varchar
        End If
        If Not IsEmpty(Request.Form("txtDate3")) Then
			rs("UDFChar45") = Trim(Mid(Request.Form("txtDate3"),1,50))	' Nullable: YES Type: varchar
        End If
                If Not IsEmpty(Request.Form("txtSetBy3")) Then
			rs("UDFChar46") = Trim(Mid(Request.Form("txtSetBy3"),1,50))	' Nullable: YES Type: varchar
            rs("UDFChar47") = Trim(Mid(Request.Form("txtSetBy3DescH").Item,1,50))	' Nullable: YES Type: varchar
        End If
        rs("UDFChar17") = Not Request.Form("txtMaintenanceHistory") = ""
        rs("UDFChar27") = Not Request.Form("txtAssistantDirector") = ""
        'rs("PlanForImprovement") =  Request.Form("txtPlanForImprovement") = ""	' Nullable: YES Type: bit
            ''''''''''''''''''''''''''rs("UDFChar27") = Trim(Mid(Request.Form("RemiRisk4DescH").Item,1,50))	' Nullable: YES Type: varchar
		'''''''''End If


	'CB AUTOCALC FIX 1/10/2013
	'============================================================================================================================
	If CheckIfFieldExists(RS,"RiskLevelAutoCalc") Then
		rs("RiskLevelAutoCalc") = Not Request.Form("txtRiskLevelAutoCalc") = ""	' Nullable: YES Type: bit
		rs("PMRequiredAutoCalc") = Not Request.Form("txtPMRequiredAutoCalc") = ""	' Nullable: YES Type: bit
	End If

	If Not IsEmpty(Request.Form("txtRiskLevel")) Then
		If Len(Trim(Request.Form("txtRiskLevel"))) > 0 Then
			rs("RiskLevel") = Request.Form("txtRiskLevel")	' Nullable: YES Type: smallint
		End If
	End If
	rs("PMRequired") = Not Request.Form("txtPMRequired") = ""	' Nullable: YES Type: bit
	rs("PlanForImprovement") = Not Request.Form("txtPlanForImprovement") = ""	' Nullable: YES Type: bit
	rs("HIPPARelated") = Not Request.Form("txtHIPPARelated") = ""	' Nullable: YES Type: bit
	rs("StatementOfConditions") = Not Request.Form("txtStatementOfConditions") = ""	' Nullable: YES Type: bit
	rs("StatementOfConditionsCompliant") = Not Request.Form("txtStatementOfConditionsCompliant") = ""	' Nullable: YES Type: bit



	If Not IsEmpty(Request.Form("txtZonePK").Item) Then
		If Len(Trim(Request.Form("txtZonePK").Item)) > 0 Then
			rs("ZonePK") = Request.Form("txtZonePK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtZone").Item) Then
				rs("ZoneID") = Trim(Mid(Request.Form("txtZone").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtZoneDescH").Item) Then
				rs("ZoneName") = Trim(Mid(Request.Form("txtZoneDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("ZonePK") = Null
			rs("ZoneID") = Null
			rs("ZoneName") = Null
		End If
	End If

	If Not IsEmpty(Request.Form("txtZoneColor").Item) Then
		If Len(Trim(Request.Form("txtZoneColor").Item)) > 0 Then
			rs("ZoneColor") = Trim(Mid(Request.Form("txtZoneColor").Item,1,10))	' Nullable: YES Type: nvarchar
		Else
			rs("ZoneColor") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtCounty")) Then
		rs("County") = Trim(Mid(Request.Form("txtCounty"),1,50))	' Nullable: YES Type: varchar
	End If
	If Not IsEmpty(Request.Form("txtYearBuilt")) Then
		If Not Request.Form("txtYearBuilt") = "" Then
			rs("YearBuilt") = SQLdatetimeADO(Request.Form("txtYearBuilt"))	' Nullable: YES Type: datetime
		End If
	End If
	If Not IsEmpty(Request.Form("txtMajorRenovations")) Then
		rs("MajorRenovations") = Trim(Mid(Request.Form("txtMajorRenovations"),1,50))	' Nullable: YES Type: varchar
	End If
	If Not IsEmpty(Request.Form("txtSquareFootage")) Then
		If Len(Trim(Request.Form("txtSquareFootage"))) > 0 Then
			rs("SquareFootage") = FixInternationalNumber(Request.Form("txtSquareFootage"))	' Nullable: YES Type: real
		End If
	End If
	If Not IsEmpty(Request.Form("txtConstructionCode")) Then
		If Len(Trim(Request.Form("txtConstructionCode"))) > 0 Then
			rs("ConstructionCode") = Trim(Mid(Request.Form("txtConstructionCode"),1,25))	' Nullable: YES Type: varchar
			If Not IsEmpty(Request.Form("txtConstructionCodeDescH")) Then
				rs("ConstructionCodeDesc") = Trim(Mid(Request.Form("txtConstructionCodeDescH"),1,50))	' Nullable: YES Type: varchar
			End If
		Else
			rs("ConstructionCode") = Null
			rs("ConstructionCodeDesc") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtNumberOfStories")) Then
		If Len(Trim(Request.Form("txtNumberOfStories"))) > 0 Then
			rs("NumberOfStories") = Request.Form("txtNumberOfStories")	' Nullable: YES Type: smallint
		End If
	End If
	If Not IsEmpty(Request.Form("txtISOProtection")) Then
		If Len(Trim(Request.Form("txtISOProtection"))) > 0 Then
			rs("ISOProtection") = Trim(Mid(Request.Form("txtISOProtection"),1,25))	' Nullable: YES Type: varchar
			If Not IsEmpty(Request.Form("txtISOProtectionDescH")) Then
				rs("ISOProtectionDesc") = Trim(Mid(Request.Form("txtISOProtectionDescH"),1,50))	' Nullable: YES Type: varchar
			End If
		Else
			rs("ISOProtection") = Null
			rs("ISOProtectionDesc") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtAutoSprinkler")) Then
		If Len(Trim(Request.Form("txtAutoSprinkler"))) > 0 Then
			rs("AutoSprinkler") = Trim(Mid(Request.Form("txtAutoSprinkler"),1,25))	' Nullable: YES Type: varchar
			If Not IsEmpty(Request.Form("txtAutoSprinklerDescH")) Then
				rs("AutoSprinklerDesc") = Trim(Mid(Request.Form("txtAutoSprinklerDescH"),1,50))	' Nullable: YES Type: varchar
			End If
		Else
			rs("AutoSprinkler") = Null
			rs("AutoSprinklerDesc") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtSmokeAlarm")) Then
		If Len(Trim(Request.Form("txtSmokeAlarm"))) > 0 Then
			rs("SmokeAlarm") = Trim(Mid(Request.Form("txtSmokeAlarm"),1,25))	' Nullable: YES Type: varchar
			If Not IsEmpty(Request.Form("txtSmokeAlarmDescH")) Then
				rs("SmokeAlarmDesc") = Trim(Mid(Request.Form("txtSmokeAlarmDescH"),1,50))	' Nullable: YES Type: varchar
			End If
		Else
			rs("SmokeAlarm") = Null
			rs("SmokeAlarmDesc") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtHeatAlarm")) Then
		If Len(Trim(Request.Form("txtHeatAlarm"))) > 0 Then
			rs("HeatAlarm") = Trim(Mid(Request.Form("txtHeatAlarm"),1,25))	' Nullable: YES Type: varchar
			If Not IsEmpty(Request.Form("txtHeatAlarmDescH")) Then
				rs("HeatAlarmDesc") = Trim(Mid(Request.Form("txtHeatAlarmDescH"),1,50))	' Nullable: YES Type: varchar
			End If
		Else
			rs("HeatAlarm") = Null
			rs("HeatAlarmDesc") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtFloodZone")) Then
		rs("FloodZone") = Trim(Mid(Request.Form("txtFloodZone"),1,50))	' Nullable: YES Type: varchar
	End If
	If Not IsEmpty(Request.Form("txtQuakeZone")) Then
		rs("QuakeZone") = Trim(Mid(Request.Form("txtQuakeZone"),1,50))	' Nullable: YES Type: varchar
	End If
	If Not IsEmpty(Request.Form("txtExt100Feet")) Then
		rs("Ext100Feet") = Trim(Mid(Request.Form("txtExt100Feet"),1,50))	' Nullable: YES Type: varchar
	End If
	If Not IsEmpty(Request.Form("txtOperatingUnits")) Then
		If Len(Trim(Request.Form("txtOperatingUnits"))) > 0 Then
			rs("OperatingUnits") = Request.Form("txtOperatingUnits")	' Nullable: YES Type: smallint
		End If
	End If
	If Not IsEmpty(Request.Form("txtEstimatedValue")) Then
		If Len(Trim(Request.Form("txtEstimatedValue"))) > 0 Then
			rs("EstimatedValue") = FixInternationalNumber(Request.Form("txtEstimatedValue"))	' Nullable: YES Type: float
		End If
	End If
	If Not IsEmpty(Request.Form("txtResponsibilityRepairPK")) Then
		If Len(Trim(Request.Form("txtResponsibilityRepairPK"))) > 0 Then
			rs("ResponsibilityRepairPK") = Request.Form("txtResponsibilityRepairPK")	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtResponsibilityRepair")) Then
				rs("ResponsibilityRepairID") = Trim(Mid(Request.Form("txtResponsibilityRepair"),1,25))	' Nullable: YES Type: varchar
			End If
			If Not IsEmpty(Request.Form("txtResponsibilityRepairDescH")) Then
				rs("ResponsibilityRepairName") = Trim(Mid(Request.Form("txtResponsibilityRepairDescH"),1,50))	' Nullable: YES Type: varchar
			End If
		Else
			rs("ResponsibilityRepairPK") = Null
			rs("ResponsibilityRepairID") = Null
			rs("ResponsibilityRepairName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtResponsibilityPMPK")) Then
		If Len(Trim(Request.Form("txtResponsibilityPMPK"))) > 0 Then
			rs("ResponsibilityPMPK") = Request.Form("txtResponsibilityPMPK")	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtResponsibilityPM")) Then
				rs("ResponsibilityPMID") = Trim(Mid(Request.Form("txtResponsibilityPM"),1,25))	' Nullable: YES Type: varchar
			End If
			If Not IsEmpty(Request.Form("txtResponsibilityPMDescH")) Then
				rs("ResponsibilityPMName") = Trim(Mid(Request.Form("txtResponsibilityPMDescH"),1,50))	' Nullable: YES Type: varchar
			End If
		Else
			rs("ResponsibilityPMPK") = Null
			rs("ResponsibilityPMID") = Null
			rs("ResponsibilityPMName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtResponsibilitySafetyPK")) Then
		If Len(Trim(Request.Form("txtResponsibilitySafetyPK"))) > 0 Then
			rs("ResponsibilitySafetyPK") = Request.Form("txtResponsibilitySafetyPK")	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtResponsibilitySafety")) Then
				rs("ResponsibilitySafetyID") = Trim(Mid(Request.Form("txtResponsibilitySafety"),1,25))	' Nullable: YES Type: varchar
			End If
			If Not IsEmpty(Request.Form("txtResponsibilitySafetyDescH")) Then
				rs("ResponsibilitySafetyName") = Trim(Mid(Request.Form("txtResponsibilitySafetyDescH"),1,50))	' Nullable: YES Type: varchar
			End If
		Else
			rs("ResponsibilitySafetyPK") = Null
			rs("ResponsibilitySafetyID") = Null
			rs("ResponsibilitySafetyName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtServiceRepairPK")) Then
		If Len(Trim(Request.Form("txtServiceRepairPK"))) > 0 Then
			rs("ServiceRepairPK") = Request.Form("txtServiceRepairPK")	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtServiceRepair")) Then
				rs("ServiceRepairID") = Trim(Mid(Request.Form("txtServiceRepair"),1,25))	' Nullable: YES Type: varchar
			End If
			If Not IsEmpty(Request.Form("txtServiceRepairDescH")) Then
				rs("ServiceRepairName") = Trim(Mid(Request.Form("txtServiceRepairDescH"),1,50))	' Nullable: YES Type: varchar
			End If
		Else
			rs("ServiceRepairPK") = Null
			rs("ServiceRepairID") = Null
			rs("ServiceRepairName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtServicePMPK")) Then
		If Len(Trim(Request.Form("txtServicePMPK"))) > 0 Then
			rs("ServicePMPK") = Request.Form("txtServicePMPK")	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtServicePM")) Then
				rs("ServicePMID") = Trim(Mid(Request.Form("txtServicePM"),1,25))	' Nullable: YES Type: varchar
			End If
			If Not IsEmpty(Request.Form("txtServicePMDescH")) Then
				rs("ServicePMName") = Trim(Mid(Request.Form("txtServicePMDescH"),1,50))	' Nullable: YES Type: varchar
			End If
		Else
			rs("ServicePMPK") = Null
			rs("ServicePMID") = Null
			rs("ServicePMName") = Null
		End If
	End If

	Call SaveUDF(rs)

	' CUSTOMIZED
	'-----------------------------------------------------------------------------------------------------------------
	rs("RowVersionAction") = "EDIT"
	'-----------------------------------------------------------------------------------------------------------------

	Call db_version(rs)

	'rs.Properties("Update Resync") = adResyncAutoIncrement
	'rs.Properties("Server Data On Insert").Value = True

	db.dobatchupdate rs

	If db.dok Then

		keyvalue = rs("AssetPK")
		Call CloseObj(rs)

		If newrecord or duprecord Then
			Set rs = db.RunSQLReturnRS_RW("SELECT TOP 0 * FROM AssetHierarchy","")
			If Not db.dok Then
				Exit Sub
			End If
			rs.AddNew()
			AssetUpdate = False
			rs("System") = "MC"
			rs("AssetPK") = keyvalue
			rs("ParentPK") = Trim(Request.Form("txtParentPK").Item)	' Nullable: NO Type: int
			db.dobatchupdate rs
			IsParentChange = True
		Else
			If Not Trim(Request.Form("txtParentPK").Item) = "" Then

				Set rs = db.RunSQLReturnRS("SELECT * FROM AssetHierarchy WITH (NOLOCK) WHERE AssetPK = (" & keyvalue & ") AND (System = 'MC')","")
				If Not db.dok Then
					Exit Sub
				End If

				If Not Trim(rs("ParentPK")) = Trim(Request.Form("txtParentPK").Item) Then
					If NullCheck(rs("OrderType")) = "-7777" Then
						Call db.RunSQL("UPDATE AssetHierarchy WITH (ROWLOCK) SET OrderType = 0 WHERE AssetPK = (" & keyvalue & ") AND (System = 'MC')","")
					Else
						IsParentChange = True

						'@$CUSTOMISED
						Call db.RunSP("MC_CheckCompoundAssetStockroomTransfer",Array(Array("@FromPK", adInteger, adParamInput, 4, keyvalue),Array("@ToPK", adInteger, adParamInput, 4, Trim(Request.Form("txtParentPK").Item)),Array("@ReturnCode", adInteger, adParamOutPut, 4, "")),OutArray)				
						'Response.Write OutArray(2)
						'Response.End
						If OutArray(2) = -1 Then
							db.dok = False
							db.derror = "Unable to Transfer Compound Assets via this method.<br>Compound Assets can only be transferred via the Asset Tree, and only if they do not include Rotating Parts."
							'Call db.RunSP("MC_MoveCompoundAsset",Array(Array("@FromPK", adInteger, adParamInput, 4, keyvalue),Array("@ToPK", adInteger, adParamInput, 4, Trim(Request.Form("txtParentPK").Item)),Array("@UserPK", adInteger, adParamInput, 4, Request.Form("txtRowVersionUserPK").Item),Array("@Initials", MC_ADVARCHAR, adParamInput, 5, Trim(Mid(Request.Form("txtRowVersionInitials").Item,1,5))),Array("@ErrorCode", adInteger, adParamOutPut, 4, "")),OutArray)
						Else
							Call db.RunSP("MC_MoveAssetTree",Array(Array("@FromPK", adInteger, adParamInput, 4, keyvalue),Array("@ToPK", adInteger, adParamInput, 4, Trim(Request.Form("txtParentPK").Item)),Array("@UserPK", adInteger, adParamInput, 4, Request.Form("txtRowVersionUserPK").Item),Array("@Initials", MC_ADVARCHAR, adParamInput, 5, Trim(Mid(Request.Form("txtRowVersionInitials").Item,1,5))),Array("@ErrorCode", adInteger, adParamOutPut, 4, "")),OutArray)
						End If

						'Call db.RunSP("MC_MoveAssetTree",Array(Array("@FromPK", adInteger, adParamInput, 4, keyvalue),Array("@ToPK", adInteger, adParamInput, 4, Trim(Request.Form("txtParentPK").Item)),Array("@UserPK", adInteger, adParamInput, 4, Request.Form("txtRowVersionUserPK").Item),Array("@Initials", MC_ADVARCHAR, adParamInput, 5, Trim(Mid(Request.Form("txtRowVersionInitials").Item,1,5))),Array("@ErrorCode", adInteger, adParamOutPut, 4, "")),OutArray)

						' AssetUpdate is done is MC_MoveAsset
						AssetUpdate = False

						'@$CUSTOMISED
						If OutArray(4) = -1 Then
							db.dok = False
							db.derror = "The Parent is an ancestor of the Asset. This is not permitted. Please select a Parent that does not belong to the Asset."
						ElseIf OutArray(4) = -2 Then
							db.dok = False
							db.derror = "Unable to Transfer Rotating Parts.<br>Rotating Parts can only be transferred via Work Order Cost functionality."
						End If
					End If
				End If

			End If
		End If

		Call CloseObj(rs)

		If AssetUpdate Then
			Dim tmpParentPK
			tmpParentPK = Trim(Request.Form("txtParentPK").Item)
			If tmpParentPK = "" Then
				tmpParentPK = Null
			End If


			Call db.RunSP("MC_AssetHierarchyUpdateAll",Array(Array("@AssetPK", adInteger, adParamInput, 4, keyvalue)),"")

		End If

	Else
		Call CloseObj(rs)
	End If

End Sub

Sub db_sp(rs,isinsert,suffix,htmltable,theindex)

	Dim fp
	fp = "txt" & LCase(htmltable) & suffix

	' -- Start Table Fields ------------------------------------------------

	rs("AssetPK") = keyvalue

	If Len(Trim(Request.Form(fp & "Specification" & theindex & "PK").Item)) > 0 Then
		rs("SpecificationPK") = Request.Form(fp & "Specification" & theindex & "PK").Item	' Nullable: No Type: int
	End If
	If Not IsEmpty(Request.Form(fp & "Specification" & theindex & "DescH").Item) Then
		rs("SpecificationName") = Trim(Mid(Request.Form(fp & "Specification" & theindex & "DescH").Item,1,50))	' Nullable: YES Type: nvarchar
	End If
	If Not IsEmpty(Request.Form(fp & "ValueText" & theindex).Item) Then
		If Len(Request.Form(fp & "ValueText" & theindex).Item) > 0 Then
			rs("ValueText") = Trim(Mid(Request.Form(fp & "ValueText" & theindex).Item,1,6000))	' Nullable: YES Type: nvarchar
		Else
			rs("ValueText") = Null
		End If
	End If
	If Len(Trim(Request.Form(fp & "ValueNumeric" & theindex).Item)) > 0 Then
		rs("ValueNumeric") = FixInternationalNumber(Request.Form(fp & "ValueNumeric" & theindex).Item)	' Nullable: No Type: real
	Else
		rs("ValueNumeric") = Null
	End If
	If Not IsEmpty(Request.Form(fp & "ValueDate" & theindex).Item) Then
		If Not Request.Form(fp & "ValueDate" & theindex).Item = "" Then
			rs("ValueDate") = SQLdatetimeADO(Request.Form(fp & "ValueDate" & theindex).Item)	' Nullable: YES Type: datetime
		Else
			rs("ValueDate") = Null
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "ValueHi" & theindex).Item) Then
		If Len(Trim(Request.Form(fp & "ValueHi" & theindex).Item)) > 0 Then
			rs("ValueHi") = FixInternationalNumber(Request.Form(fp & "ValueHi" & theindex).Item)	' Nullable: YES Type: real
		Else
			rs("ValueHi") = Null
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "ValueLow" & theindex).Item) Then
		If Len(Trim(Request.Form(fp & "ValueLow" & theindex).Item)) > 0 Then
			rs("ValueLow") = FixInternationalNumber(Request.Form(fp & "ValueLow" & theindex).Item)	' Nullable: YES Type: real
		Else
			rs("ValueLow") = Null
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "ValueOptimal" & theindex).Item) Then
		If Len(Trim(Request.Form(fp & "ValueOptimal" & theindex).Item)) > 0 Then
			rs("ValueOptimal") = FixInternationalNumber(Request.Form(fp & "ValueOptimal" & theindex).Item)	' Nullable: YES Type: real
		Else
			rs("ValueOptimal") = Null
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "Comments" & theindex).Item) Then
		If Len(Request.Form(fp & "Comments" & theindex).Item) > 0 Then
			rs("Comments") = Trim(Mid(Request.Form(fp & "Comments" & theindex).Item,1,4000))	' Nullable: YES Type: nvarchar
		Else
			rs("Comments") = Null
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "Procedure" & theindex & "PK")) Then
		If Not Request.Form(fp & "Procedure" & theindex & "PK") = "UNCHANGED" Then
			If Len(Trim(Request.Form(fp & "Procedure" & theindex & "PK"))) > 0 Then
				rs("ProcedurePK") = Request.Form(fp & "Procedure" & theindex & "PK")	' Nullable: YES Type: int
			Else
				rs("ProcedurePK") = Null
			End If
		End If
	End If

	rs("ValueOutOfRangeWO") = (InStr(Request.Form(fp & "ValueOutOfRangeWO" & theindex).Item,"taskchecked.gif") > 0) ' Nullable: No Type: bit
	rs("TrackHistory") = (InStr(Request.Form(fp & "TrackHistory" & theindex).Item,"taskchecked.gif") > 0) ' Nullable: No Type: bit

	' -- End Table Fields ------------------------------------------------

	Call db_version(rs)

End Sub

Sub db_am(rs,isinsert,suffix,htmltable,theindex)

	Dim fp
	fp = "txt" & LCase(htmltable) & suffix

	' -- Start Table Fields ------------------------------------------------

	rs("AssetPK") = keyvalue

	If Len(Trim(Request.Form(fp & "Part" & theindex & "PK").Item)) > 0 Then
		rs("PartPK") = Request.Form(fp & "Part" & theindex & "PK").Item	' Nullable: No Type: int
	End If
	If Not IsEmpty(Request.Form(fp & "Qty" & theindex).Item) Then
		If Len(Trim(Request.Form(fp & "Qty" & theindex).Item)) > 0 Then
			rs("Qty") = Request.Form(fp & "Qty" & theindex).Item	' Nullable: YES Type: int
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "Comments" & theindex).Item) Then
		If Len(Request.Form(fp & "Comments" & theindex).Item) > 0 Then
			rs("Comments") = Trim(Mid(Request.Form(fp & "Comments" & theindex).Item,1,2000))	' Nullable: YES Type: nvarchar
		Else
			rs("Comments") = Null
		End If
	End If

	' -- End Table Fields ------------------------------------------------

	Call db_version(rs)

End Sub

Sub db_ap(rs,isinsert,suffix,htmltable,theindex)

	Dim fp
	fp = "txt" & LCase(htmltable) & suffix

	' -- Start Table Fields ------------------------------------------------

	rs("AssetPK") = keyvalue

	If Len(Trim(Request.Form(fp & "PM" & theindex & "PK").Item)) > 0 Then
		rs("PMPK") = Request.Form(fp & "PM" & theindex & "PK").Item	' Nullable: No Type: int
	End If
	If Len(Trim(Request.Form(fp & "RouteOrder" & theindex).Item)) > 0 Then
		rs("RouteOrder") = Request.Form(fp & "RouteOrder" & theindex).Item	' Nullable: No Type: real
	End If
	If Not IsEmpty(Request.Form(fp & "LastGeneratedDate" & theindex).Item) Then
		If Not Request.Form(fp & "LastGeneratedDate" & theindex).Item = "" Then
			rs("LastGeneratedDate") = SQLdatetimeADO(Request.Form(fp & "LastGeneratedDate" & theindex).Item)	' Nullable: YES Type: datetime
		Else
			rs("LastGeneratedDate") = Null
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "LastCompletedDate" & theindex).Item) Then
		If Not Request.Form(fp & "LastCompletedDate" & theindex).Item = "" Then
			rs("LastCompletedDate") = SQLdatetimeADO(Request.Form(fp & "LastCompletedDate" & theindex).Item)	' Nullable: YES Type: datetime
		Else
			rs("LastCompletedDate") = Null
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "NextScheduledDate" & theindex).Item) Then
		If Not Request.Form(fp & "NextScheduledDate" & theindex).Item = "" Then
			rs("NextScheduledDate") = SQLdatetimeADO(Request.Form(fp & "NextScheduledDate" & theindex).Item)	' Nullable: YES Type: datetime
		Else
			rs("NextScheduledDate") = Null
		End If
	End If
	If Len(Trim(Request.Form(fp & "TimesCounter" & theindex).Item)) > 0 Then
		rs("TimesCounter") = Request.Form(fp & "TimesCounter" & theindex).Item	' Nullable: No Type: real
	End If
	If Len(Trim(Request.Form(fp & "PMCounter" & theindex).Item)) > 0 Then
		rs("PMCounter") = Request.Form(fp & "PMCounter" & theindex).Item	' Nullable: No Type: real
	End If
	If Not IsEmpty(Request.Form(fp & "Meter1ReadingLastInterval" & theindex).Item) Then
		If Len(Trim(Request.Form(fp & "Meter1ReadingLastInterval" & theindex).Item)) > 0 Then
			rs("Meter1ReadingLastInterval") = Request.Form(fp & "Meter1ReadingLastInterval" & theindex).Item	' Nullable: YES Type: real
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "Meter1NextInterval" & theindex).Item) Then
		If Len(Trim(Request.Form(fp & "Meter1NextInterval" & theindex).Item)) > 0 Then
			rs("Meter1NextInterval") = Request.Form(fp & "Meter1NextInterval" & theindex).Item	' Nullable: YES Type: real
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "Meter2ReadingLastInterval" & theindex).Item) Then
		If Len(Trim(Request.Form(fp & "Meter2ReadingLastInterval" & theindex).Item)) > 0 Then
			rs("Meter2ReadingLastInterval") = Request.Form(fp & "Meter2ReadingLastInterval" & theindex).Item	' Nullable: YES Type: real
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "Meter2NextInterval" & theindex).Item) Then
		If Len(Trim(Request.Form(fp & "Meter2NextInterval" & theindex).Item)) > 0 Then
			rs("Meter2NextInterval") = Request.Form(fp & "Meter2NextInterval" & theindex).Item	' Nullable: YES Type: real
		End If
	End If

	If Not IsEmpty(Request.Form(fp & "RepairCenter" & theindex & "PK").Item) Then
		If Not Request.Form(fp & "RepairCenter" & theindex & "PK").Item = "UNCHANGED" Then
			If Len(Trim(Request.Form(fp & "RepairCenter" & theindex & "PK").Item)) > 0 Then
				rs("RepairCenterPK") = Request.Form(fp & "RepairCenter" & theindex & "PK").Item	' Nullable: No Type: int
				rs("RepairCenterID") = Trim(Mid(Request.Form(fp & "RepairCenter" & theindex).Item,1,25))	' Nullable: No Type: nvarchar
				If Not IsEmpty(Request.Form(fp & "RepairCenter" & theindex & "DescH").Item) Then
					rs("RepairCenterName") = Trim(Mid(Request.Form(fp & "RepairCenter" & theindex & "DescH").Item,1,50))	' Nullable: YES Type: nvarchar
				End If
			Else
				'rs("RepairCenterPK") = Null
				'rs("RepairCenterID") = Null
				'rs("RepairCenterName") = Null
			End If
		End If
	End If

	If Not IsEmpty(Request.Form(fp & "StockRoom" & theindex & "PK").Item) Then
		If Not Request.Form(fp & "StockRoom" & theindex & "PK").Item = "UNCHANGED" Then
			If Len(Trim(Request.Form(fp & "StockRoom" & theindex & "PK").Item)) > 0 Then
				rs("StockRoomPK") = Request.Form(fp & "StockRoom" & theindex & "PK").Item	' Nullable: No Type: int
				rs("StockRoomID") = Trim(Mid(Request.Form(fp & "StockRoom" & theindex).Item,1,25))	' Nullable: No Type: nvarchar
				If Not IsEmpty(Request.Form(fp & "StockRoom" & theindex & "DescH").Item) Then
					rs("StockRoomName") = Trim(Mid(Request.Form(fp & "StockRoom" & theindex & "DescH").Item,1,50))	' Nullable: YES Type: nvarchar
				End If
			Else
				'rs("StockRoomPK") = Null
				'rs("StockRoomID") = Null
				'rs("StockRoomName") = Null
			End If
		End If
	End If

	If Not IsEmpty(Request.Form(fp & "ToolRoom" & theindex & "PK").Item) Then
		If Not Request.Form(fp & "ToolRoom" & theindex & "PK").Item = "UNCHANGED" Then
			If Len(Trim(Request.Form(fp & "ToolRoom" & theindex & "PK").Item)) > 0 Then
				rs("ToolRoomPK") = Request.Form(fp & "ToolRoom" & theindex & "PK").Item	' Nullable: No Type: int
				If Not IsEmpty(Request.Form(fp & "ToolRoom" & theindex).Item) Then
					rs("ToolRoomID") = Trim(Mid(Request.Form(fp & "ToolRoom" & theindex).Item,1,25))	' Nullable: No Type: nvarchar
				End If
				If Not IsEmpty(Request.Form(fp & "ToolRoom" & theindex & "DescH").Item) Then
					rs("ToolRoomName") = Trim(Mid(Request.Form(fp & "ToolRoom" & theindex & "DescH").Item,1,50))	' Nullable: YES Type: nvarchar
				End If
			Else
				'rs("ToolRoomPK") = Null
				'rs("ToolRoomID") = Null
				'rs("ToolRoomName") = Null
			End If
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "Shop" & theindex & "PK").Item) Then
		If Not Request.Form(fp & "Shop" & theindex & "PK").Item = "UNCHANGED" Then
			If Len(Trim(Request.Form(fp & "Shop" & theindex & "PK").Item)) > 0 Then
				rs("ShopPK") = Request.Form(fp & "Shop" & theindex & "PK").Item	' Nullable: YES Type: int
				If Not IsEmpty(Request.Form(fp & "Shop" & theindex).Item) Then
					rs("ShopID") = Trim(Mid(Request.Form(fp & "Shop" & theindex).Item,1,25))	' Nullable: YES Type: nvarchar
				End If
				If Not IsEmpty(Request.Form(fp & "Shop" & theindex & "DescH").Item) Then
					rs("ShopName") = Trim(Mid(Request.Form(fp & "Shop" & theindex & "DescH").Item,1,50))	' Nullable: YES Type: nvarchar
				End If
			Else
				rs("ShopPK") = Null
				rs("ShopID") = Null
				rs("ShopName") = Null
			End If
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "Shift" & theindex & "PK").Item) Then
		If Not Request.Form(fp & "Shift" & theindex & "PK").Item = "UNCHANGED" Then
			If Len(Trim(Request.Form(fp & "Shift" & theindex & "PK").Item)) > 0 Then
				rs("ShiftPK") = Request.Form(fp & "Shift" & theindex & "PK").Item	' Nullable: YES Type: int
				If Not IsEmpty(Request.Form(fp & "Shift" & theindex).Item) Then
					rs("ShiftID") = Trim(Mid(Request.Form(fp & "Shift" & theindex).Item,1,25))	' Nullable: YES Type: nvarchar
				End If
				If Not IsEmpty(Request.Form(fp & "Shift" & theindex & "DescH").Item) Then
					rs("ShiftName") = Trim(Mid(Request.Form(fp & "Shift" & theindex & "DescH").Item,1,50))	' Nullable: YES Type: nvarchar
				End If	
			Else
				rs("ShiftPK") = Null
				rs("ShiftID") = Null
				rs("ShiftName") = Null
			End If
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "Supervisor" & theindex & "PK").Item) Then
		If Not Request.Form(fp & "Supervisor" & theindex & "PK").Item = "UNCHANGED" Then
			If Len(Trim(Request.Form(fp & "Supervisor" & theindex & "PK").Item)) > 0 Then
				rs("SupervisorPK") = Request.Form(fp & "Supervisor" & theindex & "PK").Item	' Nullable: YES Type: int
				If Not IsEmpty(Request.Form(fp & "Supervisor" & theindex).Item) Then
					rs("SupervisorID") = Trim(Mid(Request.Form(fp & "Supervisor" & theindex).Item,1,25))	' Nullable: YES Type: nvarchar
				End If
				If Not IsEmpty(Request.Form(fp & "Supervisor" & theindex & "DescH").Item) Then
					rs("SupervisorName") = Trim(Mid(Request.Form(fp & "Supervisor" & theindex & "DescH").Item,1,50))	' Nullable: YES Type: nvarchar
				End If
			Else
				rs("SupervisorPK") = Null
				rs("SupervisorID") = Null
				rs("SupervisorName") = Null
			End If
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "Account" & theindex & "PK").Item) Then
		If Not Request.Form(fp & "Account" & theindex & "PK").Item = "UNCHANGED" Then
			If Len(Trim(Request.Form(fp & "Account" & theindex & "PK").Item)) > 0 Then
				rs("AccountPK") = Request.Form(fp & "Account" & theindex & "PK").Item	' Nullable: YES Type: int
				If Not IsEmpty(Request.Form(fp & "Account" & theindex).Item) Then
					rs("AccountID") = Trim(Mid(Request.Form(fp & "Account" & theindex).Item,1,25))	' Nullable: YES Type: nvarchar
				End If
				If Not IsEmpty(Request.Form(fp & "Account" & theindex & "DescH").Item) Then
					rs("AccountName") = Trim(Mid(Request.Form(fp & "Account" & theindex & "DescH").Item,1,50))	' Nullable: YES Type: nvarchar
				End If
			Else
				rs("AccountPK") = Null
				rs("AccountID") = Null
				rs("AccountName") = Null
			End If
		End If
	End If

	If Not IsEmpty(Request.Form(fp & "Department" & theindex & "PK").Item) Then
		If Not Request.Form(fp & "Department" & theindex & "PK").Item = "UNCHANGED" Then
			If Len(Trim(Request.Form(fp & "Department" & theindex & "PK").Item)) > 0 Then
				rs("DepartmentPK") = Request.Form(fp & "Department" & theindex & "PK").Item	' Nullable: YES Type: int
				If Not IsEmpty(Request.Form(fp & "Department" & theindex).Item) Then
					rs("DepartmentID") = Trim(Mid(Request.Form(fp & "Department" & theindex).Item,1,25))	' Nullable: YES Type: nvarchar
				End If
				If Not IsEmpty(Request.Form(fp & "Department" & theindex & "DescH").Item) Then
					rs("DepartmentName") = Trim(Mid(Request.Form(fp & "Department" & theindex & "DescH").Item,1,100))	' Nullable: YES Type: nvarchar
				End If
			Else
				rs("DepartmentPK") = Null
				rs("DepartmentID") = Null
				rs("DepartmentName") = Null
			End If
		End If
	End If

	If Not IsEmpty(Request.Form(fp & "Tenant" & theindex & "PK").Item) Then
		If Not Request.Form(fp & "Tenant" & theindex & "PK").Item = "UNCHANGED" Then
			If Len(Trim(Request.Form(fp & "Tenant" & theindex & "PK").Item)) > 0 Then
				rs("TenantPK") = Request.Form(fp & "Tenant" & theindex & "PK").Item	' Nullable: YES Type: int
				If Not IsEmpty(Request.Form(fp & "Tenant" & theindex).Item) Then
					rs("TenantID") = Trim(Mid(Request.Form(fp & "Tenant" & theindex).Item,1,25))	' Nullable: YES Type: nvarchar
				End If
				If Not IsEmpty(Request.Form(fp & "Tenant" & theindex & "DescH").Item) Then
					rs("TenantName") = Trim(Mid(Request.Form(fp & "Tenant" & theindex & "DescH").Item,1,100))	' Nullable: YES Type: nvarchar
				End If
			Else
				rs("TenantPK") = Null
				rs("TenantID") = Null
				rs("TenantName") = Null
			End If
		End If
	End If

	If Not IsEmpty(Request.Form(fp & "Project" & theindex & "PK").Item) Then
		If Not Request.Form(fp & "Project" & theindex & "PK").Item = "UNCHANGED" Then
			If Len(Trim(Request.Form(fp & "Project" & theindex & "PK").Item)) > 0 Then
				rs("ProjectPK") = Request.Form(fp & "Project" & theindex & "PK").Item	' Nullable: YES Type: int
				If Not IsEmpty(Request.Form(fp & "Project" & theindex).Item) Then
					rs("ProjectID") = Trim(Mid(Request.Form(fp & "Project" & theindex).Item,1,25))	' Nullable: YES Type: nvarchar
				End If
				If Not IsEmpty(Request.Form(fp & "Project" & theindex & "DescH").Item) Then
					rs("ProjectName") = Trim(Mid(Request.Form(fp & "Project" & theindex & "DescH").Item,1,50))	' Nullable: YES Type: nvarchar
				End If
			Else
				rs("ProjectPK") = Null
				rs("ProjectID") = Null
				rs("ProjectName") = Null
			End If
		End If
	End If

	' --------------- Customized ------------------
	'rs("ShutdownBox") = (InStr(Request.Form(fp & "ShutdownBox" & theindex).Item,"taskchecked.gif") > 0) ' Nullable: No Type: bit
	' ---------------------------------------------

	' -- End Table Fields ------------------------------------------------

	rs("ScheduleDisabled") = (InStr(Request.Form(fp & "ScheduleDisabled" & theindex).Item,"taskline.gif") > 0) ' Nullable: No Type: bit

	Call db_version(rs)

End Sub

Sub db_al(rs,isinsert,suffix,htmltable,theindex)

	Dim fp
	fp = "txt" & LCase(htmltable) & suffix

	' -- Start Table Fields ------------------------------------------------

	rs("AssetPK") = keyvalue

	If Len(Trim(Request.Form(fp & "Labor" & theindex & "PK").Item)) > 0 Then
		rs("LaborPK") = Request.Form(fp & "Labor" & theindex & "PK").Item	' Nullable: No Type: int
	End If
	If Not IsEmpty(Request.Form(fp & "Priority" & theindex).Item) Then
		If Len(Trim(Request.Form(fp & "Priority" & theindex).Item)) > 0 Then
			rs("Priority") = Request.Form(fp & "Priority" & theindex).Item	' Nullable: YES Type: smallint
		End If
	End If
	rs("AutoAssign") = (InStr(Request.Form(fp & "AutoAssign" & theindex).Item,"taskchecked.gif") > 0) ' Nullable: Yes Type: bit
	rs("BackupResource") = (InStr(Request.Form(fp & "BackupResource" & theindex).Item,"taskchecked.gif") > 0) ' Nullable: Yes Type: bit

	' -- End Table Fields ------------------------------------------------

	Call db_version(rs)

End Sub

Sub db_ao(rs,isinsert,suffix,htmltable,theindex)

	Dim fp
	fp = "txt" & LCase(htmltable) & suffix

	' -- Start Table Fields ------------------------------------------------

	rs("AssetPK") = keyvalue

	If Len(Trim(Request.Form(fp & "Company" & theindex & "PK").Item)) > 0 Then
		rs("CompanyPK") = Request.Form(fp & "Company" & theindex & "PK").Item	' Nullable: No Type: int
	End If
	If Not IsEmpty(Request.Form(fp & "PeriodStart" & theindex).Item) Then
		If Not Request.Form(fp & "PeriodStart" & theindex).Item = "" Then
			rs("PeriodStart") = SQLdatetimeADO(Request.Form(fp & "PeriodStart" & theindex).Item)	' Nullable: YES Type: datetime
		Else
			rs("PeriodStart") =	Null
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "PeriodEnd" & theindex).Item) Then
		If Not Request.Form(fp & "PeriodEnd" & theindex).Item = "" Then
			rs("PeriodEnd") = SQLdatetimeADO(Request.Form(fp & "PeriodEnd" & theindex).Item)	' Nullable: YES Type: datetime
		Else
			rs("PeriodEnd") = Null
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "VendorContractNum" & theindex).Item) Then
		If Len(Request.Form(fp & "VendorContractNum" & theindex).Item) > 0 Then
			rs("VendorContractNum") = Trim(Mid(Request.Form(fp & "VendorContractNum" & theindex).Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("VendorContractNum") = Null
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "ContractSummary" & theindex).Item) Then
		If Len(Request.Form(fp & "ContractSummary" & theindex).Item) > 0 Then
			rs("ContractSummary") = Trim(Mid(Request.Form(fp & "ContractSummary" & theindex).Item,1,7900))	' Nullable: YES Type: nvarchar
		Else
			rs("ContractSummary") = Null
		End If
	End If

	' -- End Table Fields ------------------------------------------------

	Call db_version(rs)

End Sub

Sub db_ar(rs,isinsert,suffix,htmltable,theindex)

	Dim fp
	fp = "txt" & LCase(htmltable) & suffix

	' -- Start Table Fields ------------------------------------------------

	rs("AssetPK") = keyvalue

	If Len(Trim(Request.Form(fp & "Labor" & theindex & "PK").Item)) > 0 Then
		rs("LaborPK") = Request.Form(fp & "Labor" & theindex & "PK").Item	' Nullable: No Type: int
	End If
	If Not IsEmpty(Request.Form(fp & "StartDate" & theindex).Item) Then
		If Not Request.Form(fp & "StartDate" & theindex).Item = "" Then
			rs("StartDate") = SQLdatetimeADO(Request.Form(fp & "StartDate" & theindex).Item)	' Nullable: YES Type: datetime
		Else
			rs("StartDate") = Null
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "EndDate" & theindex).Item) Then
		If Not Request.Form(fp & "EndDate" & theindex).Item = "" Then
			rs("EndDate") = SQLdatetimeADO(Request.Form(fp & "EndDate" & theindex).Item)	' Nullable: YES Type: datetime
		Else
			rs("EndDate") = Null
		End If
	End If
	' --------------- Customized ------------------
	rs("IsLease") = (InStr(Request.Form(fp & "IsLease" & theindex).Item,"taskchecked.gif") > 0) ' Nullable: No Type: bit
	' ---------------------------------------------
	If Not IsEmpty(Request.Form(fp & "ExpireDate" & theindex).Item) Then
		If Not Request.Form(fp & "ExpireDate" & theindex).Item = "" Then
			rs("ExpireDate") = Request.Form(fp & "ExpireDate" & theindex).Item	' Nullable: YES Type: datetime
		Else
			rs("ExpireDate") = Null
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "Percentage" & theindex).Item) Then
		If Len(Trim(Request.Form(fp & "Percentage" & theindex).Item)) > 0 Then
			rs("Percentage") = Request.Form(fp & "Percentage" & theindex).Item	' Nullable: YES Type: datetime
		End If
	End If
	' --------------- Customized ------------------
	rs("Active") = (InStr(Request.Form(fp & "Active" & theindex).Item,"taskchecked.gif") > 0) ' Nullable: No Type: bit
	' ---------------------------------------------
	' --------------- Customized ------------------
	rs("IsPrimary") = (InStr(Request.Form(fp & "IsPrimary" & theindex).Item,"taskchecked.gif") > 0) ' Nullable: No Type: bit
	' ---------------------------------------------

	' -- End Table Fields ------------------------------------------------

	Call db_version(rs)

End Sub

Sub db_at(rs,isinsert,suffix,htmltable,theindex)

	Dim fp
	fp = "txt" & LCase(htmltable) & suffix

	' -- Start Table Fields ------------------------------------------------

	rs("AssetPK") = keyvalue

	If Len(Trim(Request.Form(fp & "Document" & theindex & "PK").Item)) > 0 Then
		rs("DocumentPK") = Request.Form(fp & "Document" & theindex & "PK").Item	' Nullable: No Type: int
	End If
	' --------------- Customized ------------------
	If Not IsEmpty(Request.Form(fp & "ModuleID" & theindex).Item) Then
		rs("ModuleID") = Request.Form(fp & "ModuleID" & theindex).Item	' Nullable: No Type: nchar
	End If
	' ---------------------------------------------
	rs("DisplayLink") = (InStr(Request.Form(fp & "DisplayLink" & theindex),"taskchecked.gif") > 0) ' Nullable: No Type: bit
	rs("PrintWithWO") = (InStr(Request.Form(fp & "PrintWithWO" & theindex).Item,"taskchecked.gif") > 0) ' Nullable: No Type: bit
	rs("SendWithEmail") = (InStr(Request.Form(fp & "SendWithEmail" & theindex).Item,"taskchecked.gif") > 0)	' Nullable: No Type: bit
	' ---------------------------------------------

	' -- End Table Fields ------------------------------------------------

	Call db_version(rs)

End Sub

Sub db_no(rs,isinsert,suffix,htmltable,theindex)

	Dim fp
	fp = "txt" & LCase(htmltable) & suffix

	' -- Start Table Fields ------------------------------------------------

	rs("AssetPK") = keyvalue

	If IsDate(Request.Form(fp & "NoteDate" & theindex)) Then
		If Not IsEmpty(Request.Form(fp & "NoteTime" & theindex)) Then
			rs("NoteDate") = SQLdatetimeADOAT(Request.Form(fp & "NoteDate" & theindex) & " " & Request.Form(fp & "NoteTime" & theindex))	' Nullable: No Type: datetime
		Else
			rs("NoteDate") = SQLdatetimeADO(Request.Form(fp & "NoteDate" & theindex) & " " & Time())	' Nullable: No Type: datetime
		End If
	Else
		rs("NoteDate") = SQLdatetimeADO(DateTimeNullCheck(Now()))	' Nullable: No Type: datetime
	End If
	If Not IsEmpty(Request.Form(fp & "Initials" & theindex)) Then
		If Len(Request.Form(fp & "Initials" & theindex)) > 0 Then
			rs("Initials") = Trim(Mid(Request.Form(fp & "Initials" & theindex),1,5))	' Nullable: YES Type: varchar
		Else
			rs("Initials") = Null
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "Note" & theindex)) Then
		If Len(Request.Form(fp & "Note" & theindex)) > 0 Then
			rs("Note") = Trim(Mid(Request.Form(fp & "Note" & theindex),1,7500))	' Nullable: YES Type: varchar
		Else
			rs("Note") = Null
		End If
	End If

	' ---------------------------------------------
	rs("Custom1") = (InStr(Request.Form(fp & "Custom1" & theindex),"taskchecked.gif") > 0) ' Nullable: No Type: bit
	rs("Custom2") = (InStr(Request.Form(fp & "Custom2" & theindex),"taskchecked.gif") > 0)	' Nullable: No Type: bit
	' ---------------------------------------------

	' -- End Table Fields ------------------------------------------------

	Call db_version(rs)

End Sub
	%>
