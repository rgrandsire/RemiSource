<%@ EnableSessionState=False Language=VBScript %>
<% Option Explicit %>

<!--#INCLUDE FILE="../common/mc_all_cache.asp" --> 
<!--#INCLUDE FILE="../common/mc_tabcontrol.asp" -->
    
<%
Dim keyvalue,keyvalues,poclass,podefault,lastaction,errorfield,errortabinfo,returnmessage,returnclass,newrecord,duprecord,mcmode,findtabtext,treefiltervalue,norecord,firstload,curtab
Dim ServerDate,addontheflymode,addontheflymodule,currentmodule,overrideApprovalValue

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
		poclass = Trim(Request.QueryString("poclass"))
		podefault = Trim(Request.QueryString("podefault"))
	Else
		poclass = "1"
		podefault = ""
	End If

	' This is NOT the SQL Server Date - rather the Web Server Date! (for now)
	ServerDate = DateTimeNullCheck(Now())

End Sub

Sub displayhtml

	' ModuleSpecific
	Dim windowonloadjs, headercmds
	headercmds = _
	"<script language=""javascript"" src=""../../javascript/normal/mc_tabcontrol.js""></script>" & _
	"<script language=""javascript"" src=""../../javascript/normal/mc_autocomplete.js""></script>" & _
	"<link rel=""stylesheet"" type=""text/css"" href=""../../css/mc_tabstyle.css"">"
	windowonloadjs = ""
	'windowonloadjs = windowonloadjs + "top.fraTopic.document.body.style.backgroundColor = '#808080';" + nl
	'windowonloadjs = windowonloadjs + "top.fraTopic.document.body.style.backgroundImage = '';" + nl
	windowonloadjs = windowonloadjs + "top.fraTabbar.document.images.maxmin.style.marginRight='0';" + nl
	windowonloadjs = windowonloadjs + "top.fraTabbar.document.images.closemod.style.display='';" + nl
	windowonloadjs = windowonloadjs + "pagetabs_current = self.document.getElementById('pagetabs_21');" + nl
	windowonloadjs = windowonloadjs + "top.endprocess();" + nl
	windowonloadjs = windowonloadjs + "top.navcurrent(null,null,true);" + nl

    ' Purchase Order AutoComplete Fields
    windowonloadjs = windowonloadjs & "try {"
    windowonloadjs = windowonloadjs & "if (top.acom == true) {"

    windowonloadjs = windowonloadjs & "var txtVendor_AC = new actb('CM','PO',document.mcform.txtVendor);"
    windowonloadjs = windowonloadjs & "var txtBuyer_AC = new actb('LA_TAKENBY','PO',document.mcform.txtBuyer);"
    windowonloadjs = windowonloadjs & "var txtBuyerCompany_AC = new actb('CM','PO',document.mcform.txtBuyerCompany);"
    windowonloadjs = windowonloadjs & "var txtRequester_AC = new actb('LA_REQUESTER','PO',document.mcform.txtRequester);"
    windowonloadjs = windowonloadjs & "var txtTenant_AC = new actb('TN','PO',document.mcform.txtTenant);"
    windowonloadjs = windowonloadjs & "var txtAccount_AC = new actb('AC','PO',document.mcform.txtAccount);"
    windowonloadjs = windowonloadjs & "var txtDepartment_AC = new actb('DP','PO',document.mcform.txtDepartment);"
    windowonloadjs = windowonloadjs & "var txtRepairCenter_AC = new actb('RC','PO',document.mcform.txtRepairCenter);"
    windowonloadjs = windowonloadjs & "var txtFreightTerms_AC = new actb('LOOKUP_SHIPTERM','PO',document.mcform.txtFreightTerms);"
    windowonloadjs = windowonloadjs & "var txtShippingMethod_AC = new actb('LOOKUP_SHIPMETHOD','PO',document.mcform.txtShippingMethod);"
    windowonloadjs = windowonloadjs & "var txtShipTo_AC = new actb('CM','PO',document.mcform.txtShipTo);"
    windowonloadjs = windowonloadjs & "var txtTerms_AC = new actb('LOOKUP_paymentterms','PO',document.mcform.txtTerms);"
    windowonloadjs = windowonloadjs & "var txtBillTo_AC = new actb('CM','PO',document.mcform.txtBillTo);"
    windowonloadjs = windowonloadjs & "var txtCurrency_AC = new actb('LOOKUP_currency','PO',document.mcform.txtCurrency);"
    windowonloadjs = windowonloadjs & "var txtPriority_AC = new actb('LOOKUP_popriority','PO',document.mcform.txtPriority);"
    windowonloadjs = windowonloadjs & "try {var txtSubstatus_AC = new actb('LOOKUP_POSUBSTATUS','PO',document.mcform.txtSubStatus);} catch(e) { }"

    windowonloadjs = windowonloadjs & "}"
    windowonloadjs = windowonloadjs & "} catch(e) {};"

	Call domctop("",True,headercmds,True)
	FlushIt
	%>

	<div align="right" id="fakewomenu" style="display:none; MARGIN-TOP: 0px;" height="0" width="100%" onclick="top.showMenu('pomenu',top.recordkey,true,top.fraTopic,'FORCETOPRIGHT');"></div>
	<div id="mcpage" name="purchaseorder" moduleid="PO" allownewrecords="Y">

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
					<div id="splash" name="tab61" STYLE="display: none;" mcfocus="NONE" watermark="images/po_watermark.jpg">
					<!--#INCLUDE FILE="purchaseorder_splash.htm" -->
					</div>

					<!-- Tab #1  -->
					<div id="po_details" name="tab61" dataloaded="Y" STYLE="display: none;" showheader="N" multipage="Y">
					<!--#INCLUDE FILE="_purchaseorder_details_Sandvik.htm" -->
					</div>

					<% FlushIt %>

					<!-- Tab #2  -->
					<div id="po_lineitems" name="tab62" dataloaded="N" STYLE="display: none;" mcfocus="NONE" showheader="Y">
					<!--#INCLUDE FILE="purchaseorder_lineitems.htm" -->
					</div>

					<!-- Tab #3  -->
					<div id="po_receipts" name="tab63" dataloaded="Y" STYLE="display: none;" mcfocus="NONE" showscroll="N" showheader="Y" noaccessinedit="Y">
					<!--#INCLUDE FILE="purchaseorder_receipts.htm" -->
					</div>

					<!-- Tab #4  -->
					<div id="po_rma" name="tab64" dataloaded="Y" STYLE="display: none;" mcfocus="NONE" showscroll="N" showheader="Y" noaccessinedit="Y">
					<!--#INCLUDE FILE="purchaseorder_rma.htm" -->
					</div>

					<% FlushIt %>

					<!-- Tab #5  -->
					<div id="po_attach" name="tab65" dataloaded="N" STYLE="display: none;" showheader="Y" mcfocus="NONE" watermark="images/attach_watermark2.jpg">
					<!--#INCLUDE FILE="purchaseorder_attach.htm" -->
					</div>

					<!-- Tab #6  -->
					<div id="po_report" name="tab66" dataloaded="Y" STYLE="display: none;" mcfocus="NONE" showscroll="N" showheader="N" noaccessinedit="Y">
					<!--#INCLUDE FILE="purchaseorder_report.htm" -->
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
					<% Call OutputBottomTabs(2) %>		            
		            </td>
		            <td width="65%" valign="top" align="right" style="padding-top:3px;white-space:nowrap;">

						<span id="bg_requested" class="headerbuttongrp" style="display:none;">
                            <% Call ButtonGenerate("ISSUE","top.MODclickMenu('mnuPOInitiate');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("DENY","top.MODclickMenu('mnuPODeny');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("CLONE","","") %>
                            <% Call ButtonGenerate("DELETE","top.MODclickMenu('mnuPODelete');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("HISTORY","","") %>
                            <% Call ButtonGenerate("PRINT2","if (self.checkifpocanprint()) {top.printporeport();event.cancelBubble=true;}","") %>
                        </span>
						<span id="bg_generated" class="headerbuttongrp" style="display:none;">
                            <% Call ButtonGenerate("ISSUE","top.MODclickMenu('mnuPOInitiate');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("CANCEL2","top.MODclickMenu('mnuPOCancel');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("CLONE","","") %>
                            <% Call ButtonGenerate("DELETE","top.MODclickMenu('mnuPODelete');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("HISTORY","","") %>
                            <% Call ButtonGenerate("PRINT2","if (self.checkifpocanprint()) {top.printporeport();event.cancelBubble=true;}","") %>
                        </span>
						<span id="bg_issued" class="headerbuttongrp" style="display:none;">
                            <% Call ButtonGenerate("CLOSE","top.MODclickMenu('mnuPOClose');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("CLONE","","") %>
                            <% Call ButtonGenerate("DELETE","top.MODclickMenu('mnuPODelete');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("HISTORY","","") %>
                            <% Call ButtonGenerate("PRINT2","if (self.checkifpocanprint()) {top.printporeport();event.cancelBubble=true;}","") %>
                        </span>
						<span id="bg_onhold" class="headerbuttongrp" style="display:none;">
                            <% Call ButtonGenerate("REISSUE","top.MODclickMenu('mnuPOInitiate');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("CLOSE","top.MODclickMenu('mnuPOClose');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("CANCEL2","top.MODclickMenu('mnuPOCancel');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("CLONE","","") %>
                            <% Call ButtonGenerate("DELETE","top.MODclickMenu('mnuPODelete');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("HISTORY","","") %>
                            <% Call ButtonGenerate("PRINT2","if (self.checkifpocanprint()) {top.printporeport();event.cancelBubble=true;}","") %>
                        </span>
						<span id="bg_denied" class="headerbuttongrp" style="display:none;margin-left:0px;">
                            <% Call ButtonGenerate("REISSUE","top.MODclickMenu('mnuPOInitiate');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("CLONE","","") %>
                            <% Call ButtonGenerate("DELETE","top.MODclickMenu('mnuPODelete');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("HISTORY","","") %>
                            <% Call ButtonGenerate("PRINT2","if (self.checkifpocanprint()) {top.printporeport();event.cancelBubble=true;}","") %>
                        </span>
						<span id="bg_closed" class="headerbuttongrp" style="display:none;margin-left:0px">
                            <% Call ButtonGenerate("REISSUE","top.MODclickMenu('mnuPOInitiate');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("CLONE","","") %>
                            <% Call ButtonGenerate("DELETE","top.MODclickMenu('mnuPODelete');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("HISTORY","","") %>
                            <% Call ButtonGenerate("PRINT2","if (self.checkifpocanprint()) {top.printporeport();event.cancelBubble=true;}","") %>
                        </span>
						<span id="bg_canceled" class="headerbuttongrp" style="display:none;margin-left:0px">
                            <% Call ButtonGenerate("REISSUE","top.MODclickMenu('mnuPOInitiate');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("CLONE","","") %>
                            <% Call ButtonGenerate("DELETE","top.MODclickMenu('mnuPODelete');event.cancelBubble=true;","") %>
                            <% Call ButtonGenerate("HISTORY","","") %>
                            <% Call ButtonGenerate("PRINT2","if (self.checkifpocanprint()) {top.printporeport();event.cancelBubble=true;}","") %>
						</span>

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

		    function ChangeAuthClick(inType) {
		        if (inType == '1') {
		            $('.authOption').attr('id', 'mnuPOAuthApprove');
		        }
		        else {
		            $('.authOption').attr('id', 'mnuPOAuthApproveByRC');
		        }
		        //top.MODclickMenu('poauthmenu', top.fraTopic)
		    }
		</script>

		<!-- Purchase Order Status Menu -->

		<div id="pomenu" onclick="top.MODclickMenu('pomenu',top.fraTopic)" onmouseover="top.toggleMenu(top.fraTopic)" onmouseout="top.toggleMenu(top.fraTopic)" oncontextmenu="return false;" style="position:absolute;display:none; background-Color:#FEFEFE; width:140px; border: 2px outset #FFFFFF; padding-top:5px; padding-bottom:5px;background-image: url('../../images/menuleftbg5.gif');">
		  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/status_inprogress_g.gif" style="margin-left:2px;margin-right:9px;" WIDTH="16" HEIGHT="13"></td><td class="menuItemImg" id="mnuPOInitiate">Issue</td></tr></table>
		  <div class="menuItemHR"></div>
		  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/status_no_g.gif" style="margin-left:2px;margin-right:9px;" WIDTH="16" HEIGHT="13"></td><td class="menuItemImg" id="mnuPODeny">Deny</td></tr></table>
		  <div class="menuItemHR"></div>
		  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/status_waiting_g.gif" style="margin-left:2px;margin-right:9px;" WIDTH="16" HEIGHT="13"></td><td class="menuItemImg" id="mnuPOPending">On-Hold</td></tr></table>
		  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/status_canceled_g.gif" style="margin-left:2px;margin-right:9px;" WIDTH="16" HEIGHT="13"></td><td class="menuItemImg" id="mnuPOCancel">Cancel</td></tr></table>
		  <div class="menuItemHR"></div>
		  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/status_approvalrequired_g.gif" style="margin-left:2px;margin-right:9px;" WIDTH="16" HEIGHT="13"></td><td class="menuItemImg" id="mnuPOClose">Close</td></tr></table>
		  <div class="menuItemHR"></div>
		  <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/mnu_delete_g.gif" style="margin-left:3px;margin-right:9px;" WIDTH="14" HEIGHT="13"></td><td class="menuItemImg" id="mnuPODelete">Delete</td></tr></table>
		  <div class="menuItemHR"></div>
		  <div class="menuItem" id="mnuCloseMenu">Close Menu</div>
		</div>

		<!-- Purchase Order Authorization Menu -->
		<div id="poauthmenu" onclick="top.MODclickMenu('poauthmenu', top.fraTopic);" onmouseover="top.toggleMenu(top.fraTopic)" onmouseout="top.toggleMenu(top.fraTopic)" oncontextmenu="return false;" style="position:absolute;display:none; background-Color:#FEFEFE; width:140px; border: 2px outset #FFFFFF; padding-top:5px; padding-bottom:5px;background-image: url('../../images/menuleftbg5.gif');">
          <div id="canAuthMenu" style="display:;">
		      <table id='approvalNormal' cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/status_closed_g.gif" style="margin-left:2px;margin-right:9px;" WIDTH="16" HEIGHT="13"></td><td class="menuItemImg authOption" id="mnuPOAuthApprove">Approve</td></tr></table>
		      <div class="menuItemHR"></div>
		      <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/status_none_g.gif" style="margin-left:2px;margin-right:9px;" WIDTH="16" HEIGHT="13"></td><td class="menuItemImg" id="mnuPOAuthUnapprove">Unapprove</td></tr></table>
		      <div class="menuItemHR"></div>
		      <table cellpadding="0" cellspacing="0"><tr><td><img src="../../images/icons/status_canceled_g.gif" style="margin-left:2px;margin-right:9px;" WIDTH="16" HEIGHT="13"></td><td class="menuItemImg" id="mnuPOAuthReject">Reject</td></tr></table>
          </div>
          <div class="menuItemHR approvalHide" id="mnuLine"><hr size="1" class="menuhr"></div>
		  <div class="menuItem" id="mnuCloseMenu">Close Menu</div>

		</div>

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

Function GetRequiredStatus(db, ipopk, ircpk)
    Dim iRow, POApprovalType, ApprovalLevel, AmountA, AmountB, Condition, tmpTotalCost,poAuthStatus,poICount,rString
    Dim deString
    deString = ""
    Set iRow = db.RunSQLReturnRS("SELECT TOP 1 POApprovalType FROM RepairCenter WITH ( NOLOCK ) WHERE RepairCenterPK=" & FixInt(ircpk), "")
    If Not iRow.EOF Then
        POApprovalType = iRow("POApprovalType")
    Else
        POApprovalType = 1
    End If
    Set iRow = Nothing

    Set iRow = db.RunSQLReturnRS("SELECT TOP 1 Total FROM PurchaseOrder WITH ( NOLOCK ) WHERE POPK = " & FixInt(ipopk), "")
    If Not iRow.EOF Then
        tmpTotalCost = iRow(0)
    Else
        tmpTotalCost = 0
    End If
    Set iRow = Nothing
    rString = Request("txtAuthStatus")
    Response.Write("<script type='text/javascript'>")
    If Clng(POApprovalType) = 0 Then
			Response.Write("top.addpoauth('NOTREQUIRED');")+nl
		    Response.Write("	top.sethidden('txtAuthStatus','NOTREQUIRED');")+nl
            rString = "NOTREQUIRED"
            deString = "0: NOTREQUIRED"
    ElseIf Clng(POApprovalType) = 2 Then
        Select Case UCase(poAuthStatus)
            Case "REQUIRED1"
			    Response.Write("top.addpoauth('REQUIRED1');")+nl
		        Response.Write("	top.sethidden('txtAuthStatus','REQUIRED1');")+nl
                rString = "REQUIRED1"
                deString = "2: REQUIRED1"
            Case "REQUIRED2"
			    Response.Write("top.addpoauth('REQUIRED2');")+nl
		        Response.Write("	top.sethidden('txtAuthStatus','REQUIRED2');")+nl
                rString = "REQUIRED2"
                deString = "2: REQUIRED2"
            Case "REQUIRED3"
			    Response.Write("top.addpoauth('REQUIRED3');")+nl
		        Response.Write("	top.sethidden('txtAuthStatus','REQUIRED3');")+nl
                rString = "REQUIRED3"
                deString = "2: REQUIRED3"
            Case "REQUIRED4"
			    Response.Write("top.addpoauth('REQUIRED4');")+nl
		        Response.Write("	top.sethidden('txtAuthStatus','REQUIRED4');")+nl
                rString = "REQUIRED4"
                deString = "2: REQUIRED4"
            Case "REQUIRED5"
			    Response.Write("top.addpoauth('REQUIRED5');")+nl
		        Response.Write("	top.sethidden('txtAuthStatus','REQUIRED5');")+nl
                rString = "REQUIRED5"
                deString = "2: REQUIRED5"
        End Select
    Else
		If GetSession("POAuthReq") = "0" Then
			Response.Write("top.addpoauth('NOTREQUIRED');")+nl
		    Response.Write("	top.sethidden('txtAuthStatus','NOTREQUIRED');")+nl
            rString = "NOTREQUIRED"
            deString = "1: NOTREQUIRED"
		Else
			Response.Write("top.addpoauth('REQUIRED" & GetSession("POAuthReq") & "');")+nl
		    Response.Write("	top.sethidden('txtAuthStatus','REQUIRED" & GetSession("POAuthReq") & "');")+nl
            rString = "REQUIRED" & GetSession("POAuthReq")
            deString = "1: REQUIRED" & GetSession("POAuthReq")
		End If
    End If
    Response.Write("</script>")
    GetRequiredStatus = rString
End Function

Sub DoApprovals(db)
    Dim iRow, POApprovalType, ApprovalLevel, AmountA, AmountB, Condition, tmpTotalCost,poAuthStatus,poICount

    Set iRow = db.RunSQLReturnRS("SELECT TOP 1 POApprovalType FROM RepairCenter WITH ( NOLOCK ) WHERE RepairCenterPK=" & FixInt(Request("txtRepairCenterPK")), "")
    If Not iRow.EOF Then
        POApprovalType = iRow("POApprovalType")
    Else
        POApprovalType = 1
    End If
    Set iRow = Nothing

    Set iRow = db.RunSQLReturnRS("SELECT TOP 1 Total FROM PurchaseOrder WITH ( NOLOCK ) WHERE POPK = " & keyvalue, "")
    If Not iRow.EOF Then
        tmpTotalCost = iRow(0)
    Else
        tmpTotalCost = 0
    End If
    Set iRow = Nothing
    If POApprovalType = 0 Then
        Call db.RunSQL("SET NOCOUNT ON UPDATE PurchaseOrder WITH ( ROWLOCK ) SET AuthStatus = 'NOTREQUIRED', AuthStatusDesc = '(Not Required)', AuthLevelsRequired = 0 WHERE POPK = " & keyvalue & " SET NOCOUNT OFF","")
    End If
    If POApprovalType = 2 Then
        dim poAuthStatusDesc,poAuthLevelsRequired
        poICount = 1
        poAuthStatus = "NOTREQUIRED"
        Set iRow = db.RunSQLReturnRS("SELECT TOP 5 ApprovalLevel, AmountA, AmountB, Condition FROM RepairCenterApproval WITH ( NOLOCK ) WHERE ModuleID = 'PO' AND RepairCenterPK = " & FixInt(Request("txtRepairCenterPK")) & " ORDER BY ApprovalLevel", "")
        Do While Not iRow.EOF
        If CSng(tmpTotalCost) < CSng("0.00") Then
            poAuthStatusDesc = "(Not Required)"
            poAuthLevelsRequired=0
        Else
            If CLng(iRow("ApprovalLevel")) = 1 Then
		        If CSng(FixSingle(tmpTotalCost)) < CSng(FixSingle(iRow("AmountA"))) Then
                    poAuthStatus = "NOTREQUIRED"
                    poAuthStatusDesc = "(Not Required)"
                    poAuthLevelsRequired=0
                ElseIf  CSng(tmpTotalCost) < CSng("0.00") Then
                Else
		            If CSng(FixSingle(tmpTotalCost)) >= CSng(FixSingle(iRow("AmountA"))) Then
				            poAuthStatus = "REQUIRED1"
                            poAuthStatusDesc = "Required - L1"
                            poAuthLevelsRequired="1"
		            End IF
                End If
            Else
		        If CSng(FixSingle(tmpTotalCost)) >= CSng(FixSingle(iRow("AmountA"))) Then
                    If poICount = 1 Then
				        poAuthStatus = "REQUIRED" & CStr(poICount)
                        poAuthStatusDesc = "Required - L" & CStr(poICount)
                        poAuthLevelsRequired=CStr(poICount)
                    Else
                        If FixInt(iRow("AmountA")) <> 0 Then
					        poAuthStatus = "REQUIRED" & CStr(poICount)
                            poAuthStatusDesc = "Required - L" & CStr(poICount)
                            poAuthLevelsRequired=CStr(poICount)
                        End If
                    End If
		        End IF
            End If
        End IF
            poICount = poICount + 1
            iRow.MoveNext
        Loop
        Set iRow = Nothing

        'If poAuthStatus = "NOTREQUIRED" Then
        '    poAuthStatusDesc = "(Not Required)"
        '    poAuthLevelsRequired=0
        'End If
        'Call db.RunSQL("SET NOCOUNT ON UPDATE PurchaseOrder WITH ( ROWLOCK ) SET AuthStatus = '" & poAuthStatus & "', AuthStatusDesc = '" & poAuthStatusDesc & "', AuthLevelsRequired = " & poAuthLevelsRequired & " WHERE POPK = " & keyvalue & " SET NOCOUNT OFF","")

        'MCA Fix Thanks MCA! :)
        Dim POSql 
        POSql = "SET NOCOUNT ON UPDATE PurchaseOrder WITH ( ROWLOCK ) SET AuthStatus = '" & poAuthStatus & "', AuthStatusDesc = '" & poAuthStatusDesc & "', AuthLevelsRequired = " & poAuthLevelsRequired
 
        If poAuthStatus = "NOTREQUIRED" Then
            poAuthStatusDesc = "(Not Required)"
            poAuthLevelsRequired=0
            POSql = POSql + ", isApproved=1 "
        End If
        Call db.RunSQL(POSql & " WHERE POPK = " & keyvalue & " SET NOCOUNT OFF","")

    End If
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
			Dim atcwhere,licwhere,nocwhere
			Dim atnew,linew,nonew
			Dim atdwhere,lidwhere,nodwhere

			' <SERVER-SIDE VALIDATION STUFF HERE>

			aok = True
			errorfield = ""
			returnmessage = ""
			returnclass = "standardmessage"
			Set db = New ADOHelper

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


			If db.OpenClientConnection Then
				If db.OpenTransaction Then

					' <SAVE THE MAIN RECORD>

         			Call db_master(db)

					' <SAVE THE CHILDREN RECORDS>

					Call db_child(db,d,"LI","purchaseorderdetail",licwhere,linew,lidwhere)
					Call db_child(db,d,"AT","purchaseorderdocument",atcwhere,atnew,atdwhere)
					Call db_child(db,d,"NO","purchaseordernote",nocwhere,nonew,nodwhere)

					db.CloseTransaction

				End If

                'CB2->Approval Code
                Call DoApprovals(db)
			End If

			dok = db.dok
			derror = db.derror
			If dok Then
				If newrecord Then
                    Dim returnString
                    returnString = GetRequiredStatus(db, keyvalue, FixInt(Request("txtRepairCenterPK")))

					Call db.RunSP("MC_ProcessNewPurchaseOrder",Array(_
					Array("@POPK", adInteger, adParamInput, 4, keyvalue),_
					Array("@status",  MC_ADVARCHAR, adParamInput, 25, Trim(Mid(Request.Form("txtStatus"),1,15))),_
					Array("@oldstatus",  MC_ADVARCHAR, adParamInput, 25, Null),_
					Array("@authstatus",  MC_ADVARCHAR, adParamInput, 25, Trim(returnString))_
					),"")
					dok = db.dok
					derror = db.derror
				End If

			End If

			Set d = Nothing

            'CB2->Approvals
            'Check repair center, approval amounts etc.

			' <RETURN TO CLIENT HERE>


            If dok Then

				If addontheflymode = "Y" Then
					newjs = newjs + "	top.addonthefly_recupdated = true;" + nl
					If newrecord Then
						newjs = newjs & "	top.addonthefly_recaddedpk = '" & CStr(keyvalue) & "';"
					End If
				End If

				newjs = newjs + "	top.mcmode = 'EDIT';" + nl
				newjs = newjs + "	top.dirtydata = false;" + nl
				newjs = newjs + "	top.requiredfields('OFF');" + nl

				returnmessage = "The Purchase Order has been Saved"
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
				If RefreshExplorerIsTurnedOn(db,currentmodule,newrecord) Then
					Response.Write("top.refreshcurrentexplorer(true);")+nl
				End If
				'======================================================================================================================================
                'Response.Write("alert('" & GetApprovalType(db, Request("txtRepairCenterPK")) & "'); ") & nl
				db.CloseClientConnection


				Call dogenericendaction(True)
				Response.Write "	top.showtabinfo(top.currenttabinfo.id,false);" + nl
				If newrecord Then
				Response.Write "	top.navnewcomplete(top.currentmodule,'kv=" + CStr(keyvalue) + "');" + nl
				End If
				If db.warn Then
				Response.Write("   top.mcalert('info','Purchase Order Message','The Purchase Order has been saved successfully, but there were some warnings that you need to be aware of. The details of the warnings are described below.<br><br><u>Warning Details</u>:<br><br>" & Replace(db.warntext,"'","\'") & "','bg_okprint',700,370,'sounds/error.wav');") + nl
				Else
				Response.Write("	top.playsound('sounds/done.wav');")
				End If
				Response.Write "	top.dofocus();" + nl
			    Response.Write "	 try{myframe.ForceOverride(top.recordkey.replace('kv=',''));}catch(e){}" + nl
				dogenericfooter
				Set db = Nothing
				Response.End

			Else
				If Not dok Then
					'If db.isduplicate Then
					'	newjs = newjs + "   top.mcalert('warning','Purchase Order Message','You have entered a Purchase Order ID that is already in use. Please change the Purchase Order ID and then click the SAVE button.','bg_okprint',700,240,'sounds/error.wav');" + nl
					'Else
						newjs = newjs + "   top.mcalert('warning','Purchase Order Message','There was a problem saving the Purchase Order. The details of the problem are described below. You can try to SAVE the Purchase Order again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.<br><br><u>Problem Details</u>:<br><br>" & Replace(derror,"'","\'") & "','bg_okprint',700,370,'sounds/error.wav');" + nl
					'End If
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

	Dim db,OutArray,kvs,sql,dok,doaction,newjs,recinview,selrecs,criteria,ecriteria,showbox,actionwhere
	Dim POPK,POID,POName
	Dim RowUser, RowInitials

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
		kvs = Trim(UCase(Request.QueryString("kvs")))
		rowuser = Request.QueryString("txtRowVersionUserPK")
		rowinitials = Trim(Mid(Request.QueryString("txtRowVersionInitials").Item,1,5))
	Else
		lastaction = Trim(Request.Form("lastaction"))
		recinview = Trim(UCase(Request.Form("recinview")))
		selrecs = Trim(Request.Form("sel"))
		criteria = Trim(Request.Form("findercrit"))
		ecriteria = Trim(Request.Form("finderecrit"))
		showbox = Trim(UCase(Request.Form("showbox")))
		rowuser = Request.Form("txtRowVersionUserPK")
		rowinitials = Trim(Mid(Request.Form("txtRowVersionInitials").Item,1,5))
	End If

	If rowuser = "" Then
		rowuser = GetSession("USERPK")
	End If
	If rowinitials = "" Then
		rowinitials = GetSession("USERINITIALS")
	End If

	newjs = ""

	Set db = New ADOHelper

	Select Case doaction

		Case "PENDING"

				' <SERVER-SIDE VALIDATION STUFF HERE>

				aok = True
				errorfield = ""
				returnmessage = ""
				returnclass = "standardmessage"

				'If Request.Form("txtSupervisor") = "" Then
				'	aok = False
				'	errorfield = "top.frames['fraTopic'].document.forms['mcform'].txtSupervisor"
				'	errortabinfo = "wo_details"
				'	returnmessage = "Supervisor can NOT be left blank (Server Side Validation)"
				'	returnclass = "errormessage"
				'	Call DoResponse(lastaction,"",True,True,False,keyvalue)
				'End If

				actionwhere = "WHERE (POPK IN (" & keyvalue & "))"

				If DemoMode() Then
					sql = "Select 'DemoTest' FROM PurchaseOrder WITH (NOLOCK) " & _
					actionwhere & " AND PurchaseOrder.DemoLaborPK = " & GetSession("UserPK")
					Call DemoNoActionMsg(db,sql)
				End If

				sql = _
				"UPDATE PurchaseOrder WITH (ROWLOCK) " & _
				"	SET IsOpen = 1, Status = 'ONHOLD', StatusDesc = 'On-Hold', StatusDate = getDate(), RowVersionUserPK = " & rowuser & ", RowVersionInitials = '" & rowinitials & "', RowVersionAction = 'STATUS', RowVersionDate = getdate() " & _
				actionwhere

				'Response.Write sql
				'Response.End

				If db.OpenClientConnection and db.OpenTransaction Then
					If Not db.RunSQL(sql,"") Then
						Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
					End If
				'	If Not db.RunSP("MC_OrderUnorderParts",Array(Array("@POPK", adInteger, adParamInput, 4, keyvalue)),"") Then
				'		Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				'	End If
					db.CloseTransaction
				Else
					Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				End If

				returnmessage = "The Purchase Order is now On-Hold"
				If recinview = "YES" Then
					'newjs = newjs + "		top.addpostatus('ONHOLD');" + nl
    				newjs = newjs + "	    if (top.addontheflymode == true) {top.addonthefly_recupdated = true;}" + nl	
					newjs = newjs + "		top.refreshcurrentrecord();" + nl
					newjs = newjs + "		top.refreshcurrentexplorer(true);"+nl
				End If

		Case "APPROVE"
				' <SERVER-SIDE VALIDATION STUFF HERE>
                dim approveLevel, rPOApprovalType, rcRow, poIsApproved
                dim cRow, MaxLevels, tmpApprovalLevel,cCount,foundLevel,nextLevel
                rPOApprovalType = 1
                Set rcRow = db.RunSQLReturnRS("SELECT TOP 1 POApprovalType, AuthLevelsRequired, IsApproved FROM RepairCenter RC WITH ( NOLOCK ) JOIN PurchaseOrder PO ON ( RC.RepairCenterPK = PO.RepairCenterPK ) WHERE PO.POPK = " & keyvalue, "")
                If Not rcRow.EOF Then
                    rPOApprovalType = rcRow("POApprovalType")
					poIsApproved = rcRow("IsApproved")
                    MaxLevels = rcRow("AuthLevelsRequired")
                Else
                    rPOApprovalType = 1
                    poIsApproved = False
                    MaxLevels = 0
                End If
                Set rcRow = Nothing

                approveLevel=0
                Dim accessLevel1,accessLevel2,accessLevel3,accessLevel4,accessLevel5
                accessLevel1 = GetAccessRight(db,"POAuthLevel1",0)
                accessLevel2 = GetAccessRight(db,"POAuthLevel2",0)
                accessLevel3 = GetAccessRight(db,"POAuthLevel3",0)
                accessLevel4 = GetAccessRight(db,"POAuthLevel4",0)
                accessLevel5 = GetAccessRight(db,"POAuthLevel5",0)

                if accessLevel5 and approveLevel=0 Then approveLevel = 5
                if accessLevel4 and approveLevel=0 Then approveLevel = 4
                if accessLevel3 and approveLevel=0 Then approveLevel = 3
                if accessLevel2 and approveLevel=0 Then approveLevel = 2
                if accessLevel1 and approveLevel=0 Then approveLevel = 1

                Dim aCount, checkLevelAccess
                aCount = Clng(MaxLevels)
                Do While aCount >= 1
                    Set cRow = db.RunSQLReturnRS("SELECT TOP 1 ApprovalLevel = IsNull(ApprovalLevel,0) - 1 FROM MCApprovals WITH ( NOLOCK ) WHERE POPK = " & keyvalue & " AND ApprovalLevel = " & aCount, "")
                    If cRow.EOF Then
                        nextLevel = aCount
                        Exit Do
                    End If
                    Set cRow = Nothing
                    aCount = aCount - 1
                Loop

                Dim hasApprovalAccess
                hasApprovalAccess = True
                If CBool(poIsApproved)  Then
		                newjs = "   top.mcalert('warning','Purchase Order Approved','This purchase order has already been approved.','bg_okprint',440,220,'sounds/error.wav');" + nl
				        returnmessage = ""
		                Call DoResponse("",newjs,false,True,False,"")
                Else
                    If Clng(approveLevel) = 0 AND Clng(rPOApprovalType) <> 1 Then
		                newjs = "   top.mcalert('warning','No Approval Access','You do not have the correct access rights to approve purchase orders for this repair center.','bg_okprint',400,220,'sounds/error.wav');" + nl
				        returnmessage = ""
		                Call DoResponse("",newjs,false,True,False,"")
                        hasApprovalAccess = False
                    ElseIf CLng(approveLevel) < CLng(MaxLevels) Then
                        hasApprovalAccess = False
                        checkLevelAccess = CLng(approveLevel)
                        aCount = CLng(checkLevelAccess)
                        Do While aCount >= 1
                            Set cRow = db.RunSQLReturnRS("SELECT TOP 1 ApprovalLevel = IsNull(ApprovalLevel,0) - 1 FROM MCApprovals WITH ( NOLOCK ) WHERE POPK = " & keyvalue & " AND ApprovalLevel = " & aCount, "")
                            If cRow.EOF Then
                                approveLevel = aCount
                                nextLevel = aCount
                                MaxLevels = aCount
                                hasApprovalAccess = True
                                Exit Do
                            End If
                            Set cRow = Nothing
                            aCount = aCount - 1
                        Loop
                        If Not hasApprovalAccess AND Clng(rPOApprovalType) <> 1 Then
		                    newjs = "   top.mcalert('warning','No Authorization','You do not have the appropriate access rights to approve purchase orders at this level.','bg_okprint',400,220,'sounds/error.wav');" + nl
				            returnmessage = ""
		                    Call DoResponse("",newjs,false,True,False,"")
                        End If
                    End If
                    IF Clng(rPOApprovalType) = 1 Then hasApprovalAccess = True
                    If hasApprovalAccess Then
                        If Clng(nextLevel) <= 0 Then nextLevel = 1
                        If CLng(approveLevel) > CLng(nextLevel) then
                            approveLevel = nextLevel
                        Else
                            approveLevel = MaxLevels
                        End If

				        aok = True
				        errorfield = ""
				        returnmessage = ""
				        returnclass = "standardmessage"

				        actionwhere = "WHERE (POPK IN (" & keyvalue & "))"

				        If DemoMode() Then
					        sql = "Select 'DemoTest' FROM PurchaseOrder WITH (NOLOCK) " & _
					        actionwhere & " AND PurchaseOrder.DemoLaborPK = " & GetSession("UserPK")
					        Call DemoNoActionMsg(db,sql)
				        End If

                        ' v70 Bug fix for PO's created from Order Parts that have AuthType = 2 (Amount)
				        sql = _
				        "UPDATE PurchaseOrder WITH (ROWLOCK) " & _
				        "	SET IsApproved = 1 " & _
				        actionwhere & _
				        " AND PurchaseOrder.AuthStatus = 'NOTREQUIRED' "

				        sql = sql & _
				        "UPDATE PurchaseOrder WITH (ROWLOCK) " & _
				        "	SET AuthStatus = 'APPROVED', AuthStatusDesc = 'Approved', StatusDate = getDate(), AuthStatusUserPK = " & rowuser & ", AuthStatusUserInitials = '" & rowinitials & "', AuthStatusDate = getdate(), RowVersionUserPK = " & rowuser & ", RowVersionInitials = '" & rowinitials & "', RowVersionAction = 'AUTHSTATUS', RowVersionDate = getdate() " & _
				        actionwhere

                        sql = sql & " Exec MCApprovalsUpdate " & keyvalue & ", 0, 'PO', " & approveLevel & ", 1, " & rowuser & ", '" & rowinitials & "', " & rowuser & ", '" & rowinitials & "' "

				        If db.OpenClientConnection and db.OpenTransaction Then
					        If Not db.RunSQL(sql,"") Then
						        Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
					        End If
					        db.CloseTransaction
				        Else
					        Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				        End If

				        returnmessage = "The Purchase Order has been Approved"
				        If recinview = "YES" Then
					        'newjs = newjs + "		top.addpoauth('APPROVED');" + nl
            				newjs = newjs + "	    if (top.addontheflymode == true) {top.addonthefly_recupdated = true;}" + nl	
					        newjs = newjs + "		top.refreshcurrentrecord();" + nl
					        newjs = newjs + "		top.refreshcurrentexplorer(true);"+nl
				        End If
                    End If
                End If
		Case "UNAPPROVE"

				' <SERVER-SIDE VALIDATION STUFF HERE>

				aok = True
				errorfield = ""
				returnmessage = ""
				returnclass = "standardmessage"

				actionwhere = "WHERE (POPK IN (" & keyvalue & "))"

				If DemoMode() Then
					sql = "Select 'DemoTest' FROM PurchaseOrder WITH (NOLOCK) " & _
					actionwhere & " AND PurchaseOrder.DemoLaborPK = " & GetSession("UserPK")
					Call DemoNoActionMsg(db,sql)
				End If

				sql = _
				"UPDATE PurchaseOrder WITH (ROWLOCK) " & _
				"	SET IsOpen = 1, IsPartsOrdered = 0, Status = 'REQUESTED', StatusDesc = 'Requested', AuthStatus = CASE WHEN authlevelsrequired > 0 THEN 'REQUIRED'+RTRIM(CONVERT(char(12),authlevelsrequired)) ELSE 'NOTREQUIRED' END, AuthStatusDesc = CASE WHEN authlevelsrequired > 0 THEN 'Required - L'+RTRIM(CONVERT(char(12),authlevelsrequired)) ELSE '(Not Required)' END, StatusDate = getDate(), AuthStatusUserPK = " & rowuser & ", AuthStatusUserInitials = '" & rowinitials & "', AuthStatusDate = getdate(), RowVersionUserPK = " & rowuser & ", RowVersionInitials = '" & rowinitials & "', RowVersionAction = 'AUTHSTATUS', RowVersionDate = getdate() " & _
				actionwhere

                sql = sql & " DELETE FROM MCApprovals WITH ( ROWLOCK ) WHERE POPK = " & keyvalue & " "
				'Response.Write sql
				'Response.End

				If db.OpenClientConnection and db.OpenTransaction Then
					If Not db.RunSQL(sql,"") Then
						Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
					End If
				'	If Not db.RunSP("MC_OrderUnorderParts",Array(Array("@POPK", adInteger, adParamInput, 4, keyvalue)),"") Then
				'		Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				'	End If
					db.CloseTransaction
				Else
					Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				End If

				returnmessage = "The Purchase Order has been Unapproved"
				If recinview = "YES" Then
					'newjs = newjs + "		top.addpoauth('APPROVED');" + nl
    				newjs = newjs + "	    if (top.addontheflymode == true) {top.addonthefly_recupdated = true;}" + nl	
					newjs = newjs + "		top.refreshcurrentrecord();" + nl
					newjs = newjs + "		top.refreshcurrentexplorer(true);"+nl
				End If

		Case "DENY"

				' <SERVER-SIDE VALIDATION STUFF HERE>

				aok = True
				errorfield = ""
				returnmessage = ""
				returnclass = "standardmessage"

				actionwhere = "WHERE (PurchaseOrder.POPK IN (" & keyvalue & "))"

				If DemoMode() Then
					sql = "Select 'DemoTest' FROM PurchaseOrder WITH (NOLOCK) " & _
					actionwhere & " AND PurchaseOrder.DemoLaborPK = " & GetSession("UserPK")
					Call DemoNoActionMsg(db,sql)
				End If

				If db.OpenClientConnection and db.OpenTransaction Then

		            ' If we are denying - then set the OrderUnitQty = OrderUnitQtyReceived
		            ' so that the OnOrder and OnOrderOpenPO fields get updated correctly
		            ' in the PartLocation Table

				    sql = _
				    "UPDATE PurchaseOrderDetail WITH (ROWLOCK) " & _
			        "	SET OrderUnitQty = OrderUnitQtyReceived, RowVersionUserPK = " & rowuser & ", RowVersionInitials = '" & rowinitials & "', RowVersionDate = getdate() " & _
				    " WHERE POPK IN (SELECT POPK FROM PurchaseOrder " & actionwhere & ") "

				    sql = sql & _
				    "UPDATE PurchaseOrder WITH (ROWLOCK) " & _
				    "	SET IsOpen = 0, Status = 'DENIED', StatusDesc = 'Denied', StatusDate = getDate(), RowVersionUserPK = " & rowuser & ", RowVersionInitials = '" & rowinitials & "', RowVersionAction = 'STATUS', RowVersionDate = getdate() " & _
				    actionwhere

				    'Response.Write sql
				    'Response.End

					If Not db.RunSQL(sql,"") Then
						Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
					End If
				'	If Not db.RunSP("MC_OrderUnorderParts",Array(Array("@POPK", adInteger, adParamInput, 4, keyvalue)),"") Then
				'		Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				'	End If
					db.CloseTransaction
				Else
					Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				End If

				returnmessage = "The Purchase Order has been Denied"
				If recinview = "YES" Then
					'newjs = newjs + "		top.addpostatus('DENIED');" + nl
    				newjs = newjs + "	    if (top.addontheflymode == true) {top.addonthefly_recupdated = true;}" + nl	
					newjs = newjs + "		top.refreshcurrentrecord();" + nl
					newjs = newjs + "		top.refreshcurrentexplorer(true);"+nl
				End If

		Case "REJECT"

				' <SERVER-SIDE VALIDATION STUFF HERE>

				aok = True
				errorfield = ""
				returnmessage = ""
				returnclass = "standardmessage"

				actionwhere = "WHERE (POPK IN (" & keyvalue & "))"

				If DemoMode() Then
					sql = "Select 'DemoTest' FROM PurchaseOrder WITH (NOLOCK) " & _
					actionwhere & " AND PurchaseOrder.DemoLaborPK = " & GetSession("UserPK")
					Call DemoNoActionMsg(db,sql)
				End If

				sql = _
				"UPDATE PurchaseOrder WITH (ROWLOCK) " & _
				"	SET IsOpen = 0, Status = 'DENIED', StatusDesc = 'Denied', StatusDate = getDate(), AuthStatus = 'REJECTED', AuthStatusDesc = 'Rejected', AuthStatusUserPK = " & rowuser & ", AuthStatusUserInitials = '" & rowinitials & "', AuthStatusDate = getdate(), RowVersionUserPK = " & rowuser & ", RowVersionInitials = '" & rowinitials & "', RowVersionAction = 'AUTHSTATUS', RowVersionDate = getdate() " & _
				actionwhere

				'Response.Write sql
				'Response.End

				If db.OpenClientConnection and db.OpenTransaction Then
					If Not db.RunSQL(sql,"") Then
						Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
					End If
				'	If Not db.RunSP("MC_OrderUnorderParts",Array(Array("@POPK", adInteger, adParamInput, 4, keyvalue)),"") Then
				'		Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				'	End If
					db.CloseTransaction
				Else
					Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				End If

				returnmessage = "The Authorization for this Purchase Order has been Rejected"
				If recinview = "YES" Then
					'newjs = newjs + "		top.addpostatus('DENIED');" + nl
					'newjs = newjs + "		top.addpoauth('REJECTED');" + nl
    				newjs = newjs + "	    if (top.addontheflymode == true) {top.addonthefly_recupdated = true;}" + nl	
					newjs = newjs + "		top.refreshcurrentrecord();" + nl
					newjs = newjs + "		top.refreshcurrentexplorer(true);"+nl
				End If

		Case "INITIATE"

				' <SERVER-SIDE VALIDATION STUFF HERE>

				aok = True
				errorfield = ""
				returnmessage = ""
				returnclass = "standardmessage"

				actionwhere = "WHERE (POPK IN (" & keyvalue & "))"

				If DemoMode() Then
					sql = "Select 'DemoTest' FROM PurchaseOrder WITH (NOLOCK) " & _
					actionwhere & " AND PurchaseOrder.DemoLaborPK = " & GetSession("UserPK")
					Call DemoNoActionMsg(db,sql)
				End If

                ' v70 Bug fix for PO's created from Order Parts that have AuthType = 2 (Amount)
				sql = _
				"UPDATE PurchaseOrder WITH (ROWLOCK) " & _
				"	SET IsApproved = 1 " & _
				actionwhere & _
				" AND PurchaseOrder.AuthStatus = 'NOTREQUIRED' "
                
				' We need to Approve as well if the Work Order is NOT approved and Approval is Required
				sql = sql & _
				"UPDATE PurchaseOrder WITH (ROWLOCK) " & _
				"	SET AuthStatus = 'APPROVED', AuthStatusDesc = 'Approved', StatusDate = getDate(), AuthStatusUserPK = " & rowuser & ", AuthStatusUserInitials = '" & rowinitials & "', AuthStatusDate = getdate(), RowVersionUserPK = " & rowuser & ", RowVersionInitials = '" & rowinitials & "', RowVersionAction = 'AUTHSTATUS', RowVersionDate = getdate() " & _
				actionwhere & _
				" AND PurchaseOrder.IsApproved = 0 "

				sql = sql & _
				"UPDATE PurchaseOrder WITH (ROWLOCK) " & _
				"	SET IsOpen = 1, Status = 'ISSUED', StatusDesc = 'Issued', StatusDate = getDate(), RowVersionUserPK = " & rowuser & ", RowVersionInitials = '" & rowinitials & "', RowVersionAction = 'STATUS', RowVersionDate = getdate() " & _
				actionwhere & _
				" AND PurchaseOrder.IsApproved = 1 "

				'Response.Write sql
				'Response.End

				If db.OpenClientConnection and db.OpenTransaction Then
					If Not db.RunSQL(sql,"") Then
						Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
					End If
				'	If Not db.RunSP("MC_OrderUnorderParts",Array(Array("@POPK", adInteger, adParamInput, 4, keyvalue)),"") Then
				'		Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				'	End If
					db.CloseTransaction
				Else
					Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				End If

				If recinview = "YES" Then
					'newjs = newjs + "		top.addpostatus('ISSUED');" + nl
    				newjs = newjs + "	    if (top.addontheflymode == true) {top.addonthefly_recupdated = true;}" + nl	
					newjs = newjs + "		top.refreshcurrentrecord();" + nl
					newjs = newjs + "		top.refreshcurrentexplorer(true);"+nl
				End If
				returnmessage = "The Purchase Order has been Issued"

		Case "CANCEL"

				' <SERVER-SIDE VALIDATION STUFF HERE>

				aok = True
				errorfield = ""
				returnmessage = ""
				returnclass = "standardmessage"

				actionwhere = "WHERE (PurchaseOrder.POPK IN (" & keyvalue & "))"

				If DemoMode() Then
					sql = "Select 'DemoTest' FROM PurchaseOrder WITH (NOLOCK) " & _
					actionwhere & " AND PurchaseOrder.DemoLaborPK = " & GetSession("UserPK")
					Call DemoNoActionMsg(db,sql)
				End If

				If db.OpenClientConnection and db.OpenTransaction Then

	                ' If we are canceling - then set the OrderUnitQty = OrderUnitQtyReceived
	                ' so that the OnOrder and OnOrderOpenPO fields get updated correctly
	                ' in the PartLocation Table

			        sql = _
			        "UPDATE PurchaseOrderDetail WITH (ROWLOCK) " & _
			        "	SET OrderUnitQty = OrderUnitQtyReceived, RowVersionUserPK = " & rowuser & ", RowVersionInitials = '" & rowinitials & "', RowVersionDate = getdate() " & _
			        " WHERE POPK IN (SELECT POPK FROM PurchaseOrder " & actionwhere & ") "

			        sql = sql & _
				    "UPDATE PurchaseOrder WITH (ROWLOCK) " & _
				    "	SET IsOpen = 0, Status = 'CANCELED', StatusDesc = 'Canceled', StatusDate = getDate(), RowVersionUserPK = " & rowuser & ", RowVersionInitials = '" & rowinitials & "', RowVersionAction = 'STATUS', RowVersionDate = getdate() " & _
				    actionwhere

				    'Response.Write sql
				    'Response.End

					If Not db.RunSQL(sql,"") Then
						Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
					End If
				'	If Not db.RunSP("MC_OrderUnorderParts",Array(Array("@POPK", adInteger, adParamInput, 4, keyvalue)),"") Then
				'		Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				'	End If
					db.CloseTransaction
				Else
					Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				End If

				If recinview = "YES" Then
					'newjs = newjs + "		top.addpostatus('CANCELED');" + nl
    				newjs = newjs + "	    if (top.addontheflymode == true) {top.addonthefly_recupdated = true;}" + nl	
					newjs = newjs + "		top.refreshcurrentrecord();" + nl
					newjs = newjs + "		top.refreshcurrentexplorer(true);"+nl
				End If
				returnmessage = "The Purchase Order has been Canceled"


		Case "CLOSE"

				' <SERVER-SIDE VALIDATION STUFF HERE>

				aok = True
				errorfield = ""
				returnmessage = ""
				returnclass = "standardmessage"

				actionwhere = "WHERE (PurchaseOrder.POPK IN (" & keyvalue & "))"

				If DemoMode() Then
					sql = "Select 'DemoTest' FROM PurchaseOrder WITH (NOLOCK) " & _
					actionwhere & " AND PurchaseOrder.DemoLaborPK = " & GetSession("UserPK")
					Call DemoNoActionMsg(db,sql)
				End If

                ' v70 Bug fix for PO's created from Order Parts that have AuthType = 2 (Amount)
				sql = _
				"UPDATE PurchaseOrder WITH (ROWLOCK) " & _
				"	SET IsApproved = 1 " & _
				actionwhere & _
				" AND PurchaseOrder.AuthStatus = 'NOTREQUIRED' "

				' We need to Approve as well if the Purchase Order is NOT approved and Approval is Required
				sql = sql & _
				"UPDATE PurchaseOrder" & _
				"	SET AuthStatus = 'APPROVED', AuthStatusDesc = 'Approved', StatusDate = getDate(), AuthStatusUserPK = " & rowuser & ", AuthStatusUserInitials = '" & rowinitials & "', AuthStatusDate = getdate(), RowVersionUserPK = " & rowuser & ", RowVersionInitials = '" & rowinitials & "', RowVersionAction = 'AUTHSTATUS', RowVersionDate = getdate() " & _
				actionwhere & _
				" AND PurchaseOrder.IsApproved = 0 "

				' We need to Issue as well if the Purchase Order is in the Requested state
				sql = sql & _
				"UPDATE PurchaseOrder WITH (ROWLOCK) " & _
				"	SET IsOpen = 1, Status = 'ISSUED', StatusDesc = 'Issued', StatusDate = getDate(), RowVersionUserPK = " & rowuser & ", RowVersionInitials = '" & rowinitials & "', RowVersionAction = 'STATUS', RowVersionDate = getdate() " & _
				actionwhere & _
				" AND PurchaseOrder.Status = 'REQUESTED' AND PurchaseOrder.IsApproved = 1 "

                ' All Parts MUST be received for a PO to be CLOSED
	            Dim actionwhere_lineitemcheck
                actionwhere_lineitemcheck = actionwhere & " AND (PurchaseOrder.POPK IN (SELECT PurchaseOrder.POPK FROM PurchaseOrderDetail INNER JOIN PurchaseOrder ON PurchaseOrder.POPK = PurchaseOrderDetail.POPK " & actionwhere & " AND PurchaseOrder.IsApproved = 1 AND (PurchaseOrderDetail.OrderUnitQty = PurchaseOrderDetail.OrderUnitQtyReceived))) "

				sql = sql & _
				"UPDATE PurchaseOrder WITH (ROWLOCK) " & _
				"	SET IsOpen = 0, Status = 'CLOSED', StatusDesc = 'Closed', StatusDate = getDate(), RowVersionUserPK = " & rowuser & ", RowVersionInitials = '" & rowinitials & "', RowVersionAction = 'STATUS', RowVersionDate = getdate() " & _
				actionwhere_lineitemcheck

				'Response.Write sql
				'Response.End

				If db.OpenClientConnection and db.OpenTransaction Then
				If Not db.RunSQL(sql,"") Then
						Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
					End If
				'	If Not db.RunSP("MC_OrderUnorderParts",Array(Array("@POPK", adInteger, adParamInput, 4, keyvalue)),"") Then
				'		Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				'	End If
					db.CloseTransaction
				Else
					Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				End If

				returnmessage = "The Purchase Order has been Closed"
				If recinview = "YES" Then
					'newjs = newjs + "		top.addpostatus('CLOSED');" + nl
    				newjs = newjs + "	    if (top.addontheflymode == true) {top.addonthefly_recupdated = true;}" + nl	
					newjs = newjs + "		top.refreshcurrentrecord();" + nl
					newjs = newjs + "		top.refreshcurrentexplorer(true);"+nl
				End If

		Case "DELETE"

				' <SERVER-SIDE VALIDATION STUFF HERE>

				aok = True
				errorfield = ""
				returnmessage = ""
				returnclass = "standardmessage"

				actionwhere = "WHERE (POPK IN (" & keyvalue & "))"

				If DemoMode() Then
					sql = "Select 'DemoTest' FROM PurchaseOrder WITH (NOLOCK) " & _
					actionwhere & " AND PurchaseOrder.DemoLaborPK = " & GetSession("UserPK")
					Call DemoNoActionMsg(db,sql)
				End If

				sql = _
				"UPDATE PurchaseOrder WITH (ROWLOCK) SET RowVersionUserPK = '" & GetSession("UserPK") & "' " & actionwhere & " " & _
				"DELETE PurchaseOrder WITH (ROWLOCK) " & _
				actionwhere

				If Not db.RunSQL(sql,"") Then
  					If db.isdeletecolref Then
		                newjs = newjs + "   top.removemessage(); top.showactions('" & lastaction & "');top.loadmodeless('modules/common/mc_deleterecord.asp?t=PURCHASEORDER&pk="&keyvalue&"',null,995,700,false,false);"+nl
		                Call DoScript(newjs,False)
  					Else
    					Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
    				End If
				End If

				If recinview = "YES" Then

					newjs = newjs + "	top.delete_onafterevent('PO','The Purchase Order has been Deleted');" + nl

					Call DoScript(newjs,False)

				Else
					returnmessage = "The Purchase Order has been Deleted"
				End If

		Case "CLONE"

				' <SERVER-SIDE VALIDATION STUFF HERE>

				Dim ClonePK

				aok = True
				errorfield = ""
				returnmessage = ""
				returnclass = "standardmessage"

				If Not db.RunSP("MC_ClonePO",Array(Array("@PK", adInteger, adParamInput, 4, keyvalue),Array("@ClonePK", adInteger, adParamOutPut, 4, "")),OutArray) Then
					Call dok_check(db,"Purchase Order Message","There was a problem processing your request. The details of the problem are described below. You can try your request again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
				End If
				ClonePK = OutArray(1)

				If recinview = "YES" Then

					newjs = newjs + "		top.pendingmsg = 'The Purchase Order has been Cloned';" + nl
					newjs = newjs + "		if (top.addontheflymode == true) {top.addonthefly_recupdated = true;}" + nl
					newjs = newjs + "		top.recordkey = 'kv=" & CStr(ClonePK) & "';" + nl
					newjs = newjs + "		top.refreshcurrentexplorer(true);" + nl
					newjs = newjs + "		top.navadd('PO','kv=" & CStr(ClonePK) & "');" + nl

					Call DoScript(newjs,False)

				Else
					returnmessage = "The Purchase Order has been Cloned"
				End If

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
			Response.Write "top.changepage(myframe.document.getElementById('pagetabs_21'));" + nl
			'Response.Write "myform.txtPOName.focus();" + nl
			Response.Write "try {myform.txtVendor.focus();} catch(e) {};" + nl
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
		Response.Write("	myframe.document.getElementById('fraRMA').style.height = top.fraTopic.innerHeight - 150+'px';") + nl
		Response.Write("	myframe.document.getElementById('fraReports').style.height = top.fraTopic.innerHeight - 160 + 'px';") + nl
		Response.Write("	myframe.document.getElementById('fraReceipts').style.height = top.fraTopic.innerHeight - 140 +'px';") + nl
		Response.Write("	myframe.document.getElementById('fraInvoices').style.height = top.fraTopic.innerHeight - 140 + 'px';") + nl
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
			'Moved to AddPOStatus() in mcmodules.js
			'Response.Write("top.fraTabbar.document.getElementById('header_keytext').innerText='New Purchase Order';")+nl
			'Response.Write("top.fraTabbar.document.getElementById('header_keytext2').innerText='';")+nl							
            'Response.Write("top.walk(top.fraTabbar.document.getElementById('header_keydiv'), null, top, false, true);")+nl
			Response.Write("top.clearallfields();")+nl
			Response.Write("top.updatestatusbar('',null,'');")+nl
			ClearTables
			' UI Settings
			'------------------------------------------------------------------------------------

			Response.Write("myform.txtIsPartsOrdered.value = 'N';")

			Response.Write("	myframe.approvednoedit = false; ")+nl

			' Defaults
			'------------------------------------------------------------------------------------
			Response.Write("top.sethidden('txtPO','(New)');")+nl

			Set db = New ADOHelper
			Dim prefvalue, prefdesc, prefpk, prefsql, prefrs,iRCPK
            iRCPK = 0
			Response.Write("	// Write Purchase Order Defaults")+nl
			Response.Write("	// -------------------------------------------------------------------------")+nl

			'Set rs = db.RunSPReturnMultiRS("MC_GetDefaultsPO",Array(Array("@LaborPK", adInteger, adParamInput, 4, GetSession("UserPK")),Array("@RepairCenterPK", adInteger, adParamInput, 4, GetSession("RCPK"))),"")
			'Call dok_check(db,"Purchase Order Message","There was a problem retrieving the Purchase Order defaults. The details of the problem are described below. You can try to create the Work Order again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")

			'If not rs.eof Then
			'	Response.Write("	top.setpk('txtBuyerPK','" & NullCheck(RS("LaborPK")) & "') ;")+nl
			'	Response.Write("	myform.txtBuyer.value = '" & JSEncode(RS("LaborID")) & "';")+nl
			'	Response.Write("	top.setdesc(myframe.txtBuyerDesc,'" & JSEncode(RS("LaborName")) & "') ;")+nl
			'End If
			'Set rs = rs.nextrecordset
			'If not rs.eof Then
			'	Response.Write("	top.setpk('txtRepairCenterPK','" & NullCheck(RS("RepairCenterPK")) & "') ;")+nl
			'	Response.Write("	myform.txtRepairCenter.value = '" & JSEncode(RS("RepairCenterID")) & "';")+nl
			'	Response.Write("	top.setdesc(myframe.txtRepairCenterDesc,'" & JSEncode(RS("RepairCenterName")) & "') ;")+nl
			'End If

			'Response.Write("	top.setpk('txtRepairCenterPK','" & NullCheck(GetSession("RCPK")) & "') ;")+nl
			'Response.Write("	myform.txtRepairCenter.value = '" & JSEncode(GetSession("RCID")) & "';")+nl
			'Response.Write("	top.setdesc(myframe.txtRepairCenterDesc,'" & JSEncode(GetSession("RCNM")) & "') ;")+nl

			If GetPreference(db,True,"PO_DEFAULTBUYERCOMPANY",prefvalue, prefdesc, prefpk) Then
				Response.Write("	top.setpk('txtBuyerCompanyPK','" & NullCheck(prefpk) & "');")+nl
				Response.Write("	myform.txtBuyerCompany.value = '" & JSEncode(prefvalue) & "';")+nl
				Response.Write("	top.setdesc(myframe.txtBuyerCompanyDesc,'" & JSEncode(prefdesc) & "');")+nl
			End If

			If GetPreference(db,True,"PO_DEFAULTREPAIRCENTER",prefvalue, prefdesc, prefpk) Then
                iRCPK = prefpk
				Response.Write("	top.setpk('txtRepairCenterPK','" & NullCheck(prefpk) & "');")+nl
				Response.Write("	myform.txtRepairCenter.value = '" & JSEncode(prefvalue) & "';")+nl
				Response.Write("	top.setdesc(myframe.txtRepairCenterDesc,'" & JSEncode(prefdesc) & "');")+nl
			End If

			If GetPreference(db,True,"PO_DEFAULTSHIPTO",prefvalue, prefdesc, prefpk) Then
				Response.Write("	top.setpk('txtShipToPK','" & NullCheck(prefpk) & "');")+nl
				Response.Write("	myform.txtShipTo.value = '" & JSEncode(prefvalue) & "';")+nl
				Response.Write("	top.setdesc(myframe.txtShipToDesc,'" &  JSEncode(prefdesc) & "');")+nl

				prefsql = _
				"Select a.* FROM Company a WITH (NOLOCK) WHERE a.CompanyPK = " & NullCheck(prefpk) & " "

				Set prefrs = db.RunSQLReturnRS(prefsql,"")
				If db.dok Then
					If Not prefrs.eof Then
						Response.Write("	myform.txtShipToAttention.value = '" & JSEncode(prefrs("Attention")) & "';")+nl
						Response.Write("	myform.txtShipToAddress1.value = '" & JSEncode(prefrs("Address")) & "';")+nl
						Response.Write("	myform.txtShipToAddress2.value = '" & JSEncode(prefrs("Address2")) & "';")+nl
						If NullCheck(prefrs("City")) = "" Then
							Response.Write("	myform.txtShipToAddress3.value = '" & JSEncode(prefrs("State") & " " & prefrs("Zip")) & "';")+nl
						Else
							Response.Write("	myform.txtShipToAddress3.value = '" & JSEncode(prefrs("City") & ", " & prefrs("State") & " " & prefrs("Zip")) & "';")+nl
						End If
					End If
				End If
			End If

			If GetPreference(db,True,"PO_DEFAULTBILLTO",prefvalue, prefdesc, prefpk) Then
				Response.Write("	top.setpk('txtBillToPK','" & NullCheck(prefpk) & "');")+nl
				Response.Write("	myform.txtBillTo.value = '" & JSEncode(prefvalue) & "';")+nl
				Response.Write("	top.setdesc(myframe.txtBillToDesc,'" & JSEncode(prefdesc) & "');")+nl

				prefsql = _
				"Select a.* FROM Company a WITH (NOLOCK) WHERE a.CompanyPK = " & NullCheck(prefpk) & " "

				Set prefrs = db.RunSQLReturnRS(prefsql,"")
				If db.dok Then
					If Not prefrs.eof Then
						Response.Write("	myform.txtBillToAttention.value = '" & JSEncode(prefrs("Attention")) & "';")+nl
						Response.Write("	myform.txtBillToAddress1.value = '" & JSEncode(prefrs("Address")) & "';")+nl
						Response.Write("	myform.txtBillToAddress2.value = '" & JSEncode(prefrs("Address2")) & "';")+nl
						If NullCheck(prefrs("City")) = "" Then
							Response.Write("	myform.txtBillToAddress3.value = '" & JSEncode(prefrs("State") & " " & prefrs("Zip")) & "';")+nl
						Else
							Response.Write("	myform.txtBillToAddress3.value = '" & JSEncode(prefrs("City") & ", " & prefrs("State") & " " & prefrs("Zip")) & "';")+nl
						End If
					End If
				End If
			End If

			Response.Write("	top.setpk('txtBuyerPK','" & NullCheck(GetSession("USERPK")) & "') ;")+nl
			Response.Write("	myform.txtBuyer.value = '" & JSEncode(GetSession("USERID")) & "';")+nl
			Response.Write("	top.setdesc(myframe.txtBuyerDesc,'" & JSEncode(GetSession("USERNAME")) & "') ;")+nl

			Response.Write("	top.setpk('txtRequesterPK','" & NullCheck(GetSession("USERPK")) & "') ;")+nl
			Response.Write("	myform.txtRequester.value = '" & JSEncode(GetSession("USERID")) & "';")+nl
			Response.Write("	top.setdesc(myframe.txtRequesterDesc,'" & JSEncode(GetSession("USERNAME")) & "') ;")+nl

            Dim rcRow, rcOverride
            If iRCPK = "" Then iRCPK = 0
            Set rcRow = db.RunSQLReturnRS("SELECT TOP 1 POApprovalType FROM RepairCenter WITH ( NOLOCK ) WHERE RepairCenterPK=" & iRCPK, "")
            'Call SetSession("POAuth","Y")

            If Not rcRow.EOF Then
                rcOverride = rcRow("POApprovalType")
            Else
                rcOverride = 1
            End If
            Set rcRow = Nothing

			If GetSession("POAuthReq") = "0" Then
				Response.Write("myframe.authorizationheader.innerText = 'Approval';")
				Response.Write("myform.txtIsApproved.value = 'Y';")
			Else
				Response.Write("myframe.authorizationheader.innerText = 'Approval: Pending';")
				Response.Write("myform.txtIsApproved.value = 'N';")
			End If

			Response.Write("	top.addpostatus('REQUESTED');")+nl
			'Response.Write("	top.addpoauth('NOTREQUIRED');")+nl
            If Clng(rcOverride) = 0 Then
				    Response.Write("top.addpoauth('NOTREQUIRED');")+nl
            ElseIf Clng(rcOverride) = 2 Then
			    Response.Write("top.addpoauth('REQUIRED1');")+nl
            Else
			    If GetSession("POAuthReq") = "0" Then
				    Response.Write("top.addpoauth('NOTREQUIRED');")+nl
			    Else
				    Response.Write("top.addpoauth('REQUIRED" & GetSession("POAuthReq") & "');")+nl
			    End If
            End If

			Response.Write("	myform.txtPODate.value = '" + CStr(DateNullCheck(Date())) + "';")+nl
			Response.Write("	myform.txtFollowupDate.value = '" + CStr(DateNullCheck(Date()+7)) + "';")+nl

			Response.Write("	myform.txtSubtotal.value = '0';")+nl
			Response.Write("	myform.txtSubTotalPreDiscount.value = '0';")+nl
			Response.Write("	myform.txtFreightCharge.value = '0';")+nl
			Response.Write("	myform.txtTaxAmount.value = '0';")+nl
			Response.Write("	myform.txtDiscountPercentage.value = '0';")+nl
			Response.Write("	myform.txtDiscount.value = '0';")+nl
			Response.Write("	myform.txtTotal.value = '0';")+nl

			Call OutputUDFLabels(Null,"PurchaseOrder")

			If GetPreference(db,False,"PO_DefaultPriority",prefvalue, prefdesc, prefpk) Then
				Response.Write("	myform.txtPriority.value = '" & JSEncode(prefvalue) & "';")+nl
				Response.Write("	top.setdesc(myframe.txtPriorityDesc,'" & JSEncode(prefdesc) & "') ;")+nl
			End If

			Response.Write("	top.updatestatusbarpo();")+nl
            Response.Write("	try{myframe.ForceOverride(0);} catch(e){}")+nl
		Case "EDIT"

			' Standard Settings
			'------------------------------------------------------------------------------------
			Response.Write("top.recordkey = 'kv=" + CStr(keyvalue) + "'")+nl
			Response.Write("top.moduleaction = 'asave';")+nl
			Response.Write("top.showactions('aporecbar');")+nl
	End Select
    %>

    <%
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
			'Response.Write("myform.txtPmNo.disabled = true;")+nl
			'Response.Write("myform.txtPmNo.className = 'disabled';")+nl

	End Select
 	Response.Write("")+nl

End Function

Function ClearTables()
	Response.Write("")+nl

	Response.Write("myframe.po_details.setAttribute('dataloaded','Y');") + nl
	Response.Write(nl)
	Response.Write("function loadtab61()")+nl
	Response.Write("{")+nl
	Response.Write(nl)

	Response.Write("top.cleargenerictable(myframe.opostatus);")+nl
	Response.Write("top.cleargenerictable(myframe.opoauth);")+nl

	' CUSTOMIZED
	'--------------------------------------------------------------------------------------------------------------
	'Response.Write("	top.displayphoto(mydoc.images.purchaseorderphoto,'');")+nl
	'--------------------------------------------------------------------------------------------------------------

	Response.Write(nl)
	Response.Write("}")+nl
	Response.Write("loadtab61();")+nl

	Response.Write(nl)
	Response.Write("myframe.po_lineitems.setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab62()")+nl
	Response.Write("{")+nl

	Response.Write(nl)
	Response.Write("// Clear Line Item Records")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("top.cleartable(myframe.oli1);")+nl
	Response.Write("parent.PORecalc('');")+nl
	Response.Write(nl)
	Response.Write("}")+nl

	Response.Write(nl)
	Response.Write("myframe.po_receipts.setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab63()")+nl
	Response.Write("{")+nl
	Response.Write(nl)

	Response.Write("// Clear Items")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("   myframe.document.getElementById('fraReceipts').style.display = 'none';") + nl
	Response.Write("   myframe.document.getElementById('fraInvoices').style.display = 'none';") + nl

	Response.Write(nl)
	Response.Write("}")+nl

	Response.Write(nl)
	Response.Write("myframe.po_rma.setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab64()")+nl
	Response.Write("{")+nl
	Response.Write(nl)

	Response.Write("// Clear Items")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("   mydoc.fraRMA.style.display = 'none';") + nl

	Response.Write(nl)
	Response.Write("}")+nl

	Response.Write(nl)
	Response.Write("myframe.po_attach.setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab65()")+nl
	Response.Write("{")+nl
	Response.Write(nl)

	Response.Write("// Clear Documents")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("top.cleartable(myframe.oat1);")+nl

	Response.Write(nl)
	Response.Write("// Clear Notes Rows")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("top.cleartable(myframe.ono1);")+nl

	Response.Write(nl)
	Response.Write("// Clear Misc Attachments")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
    'Response.Write("	myframe.document.getElementById('fileFrame').style.display = 'none';") + nl

	Response.Write(nl)
	Response.Write("// Clear Rules Rows")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("top.cleartable(myframe.document.getElementById('orl1'));")+nl

	Response.Write(nl)
	Response.Write("}")+nl

	Response.Write(nl)
	Response.Write("myframe.po_report.setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab66()")+nl
	Response.Write("{")+nl
	Response.Write(nl)

	Response.Write("// Clear Items")+nl
	Response.Write("// -------------------------------------------------------------------------")+nl
	Response.Write("   mydoc.getElementById('fraReports').style.display = 'none';") + nl

	Response.Write(nl)
	Response.Write("}")+nl

End Function

Function writerecorddata(db)

	Dim killdbonend,rs,outarray,dok,derror,newjs,playsound

	writerecorddata = True

	If IsNull(db) or (Not IsObject(db)) Then
		Set db = New ADOHelper
		killdbonend = True
	Else
		killdbonend = False
	End If

    Dim prefvalue, prefdesc, prefpk, IN_USEINTERNALID

    If GetDefaultPreference(db,False,"IN_USEINTERNALID",prefvalue, prefdesc, prefpk) Then
        If UCase(prefvalue) = "YES" Then
	        IN_USEINTERNALID = True
        Else
	        IN_USEINTERNALID = False
        End If
    End If

	Set rs = db.RunSPReturnMultiRS("MC_GetPurchaseOrder_Sandvik",Array(Array("@POPK", adInteger, adParamInput, 4, keyvalue),Array("@UserPK", adInteger, adParamInput, 4, GetSession("UserPK"))),outarray)
	Call dok_check(db,"Purchase Order Message","There was a problem retrieving the selected Purchase Order. The details of the problem are described below. You can try to retrieve the Purchase Order again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")

	If rs.eof Then
		db.dok = False
		db.derror = "The selected Purchase Order could not be found. Please refresh your Purchase Order list (to ensure another user has not deleted this Purchase Order) and try again."
		' End Process here (and it will actually end the process again...but no big deal)
		' We need to end it here because the navadd will fail if we don't because we are
		' actually inside a process at this point in the game.
		Response.Write("top.endprocess();") + nl
		Response.Write("top.navadd('PO','kv=');") + nl
		Response.Write("top.mcalert('warning','Purchase Order Not Found','" & db.derror & "','bg_okprint',700,235,'sounds/error.wav');") + nl
		Call CloseObj(rs)
		If killdbonend Then
			db.CloseClientConnection
			Set db = Nothing
		End If
		writerecorddata = False
		Exit Function
	End If

	' Output All Fields Below

	Response.Write("")+nl

	Response.Write("	// Clear all hidden fields")+nl
	Response.Write("	// -------------------------------------------------------------------------")+nl
	Response.Write("	top.clearallhiddenfields();")+nl

	Response.Write("	// Set the Primary Key and Row Version")+nl
	Response.Write("	// -------------------------------------------------------------------------")+nl
	Response.Write("	top.sethidden('kv','" & JSEncode(RS("POPK")) & "');")+nl
	Response.Write("	top.sethidden('txtRowVersionDate','" & NullCheck(RS("RowVersionDate")) & "');")+nl
	Response.Write("")+nl

	'Moved to AddPOStatus() in mcmodules.js

	'Response.Write("	// Update the Header Bar")+nl
	'Response.Write("	// -------------------------------------------------------------------------")+nl
	'Response.Write("	top.fraTabbar.document.getElementById('header_keytext').innerHTML='Purchase Order:'")+nl
	'Response.Write("	top.fraTabbar.document.getElementById('header_keytext2').innerHTML=' ")
	'Response.Write(JSEncode(RS("POName")))
	'Response.Write("'")+nl
	'Response.Write("&nbsp;&nbsp;(")
	'Response.Write(JSEncode(RS("CategoryName")))
	'Response.Write(")'")+nl
    'Response.Write("top.walk(top.fraTabbar.document.getElementById('header_keydiv'), null, top, false, true);")+nl
	'Response.Write("")+nl

	Response.Write("	// Update the Status Bar")+nl
	Response.Write("	// -------------------------------------------------------------------------")+nl
	Response.Write("	top.updatestatusbarpo('" & JSEncode(RS("VendorName")) & "','" & JSEncode(RS("BuyerCompanyName")) & "','" & JSEncode(RS("BuyerName")) & "');")+nl
	Response.Write("")+nl

	Response.Write("	// Set the Tabs that MUST be loaded prior to Saving")+nl
	Response.Write("	// -------------------------------------------------------------------------")+nl
	Response.Write("	top.checktabsbeforesave = 'top.checktabdataloaded(\'po_details\');';")+nl
	Response.Write("")+nl

	Response.Write("	// Set the Purchase Order ID outside the loadtab61 function so its set for the addpostatus function ")+nl
	Response.Write("	// -------------------------------------------------------------------------")+nl
	Response.Write("	top.sethidden('txtPO','" & JSEncode(RS("POID")) & "');")+nl
	Response.Write("")+nl

	Response.Write("	// Output all fields on Tabs that might not be loaded prior to Saving ")+nl
	Response.Write("	// -------------------------------------------------------------------------")+nl
	Call OutputDefaults(rs,1)

	Response.Write("	myframe.po_details.setAttribute('dataloaded','N');") + nl
	Response.Write("")+nl
	Response.Write("function loadtab61()")+nl
	Response.Write("{")+nl
	Response.Write("")+nl

		Response.Write("	// Write Purchase Order Record Field Data")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl

		Call OutputDefaults(rs,2)

	Response.Write("}")+nl
	Response.Write(nl)

	Set rs = rs.NextRecordset

	Response.Write("	// Build Status Rows")+nl
	Response.Write("	// -------------------------------------------------------------------------")+nl
	Response.Write("	top.cleargenerictable(myframe.opostatus);")+nl
	Do Until rs.eof
		Response.Write("	top.addpostatus('" & JSEncode(RS("Status")) & "','" & DateNullCheck(RS("StatusDate")) & "','" & JSEncode(RS("RowVersionInitials")))
		rs.MoveNext
		If rs.eof Then
			Response.Write("',true);")+nl
		Else
			Response.Write("',false);")+nl
		End If
	Loop
	Response.Write("")+nl

	Set rs = rs.NextRecordset

	Response.Write("	// Build Authorization Rows")+nl
	Response.Write("	// -------------------------------------------------------------------------")+nl
	Response.Write("	top.cleargenerictable(myframe.opoauth);")+nl
	Do Until rs.eof
		Response.Write("	top.addpoauth('" & JSEncode(RS("Status")) & "','" & DateNullCheck(RS("StatusDate")) & "','" & JSEncode(RS("RowVersionInitials")))
		rs.MoveNext
		If rs.eof Then
			Response.Write("',true);")+nl
		Else
			Response.Write("',false);")+nl
		End If
	Loop
	Response.Write("")+nl

	Response.Write("	myframe.po_lineitems.setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab62()")+nl
	Response.Write("{")+nl

		Set rs = rs.NextRecordset

		Response.Write(nl)
		Response.Write("	// Build Purchase Order Line Items")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	top.cleartable(myframe.oli1);")+nl
		Dim PartIDorInternalID, taxable, sodi, bindata
		Do Until rs.eof
		    If IN_USEINTERNALID Then
		        PartIDorInternalID = RS("InternalPartNumber")
		    Else
		        PartIDorInternalID = RS("PartID")
		    End If
			If RS("istax") Then
				taxable = "<img src=""../../images/taskchecked.gif"">"
			Else
				taxable = "<img src=""../../images/taskline.gif"">"
			End If
			'Response.Write("	if (myform.txtli1DueDateROW_ID) { ")+nl
			'Response.Write("	top.builddatarow(myframe.oli1body,3,null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','" & RS("PartPK") & "','IN',false,'" & JSEncode(RS("Photo")) & "',null,null,'" & NullCheck(RS("LineItemNo")) & "','" & JSEncode(PartIDorInternalID) & "','" & JSEncode(RS("PartName")) & "','" & JSEncode(RS("LocationID")) & "','" & JSEncode(RS("VendorPartNumber")) & "','" & JSEncode(RS("WOID")) & "','" & JSEncode(RS("AssetID")) & "<mcbr> / " & JSEncode(RS("AssetName")) & "','" & JSEncode(RS("AccountID")) & "<mcbr> / " & JSEncode(RS("AccountName")) & "','" & JSEncode(RS("orderunits")) & "<mcbr> / " & JSEncode(RS("orderunitsdesc")) & "','" & NullCheck(RS("ConversionToIssueUnits")) & "','" & JSEncode(RS("IssueUnits")) & "<mcbr> / " & JSEncode(RS("IssueUnitsDesc")) & "','" & FormatNumber(NullCheck(RS("OrderUnitPrice")),4,-2,0,0) & "','" & NullCheck(RS("OrderUnitQty")) & "','" & NullCheck(RS("OrderUnitQtyReceived")) & "','" & NullCheck(RS("OrderUnitQtyBackOrdered")) & "','" & NullCheck(RS("Discount")) & "','" & FormatNumber(NullCheck(RS("Subtotal")),4,-2,0,0) & "','" & taxable & "','" & NullCheck(RS("TaxRate")) & "','" & FormatNumber(NullCheck(RS("TaxAmount")),4,-2,0,0) & "','" & FormatNumber(NullCheck(RS("LineItemTotal")),4,-2,0,0) & "','" & DateNullCheck(RS("DueDate")) & "','" & JSEncode(RS("Comments")) & "');")+nl
			'Response.Write("	} else { ")+nl
			'Response.Write("	top.builddatarow(myframe.oli1body,3,null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','" & RS("PartPK") & "','IN',false,'" & JSEncode(RS("Photo")) & "',null,null,'" & NullCheck(RS("LineItemNo")) & "','" & JSEncode(PartIDorInternalID) & "','" & JSEncode(RS("PartName")) & "','" & JSEncode(RS("LocationID")) & "','" & JSEncode(RS("VendorPartNumber")) & "','" & JSEncode(RS("WOID")) & "','" & JSEncode(RS("AssetID")) & "<mcbr> / " & JSEncode(RS("AssetName")) & "','" & JSEncode(RS("AccountID")) & "<mcbr> / " & JSEncode(RS("AccountName")) & "','" & JSEncode(RS("orderunits")) & "<mcbr> / " & JSEncode(RS("orderunitsdesc")) & "','" & NullCheck(RS("ConversionToIssueUnits")) & "','" & JSEncode(RS("IssueUnits")) & "<mcbr> / " & JSEncode(RS("IssueUnitsDesc")) & "','" & FormatNumber(NullCheck(RS("OrderUnitPrice")),4,-2,0,0) & "','" & NullCheck(RS("OrderUnitQty")) & "','" & NullCheck(RS("OrderUnitQtyReceived")) & "','" & NullCheck(RS("OrderUnitQtyBackOrdered")) & "','" & NullCheck(RS("Discount")) & "','" & FormatNumber(NullCheck(RS("Subtotal")),4,-2,0,0) & "','" & taxable & "','" & NullCheck(RS("TaxRate")) & "','" & FormatNumber(NullCheck(RS("TaxAmount")),4,-2,0,0) & "','" & FormatNumber(NullCheck(RS("LineItemTotal")),4,-2,0,0) & "','" & JSEncode(RS("Comments")) & "');")+nl
			'Response.Write("	}")+nl

            If NullCheck(RS("LocationPK")) = "" Then
                sodi = "(Directly Issued)"
            Else
                sodi = JSEncode(RS("LocationID")) & "<br>" & JSEncode(RS("LocationName"))
            End If

		    '@$CUSTOMISED
		    If NullCheck(RS("Bin")) = "" Then
		    	bindata = "(Not Specified)"
		    Else
		    	bindata = JSEncode(RS("Bin"))
		    End If
		    '@$END

			If Trim(UCase(RS("PartID"))) = "MANUAL ENTRY" Then
			Response.Write("	top.builddatarow(myframe.oli1body,4,null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','" & RS("PartPK") & "','IN',false,'" & JSEncode(RS("Photo")) & "',null,null,'" & NullCheck(RS("LineItemNo")) & "','" & JSEncode(PartIDorInternalID) & "','" & JSEncode(RS("PartName")) & "','" & sodi & "','" & bindata & "','" & JSEncode(RS("VendorPartNumber")) & "','" & JSEncode(RS("WOID")) & "','" & JSEncode(RS("AssetID")) & "<mcbr> / " & JSEncode(RS("AssetName")) & "','" & JSEncode(RS("AccountID")) & "<mcbr> / " & JSEncode(RS("AccountName")) & "','" & JSEncode(RS("SubAccountID")) & "<mcbr> / " & JSEncode(RS("SubAccountName")) & "','" & JSEncode(RS("orderunits")) & "<mcbr> / " & JSEncode(RS("orderunitsdesc")) & "','" & NullCheck(RS("ConversionToIssueUnits")) & "','" & JSEncode(RS("IssueUnits")) & "<mcbr> / " & JSEncode(RS("IssueUnitsDesc")) & "','" & FormatNumber(NullCheck(RS("OrderUnitPrice")),4,-2,0,0) & "','" & NullCheck(RS("OrderUnitQty")) & "','" & NullCheck(RS("OrderUnitQtyReceived")) & "','" & NullCheck(RS("OrderUnitQtyBackOrdered")) & "','" & NullCheck(RS("Discount")) & "','" & FormatNumber(NullCheck(RS("Subtotal")),4,-2,0,0) & "','" & taxable & "','" & NullCheck(RS("TaxRate")) & "','" & FormatNumber(NullCheck(RS("TaxAmount")),4,-2,0,0) & "','" & FormatNumber(NullCheck(RS("LineItemTotal")),4,-2,0,0) & "','" & DateNullCheck(RS("DueDate")) & "','" & JSEncode(RS("Comments")) & "');")+nl
			Else
			Response.Write("	top.builddatarow(myframe.oli1body,3,null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','" & RS("PartPK") & "','IN',false,'" & JSEncode(RS("Photo")) & "',null,null,'" & NullCheck(RS("LineItemNo")) & "','" & JSEncode(PartIDorInternalID) & "','" & JSEncode(RS("PartName")) & "','" & sodi & "','" & bindata & "','" & JSEncode(RS("VendorPartNumber")) & "','" & JSEncode(RS("WOID")) & "','" & JSEncode(RS("AssetID")) & "<mcbr> / " & JSEncode(RS("AssetName")) & "','" & JSEncode(RS("AccountID")) & "<mcbr> / " & JSEncode(RS("AccountName")) & "','" & JSEncode(RS("SubAccountID")) & "<mcbr> / " & JSEncode(RS("SubAccountName")) & "','" & JSEncode(RS("orderunits")) & "<mcbr> / " & JSEncode(RS("orderunitsdesc")) & "','" & NullCheck(RS("ConversionToIssueUnits")) & "','" & JSEncode(RS("IssueUnits")) & "<mcbr> / " & JSEncode(RS("IssueUnitsDesc")) & "','" & FormatNumber(NullCheck(RS("OrderUnitPrice")),4,-2,0,0) & "','" & NullCheck(RS("OrderUnitQty")) & "','" & NullCheck(RS("OrderUnitQtyReceived")) & "','" & NullCheck(RS("OrderUnitQtyBackOrdered")) & "','" & NullCheck(RS("Discount")) & "','" & FormatNumber(NullCheck(RS("Subtotal")),4,-2,0,0) & "','" & taxable & "','" & NullCheck(RS("TaxRate")) & "','" & FormatNumber(NullCheck(RS("TaxAmount")),4,-2,0,0) & "','" & FormatNumber(NullCheck(RS("LineItemTotal")),4,-2,0,0) & "','" & DateNullCheck(RS("DueDate")) & "','" & JSEncode(RS("Comments")) & "');")+nl
			End If

			rs.MoveNext
		Loop

		Response.Write(nl)
		Response.Write("parent.PORecalc('');")+nl
		Response.Write(nl)

	Response.Write("}")+nl
	Response.Write(nl)

	Response.Write("	myframe.po_receipts.setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab63()")+nl
	Response.Write("{")+nl

		Response.Write(nl)
		Response.Write("	// Write Receipts Data")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	myframe.fraReceipts.location.replace(top.path+'modules/purchaseorder/purchaseorder_receipts.asp?kv='+top.recordkey.replace('kv=','')+'&t=R');") + nl
        Response.Write("    myframe.resizeit();")+nl
		Response.Write("	myframe.document.getElementById('fraReceipts').style.display = '';") + nl

		Response.Write(nl)
		Response.Write("	// Write Invoices Data")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	myframe.fraInvoices.location.replace(top.path+'modules/purchaseorder/purchaseorder_receipts.asp?kv='+top.recordkey.replace('kv=','')+'&t=I');") + nl
        Response.Write("    myframe.resizeit();")+nl
		Response.Write("	myframe.document.getElementById('fraInvoices').style.display = '';") + nl

	Response.Write("}")+nl
	Response.Write(nl)

	Response.Write("	myframe.po_rma.setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab64()")+nl
	Response.Write("{")+nl

		Response.Write(nl)
		Response.Write("	// Write RMA Data")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	myframe.fraRMA.location.replace(top.path+'modules/purchaseorder/purchaseorder_rma.asp?kv='+top.recordkey.replace('kv=',''));") + nl
        Response.Write("    myframe.resizeit();")+nl
		Response.Write("	myframe.document.getElementById('fraRMA').style.display = '';") + nl

	Response.Write("}")+nl
	Response.Write(nl)

	Response.Write("	myframe.po_attach.setAttribute('dataloaded','N');") + nl
	Response.Write(nl)
	Response.Write("function loadtab65()")+nl
	Response.Write("{")+nl

		Set rs = rs.NextRecordset

		Response.Write(nl)
		Response.Write("	// Build Attachments Rows")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	top.cleartable(myframe.oat1);")+nl

		Call OutputAttachments(rs)

		Set rs = rs.NextRecordset

		Response.Write(nl)
		Response.Write("	// Build Note Rows")+nl
		Response.Write("	// -------------------------------------------------------------------------")+nl
		Response.Write("	top.cleartable(myframe.ono1);")+nl
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
				Response.Write("	top.builddatarow(myframe.ono1body," & NoteTemplate & ",null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','','',false,'',null,null,'" & DateNullCheckAT(RS("NoteDate")) & "','" & TimeNullCheckAT(RS("NoteDate")) & "','" & JSEncode(RS("Initials")) & "','" & JSEncode(RS("Note")) & "','" & NoteCustom1 & "','" & NoteCustom2 & "');")+nl
			Else
				Response.Write("	top.builddatarow(myframe.ono1body," & NoteTemplate & ",null,'" & NullCheck(RS("PK")) + "$" + NullCheck(RS("RowVersionDate")) & "','','',false,'',null,null,'" & DateNullCheckAT(RS("NoteDate")) & "','" & TimeNullCheckAT(RS("NoteDate")) & "','" & JSEncode(RS("Initials")) & "','" & JSEncode(RS("Note")) & "','" & NoteCustom1 & "','" & NoteCustom2 & "');")+nl
			End If
			rs.MoveNext
		Loop

        Response.Write(nl)
	    Response.Write("// Build Misc Attachments")+nl
	    Response.Write("// -------------------------------------------------------------------------")+nl
        'Response.Write("	myframe.fileFrame.location.replace(top.path+'modules/attachments/mc_genericFilesMain.asp?pid="&keyvalue&"&table=PO');")+nl					
        If Not newrecord Then
	    'Response.Write("	myframe.document.getElementById('fileFrame').style.display = '';") + nl
        End If

        Response.Write("outputrules();") + nl

		Response.Write(nl)

		Response.Write("}")+nl

		Response.Write("	myframe.po_report.setAttribute('dataloaded','N');") + nl
		Response.Write(nl)
		Response.Write("function loadtab66()")+nl
		Response.Write("{")+nl

			Response.Write(nl)
			Response.Write("	// Write Report Data")+nl
			Response.Write("	// -------------------------------------------------------------------------")+nl
			Response.Write("	top.showreportiniframe('PurchaseOrder','WHERE PurchaseOrder.POPK='+top.recordkey.replace('kv=',''),myframe.fraReports);") + nl
            Response.Write("    myframe.resizeit();")+nl
			Response.Write("	myframe.document.getElementById('fraReports').style.display = '';") + nl

		Response.Write("}")+nl
		Response.Write(nl)
	Set rs = rs.NextRecordset
			Response.write("myform.POName.value='';")
			Response.write("myform.POExport.value='';")
		If not rs.eof Then
			Response.write("myform.POName.value='"&rs("zPO")&"';")
			while not (rs.eof)
			Response.write("myform.POExport.value+='"&rs("exportString")&"\r\n';")
			rs.MoveNext
			wend
			else
			Response.write("myform.POName.value='No_data_found';")
			Response.write("myform.POExport.value='No_data_found';")
		end if
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

	If CInt(section) = 1 or CInt(section) = 5 Then

		' CUSTOMIZED
		'--------------------------------------------------------------------------------------------------------------
		Response.Write("	myform.txtPOName.value = '" & JSEncode(RS("POName")) & "';")+nl
		'--------------------------------------------------------------------------------------------------------------
        Response.Write("try{myframe.ForceOverride(top.recordkey.replace('kv=',''));} catch(e){}")+nl

		Response.Write("	myform.txtCurrency.value = '" & JSEncode(RS("Currency")) & "';")+nl
		If NullCheck(RS("CurrencySymbol")) = "" Then
		Response.Write("	myform.txtCurrencySymbol.value = top.basecurrencysymbol;")+nl
		Else
		Response.Write("	myform.txtCurrencySymbol.value = '" & JSEncode(RS("CurrencySymbol")) & "';")+nl
		End If

		If BitNullCheck(RS("PrintedBox")) = 0 Then
		    Response.Write("	myform.txtPrintedBox.checked = false;")+nl
		Else
		    Response.Write("	myform.txtPrintedBox.checked = true;")+nl
		End If

		If BitNullCheck(RS("IsPartsOrdered")) Then
			Response.Write("myform.txtIsPartsOrdered.value = 'Y';")
        Else
			Response.Write("myform.txtIsPartsOrdered.value = 'N';")
        End If

		If RS("AuthLevelsRequired") = 0 Then
			Response.Write("myframe.authorizationheader.innerText = 'Approval';")
			Response.Write("myform.txtIsApproved.value = 'Y';")
		Else
			If RS("IsApproved") Then
				Response.Write("myframe.authorizationheader.innerText = 'Approval: Complete';")
				Response.Write("myform.txtIsApproved.value = 'Y';")
			Else
				Response.Write("myframe.authorizationheader.innerText = 'Approval: Pending';")
				Response.Write("myform.txtIsApproved.value = 'N';")
			End If
		End If

		If BitNullCheck(RS("IsApproved")) and rs("AuthLevelsRequired") > 0 Then
			Response.Write("	myframe.approvednoedit = true; ")+nl
		Else
			Response.Write("	myframe.approvednoedit = false; ")+nl
		End IF

		If rs("AuthLevelsRequired") > 0 and Not BitNullCheck(RS("IsApproved")) Then
			Response.Write("	myframe.approvednoprint = true; ")+nl
		Else
			Response.Write("	myframe.approvednoprint = false; ")+nl
		End If

        If rs("status") = "REQUESTED" Then
			Response.Write("	myframe.notissuedyet = true; ")+nl
        Else
			Response.Write("	myframe.notissuedyet = false; ")+nl
        End If

		Response.Write("	myframe.OrderUnitQty_TOTAL = '" & JSEncode(rs("OrderUnitQty_TOTAL")) & "'; ")+nl
		Response.Write("	myframe.OrderUnitQtyReceived_TOTAL = '" & JSEncode(rs("OrderUnitQtyReceived_TOTAL")) & "'; ")+nl

	End If
	If CInt(section) = 2 or CInt(section) = 5 Then

		' CUSTOMIZED
		'--------------------------------------------------------------------------------------------------------------
		Response.Write("	top.sethidden('txtStatus','" & JSEncode(RS("Status")) & "');")+nl
		Response.Write("	top.sethidden('txtAuthStatus','" & JSEncode(RS("AuthStatus")) & "');")+nl

		'--------------------------------------------------------------------------------------------------------------

		Response.Write("	myform.txtPODate.value = '" & DateNullCheck(RS("PODate")) & "';")+nl
		Response.Write("	top.setpk('txtBuyerCompanyPK','" & RS("BuyerCompanyPK") & "');")+nl
		Response.Write("	myform.txtBuyerCompany.value = '" & JSEncode(RS("BuyerCompanyID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.txtBuyerCompanyDesc,'" & JSEncode(RS("BuyerCompanyName")) & "');")+nl

		Response.Write("	top.setpk('txtBuyerPK','" & RS("BuyerPK") & "');")+nl
		Response.Write("	myform.txtBuyer.value = '" & JSEncode(RS("BuyerID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.txtBuyerDesc,'" & JSEncode(RS("BuyerName")) & "');")+nl

		Response.Write("	top.setpk('txtRequesterPK','" & RS("RequesterPK") & "');")+nl
		Response.Write("	myform.txtRequester.value = '" & JSEncode(RS("RequesterID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.txtRequesterDesc,'" & JSEncode(RS("RequesterName")) & "');")+nl

		Response.Write("	top.setpk('txtTenantPK','" & RS("TenantPK") & "');")+nl
		Response.Write("	myform.txtTenant.value = '" & JSEncode(RS("TenantID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.txtTenantDesc,'" & JSEncode(RS("TenantName")) & "');")+nl

		Response.Write("	if (myform.txtAccount) {top.setpk('txtAccountPK','" & RS("AccountPK") & "');")+nl
		Response.Write("	myform.txtAccount.value = '" & JSEncode(RS("AccountID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.txtAccountDesc,'" & JSEncode(RS("AccountName")) & "');")+nl

		Response.Write("	top.setpk('txtDepartmentPK','" & RS("DepartmentPK") & "');")+nl
		Response.Write("	myform.txtDepartment.value = '" & JSEncode(RS("DepartmentID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.txtDepartmentDesc,'" & JSEncode(RS("DepartmentName")) & "');}")+nl

		'Response.Write("	top.setpk('txtDepartmentPK','" & RS("DepartmentPK") & "');")+nl
		'Response.Write("	myform.txtDepartment.value = '" & JSEncode(RS("DepartmentID")) & "';")+nl
		'Response.Write("	top.setdesc(myframe.txtDepartmentDesc,'" & JSEncode(RS("DepartmentName")) & "');")+nl

		Response.Write("	if (myform.txtSubStatus) {myform.txtSubStatus.value = '" & JSEncode(RS("SubStatus")) & "';")+nl
		Response.Write("	top.setdesc(myframe.document.getElementById('txtSubStatusDesc'),'" & JSEncode(RS("SubStatusDesc")) & "');}")+nl

		Response.Write("	myform.txtFollowupDate.value = '" & DateNullCheck(RS("FollowupDate")) & "';")+nl
		Response.Write("	myform.txtShipDate.value = '" & DateNullCheck(RS("ShipDate")) & "';")+nl
		Response.Write("	myform.txtTrackingNo.value = '" & JSEncode(RS("TrackingNo")) & "';")+nl
		Response.Write("	top.setpk('txtVendorPK','" & RS("VendorPK") & "');")+nl
		Response.Write("	myform.txtVendor.value = '" & JSEncode(RS("VendorID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.txtVendorDesc,'" & JSEncode(RS("VendorName")) & "');")+nl
        Response.Write("	myframe.txtVendorPhone.innerText = '" & JSEncode(RS("VendorPhone")) & "';")+nl
		Response.Write("	top.setpk('txtRepairCenterPK','" & RS("RepairCenterPK") & "');")+nl
		Response.Write("	myform.txtRepairCenter.value = '" & JSEncode(RS("RepairCenterID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.txtRepairCenterDesc,'" & JSEncode(RS("RepairCenterName")) & "');")+nl
		Response.Write("	myform.txtPriority.value = '" & JSEncode(RS("Priority")) & "';")+nl
		Response.Write("	top.setdesc(myframe.txtPriorityDesc,'" & JSEncode(RS("PriorityDesc")) & "');")+nl
		Response.Write("	myform.txtInvoiceNumber.value = '" & JSEncode(RS("InvoiceNumber")) & "';")+nl
		Response.Write("	myform.txtFreightTerms.value = '" & JSEncode(RS("FreightTerms")) & "';")+nl
		Response.Write("	top.setdesc(myframe.txtFreightTermsDesc,'" & JSEncode(RS("FreightTermsDesc")) & "');")+nl
		Response.Write("	myform.txtShippingMethod.value = '" & JSEncode(RS("ShippingMethod")) & "';")+nl
		Response.Write("	top.setdesc(myframe.txtShippingMethodDesc,'" & JSEncode(RS("ShippingMethodDesc")) & "');")+nl
		Response.Write("	myform.txtSubtotal.value = '" & FormatNumber(RS("Subtotal"),4,-2,0,0) & "';")+nl
		Response.Write("	myform.txtSubTotalPreDiscount.value = '" & FormatNumber(NumericNullCheck(RS("SubtotalPostDiscount")),4,-2,0,0) & "';")+nl
		Response.Write("	myform.txtFreightCharge.value = '" & FormatNumber(RS("FreightCharge"),4,-2,0,0) & "';")+nl
		Response.Write("	myform.txtTaxAmount.value = '" & FormatNumber(RS("TaxAmount"),4,-2,0,0) & "';")+nl
		Response.Write("	myform.txtDiscountPercentage.value = '" & FormatNumber(NumericNullCheck(RS("DiscountPercentage")),4,-2,0,0) & "';")+nl
		Response.Write("	myform.txtDiscount.value = '" & FormatNumber(NumericNullCheck(RS("Discount")),4,-2,0,0) & "';")+nl
		Response.Write("	myform.txtTotal.value = '" & FormatNumber(RS("Total"),4,-2,0,0) & "';")+nl
		Response.Write("	myform.txtTerms.value = '" & JSEncode(RS("Terms")) & "';")+nl
		Response.Write("	top.setdesc(myframe.txtTermsDesc,'" & JSEncode(RS("TermsDesc")) & "');")+nl
		'Response.Write("	myform.txtTermsInfo.value = '" & JSEncode(RS("TermsInfo")) & "';")+nl
		Response.Write("	top.setpk('txtShipToPK','" & RS("ShipToPK") & "');")+nl
		Response.Write("	myform.txtShipTo.value = '" & JSEncode(RS("ShipToID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.txtShipToDesc,'" & JSEncode(RS("ShipToName")) & "');")+nl
		Response.Write("	myform.txtShipToAttention.value = '" & JSEncode(RS("ShipToAttention")) & "';")+nl
		Response.Write("	myform.txtShipToAddress1.value = '" & JSEncode(RS("ShipToAddress1")) & "';")+nl
		Response.Write("	myform.txtShipToAddress2.value = '" & JSEncode(RS("ShipToAddress2")) & "';")+nl
		Response.Write("	myform.txtShipToAddress3.value = '" & JSEncode(RS("ShipToAddress3")) & "';")+nl
		Response.Write("	top.setpk('txtBillToPK','" & RS("BillToPK") & "');")+nl
		Response.Write("	myform.txtBillTo.value = '" & JSEncode(RS("BillToID")) & "';")+nl
		Response.Write("	top.setdesc(myframe.txtBillToDesc,'" & JSEncode(RS("BillToName")) & "');")+nl
		Response.Write("	myform.txtBillToAttention.value = '" & JSEncode(RS("BillToAttention")) & "';")+nl
		Response.Write("	myform.txtBillToAddress1.value = '" & JSEncode(RS("BillToAddress1")) & "';")+nl
		Response.Write("	myform.txtBillToAddress2.value = '" & JSEncode(RS("BillToAddress2")) & "';")+nl
		Response.Write("	myform.txtBillToAddress3.value = '" & JSEncode(RS("BillToAddress3")) & "';")+nl
		'Response.Write("	top.setpk('txtTakenByPK','" & RS("TakenByPK") & "');")+nl
		'Response.Write("	myform.txtTakenByInitials.value = '" & JSEncode(RS("TakenByInitials")) & "';")+nl
		'Response.Write("	top.setpk('txtAuthStatusUserPK','" & RS("AuthStatusUserPK") & "');")+nl
		'Response.Write("	myform.txtAuthStatusUserInitials.value = '" & JSEncode(RS("AuthStatusUserInitials")) & "';")+nl
		'Response.Write("	myform.txtAuthStatusDate.value = '" & DateNullCheck(RS("AuthStatusDate")) & "';")+nl
		'Response.Write("	myform.txtRequested.value = '" & DateNullCheck(RS("Requested")) & "';")+nl
		'Response.Write("	myform.txtIssued.value = '" & DateNullCheck(RS("Issued")) & "';")+nl
		'Response.Write("	myform.txtOnHold.value = '" & DateNullCheck(RS("OnHold")) & "';")+nl
		'Response.Write("	myform.txtClosed.value = '" & DateNullCheck(RS("Closed")) & "';")+nl
		'Response.Write("	myform.txtCanceled.value = '" & DateNullCheck(RS("Canceled")) & "';")+nl
		'Response.Write("	myform.txtDenied.value = '" & DateNullCheck(RS("Denied")) & "';")+nl
		'If RS("IsOpen") = 0 Then
		'Response.Write("	myform.txtIsOpen.checked = false;")+nl
		'Else
		'Response.Write("	myform.txtIsOpen.checked = true;")+nl
		'End If
        'Response.Write("alert(top.$('#cb2Debug'));")+nl
		' CUSTOMIZED
		'--------------------------------------------------------------------------------------------------------------
		'Response.Write("	myform.txtComments.value = unescape('" & JSEncode(RS("Comments")) & "');")+nl
		'--------------------------------------------------------------------------------------------------------------

		'Response.Write("	myform.txtPhoto.value = '" & JSEncode(RS("Photo")) & "';")+nl

		' CUSTOMIZED
		'--------------------------------------------------------------------------------------------------------------
		'Response.Write("	top.displayphoto(mydoc.images.purchaseorderphoto,'" & JSEncode(RS("Photo")) & "');")+nl
		'--------------------------------------------------------------------------------------------------------------
		Call OutputUDFData(rs)

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

End Function

Function validate_no1(thing,theindex)

End Function

Function validate_li1(thing,theindex)

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

	Set rs = db.RunSqlReturnRS("Select GETDATE() AS ServerDate","")
	If Not db.dok Then
		Exit Sub
	End If
	ServerDate = rs("ServerDate")

	If newrecord or duprecord Then
		Set rs = db.RunSQLReturnRS_RW("SELECT TOP 0 * FROM PurchaseOrder","")
		If Not db.dok Then
			Exit Sub
		End If
		rs.AddNew
	Else
		Set rs = db.RunSQLReturnRS_RW("SELECT * FROM PurchaseOrder WHERE POPK=" & NullCheck(keyvalue),"")
		If Not db.dok Then
			Exit Sub
		End If
	End If

	If rs.eof Then
		' Looks like another user deleted the record first
		db.dok = False
		db.derror = "Another user has deleted this Purchase Order while you were working with it. If you feel this is not the case, you can try again, otherwise please click the CANCEL button to cancel out of this Purchase Order."
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
			If InStr("STATUS",rva) > 0 Then
				ConflictIsStatus = True
			ElseIf InStr("AUTHSTATUS",rva) > 0 Then
				ConflictIsAuthStatus = True
			'ElseIf InStr("ASSIGN",Trim(rs("RowVersionAction"))) > 0 Then
			'	ConflictIsAssign = True
			ElseIf InStr("PRINT",rva) > 0 Then
				' No Need to Set Conflict Var because they can't change the Print flag anyway
			Else
				db.dok = False
				db.derror = "Another user has made modifications to this Purchase Order while you were working with it. Since you could potentially overwrite changes the other user made, you will need to Cancel your changes and start again. Please click the CANCEL button to cancel your changes."
				Exit Sub
			End If
		Else
			db.dok = False
			db.derror = "Another user has made modifications to this Purchase Order while you were working with it. Since you could potentially overwrite changes the other user made, you will need to Cancel your changes and start again. Please click the CANCEL button to cancel your changes."
			Exit Sub
		End If
	End If

	' CUSTOMIZED
	'--------------------------------------------------------------------------------------------------------------
	'rs("POID") = Trim(Mid(Request.Form("txtPO").Item,1,25))	' Nullable: No Type: nvarchar
	If newrecord or duprecord Then
		rs("POID") = ""
	End If
	'--------------------------------------------------------------------------------------------------------------

	' CUSTOMIZED
	'--------------------------------------------------------------------------------------------------------------
	If Not IsEmpty(Request.Form("txtPOName").Item) Then
		rs("POName") = Trim(Mid(Request.Form("txtPOName").Item,1,50))	' Nullable: YES Type: nvarchar
	End If
	'--------------------------------------------------------------------------------------------------------------

	' CUSTOMIZED
	'--------------------------------------------------------------------------------------------------------------
	If EditConflict and (ConflictIsStatus or ConflictIsAuthStatus) Then
		' Use Current Status
		If ConflictIsStatus Then
			If UCase(Request.Form("statuschanged")) = "Y" and _
			   Not rs("Status") = Trim(Mid(Request.Form("txtStatus").Item,1,15))Then
					db.warn = True
					db.warntext = Trim(db.warntext + " Another user has updated the Status of this Purchase Order while you were making your changes. <font color=""red"">Your changing of the Status to " & Trim(Mid(Request.Form("txtStatusDescH").Item,1,50)) & " for this Purchase Order was NOT saved.</font> Please check that the appropriate Status has been set for this Purchase Order.")
			End If
		End If
		If ConflictIsAuthStatus Then
			If UCase(Request.Form("authstatuschanged")) = "Y" and _
			   Not rs("AuthStatus") = Trim(Mid(Request.Form("txtAuthStatus").Item,1,15)) Then
				db.warn = True
				db.warntext = Trim(db.warntext + " Another user has updated the Authorization Status of this Purchase Order while you were making your changes. <font color=""red"">Your changing of the Authorization Status to " & Trim(Mid(Request.Form("txtStatus").Item,1,15)) & " for this Purchase Order was NOT saved.</font> Please check that the appropriate Authorization Status has been set for this Purchase Order.")
			End If
		End If
	Else
		rs("Status") = Trim(Mid(Request.Form("txtStatus").Item,1,15))	' Nullable: No Type: nvarchar
		If Not IsEmpty(Request.Form("txtStatusDescH").Item) Then
			rs("StatusDesc") = Trim(Mid(Request.Form("txtStatusDescH").Item,1,50))	' Nullable: YES Type: nvarchar
		End If
        Dim returnString
        returnString = GetRequiredStatus(db, keyvalue, FixInt(Request("txtRepairCenterPK")))

        If Trim(returnString) <> "" Then
		    rs("AuthStatus") = returnString 'Trim(Mid(Request.Form("txtAuthStatus").Item,1,15))	' Nullable: No Type: nvarchar
        End If
		If Not IsEmpty(Request.Form("txtAuthStatusDescH").Item) Then
			rs("AuthStatusDesc") = Trim(Mid(Request.Form("txtAuthStatusDescH").Item,1,50))	' Nullable: YES Type: nvarchar
		End If
		If Not IsEmpty(Request.Form("txtStatusDate").Item) Then
			If Not Request.Form("txtStatusDate").Item = "" Then
				rs("StatusDate") = SQLdatetimeADO(Request.Form("txtStatusDate").Item)	' Nullable: YES Type: datetime
			Else
				rs("StatusDate") = Null
			End If
		Else
			rs("StatusDate") = SQLdatetimeADO(ServerDate)	' Nullable: YES Type: datetime
		End If
	End If
	'--------------------------------------------------------------------------------------------------------------

	If Not IsEmpty(Request.Form("txtCurrency").Item) Then
		If Len(Trim(Request.Form("txtCurrency").Item)) > 0 Then
			rs("Currency") = Trim(Mid(Request.Form("txtCurrency").Item,1,25))	' Nullable: YES Type: nvarchar
		Else
			rs("Currency") = Null
		End If
	End If

	' CUSTOMIZED
	'-----------------------------------------------------------------------------------------------------------------
	rs("PrintedBox") = Not Request.Form("txtPrintedBox") = ""	' Nullable: No Type: bit
	'-----------------------------------------------------------------------------------------------------------------

	rs("PODate") = SQLdatetimeADO(Request.Form("txtPODate").Item)	' Nullable: No Type: datetime
	If Not IsEmpty(Request.Form("txtBuyerCompanyPK").Item) Then
		If Len(Trim(Request.Form("txtBuyerCompanyPK").Item)) > 0 Then
			rs("BuyerCompanyPK") = Request.Form("txtBuyerCompanyPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtBuyerCompany").Item) Then
				rs("BuyerCompanyID") = Trim(Mid(Request.Form("txtBuyerCompany").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtBuyerCompanyDescH").Item) Then
				rs("BuyerCompanyName") = Trim(Mid(Request.Form("txtBuyerCompanyDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("BuyerCompanyPK") = Null
			rs("BuyerCompanyID") = Null
			rs("BuyerCompanyName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtBuyerPK").Item) Then
		If Len(Trim(Request.Form("txtBuyerPK").Item)) > 0 Then
			rs("BuyerPK") = Request.Form("txtBuyerPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtBuyer").Item) Then
				rs("BuyerID") = Trim(Mid(Request.Form("txtBuyer").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtBuyerDescH").Item) Then
				rs("BuyerName") = Trim(Mid(Request.Form("txtBuyerDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("BuyerPK") = Null
			rs("BuyerID") = Null
			rs("BuyerName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtRequesterPK").Item) Then
		If Len(Trim(Request.Form("txtRequesterPK").Item)) > 0 Then
			rs("RequesterPK") = Request.Form("txtRequesterPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtRequester").Item) Then
				rs("RequesterID") = Trim(Mid(Request.Form("txtRequester").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtRequesterDescH").Item) Then
				rs("RequesterName") = Trim(Mid(Request.Form("txtRequesterDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("RequesterPK") = Null
			rs("RequesterID") = Null
			rs("RequesterName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtTenantPK").Item) Then
		If Len(Trim(Request.Form("txtTenantPK").Item)) > 0 Then
			rs("TenantPK") = Request.Form("txtTenantPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtTenant").Item) Then
				rs("TenantID") = Trim(Mid(Request.Form("txtTenant").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtTenantDescH").Item) Then
				rs("TenantName") = Trim(Mid(Request.Form("txtTenantDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("TenantPK") = Null
			rs("TenantID") = Null
			rs("TenantName") = Null
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
	If Not IsEmpty(Request.Form("txtDepartmentPK").Item) Then
		If Len(Trim(Request.Form("txtDepartmentPK").Item)) > 0 Then
			rs("DepartmentPK") = Request.Form("txtDepartmentPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtDepartment").Item) Then
				rs("DepartmentID") = Trim(Mid(Request.Form("txtDepartment").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtDepartmentDescH").Item) Then
				rs("DepartmentName") = Trim(Mid(Request.Form("txtDepartmentDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("DepartmentPK") = Null
			rs("DepartmentID") = Null
			rs("DepartmentName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtSubStatus")) Then
		If Len(Trim(Request.Form("txtSubStatus"))) > 0 Then
			rs("SubStatus") = Trim(Mid(Request.Form("txtSubStatus"),1,50))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtSubStatusDescH")) Then
				rs("SubStatusDesc") = Trim(Mid(Request.Form("txtSubStatusDescH"),1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("SubStatus") = Null
			rs("SubStatusDesc") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtFollowupDate").Item) Then
		If Not Request.Form("txtFollowupDate").Item = "" Then
			rs("FollowupDate") = SQLdatetimeADO(Request.Form("txtFollowupDate").Item)	' Nullable: YES Type: datetime
		Else
			rs("FollowupDate") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtShipDate").Item) Then
		If Not Request.Form("txtShipDate").Item = "" Then
			rs("ShipDate") = SQLdatetimeADO(Request.Form("txtShipDate").Item)	' Nullable: YES Type: datetime
		Else
			rs("ShipDate") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtTrackingNo").Item) Then
		If Len(Trim(Request.Form("txtTrackingNo").Item)) > 0 Then
			rs("TrackingNo") = Trim(Mid(Request.Form("txtTrackingNo").Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("TrackingNo") = Null
		End If
	End If
	If Len(Trim(Request.Form("txtVendorPK").Item)) > 0 Then
		rs("VendorPK") = Request.Form("txtVendorPK").Item	' Nullable: No Type: int
		rs("VendorID") = Trim(Mid(Request.Form("txtVendor").Item,1,25))	' Nullable: No Type: nvarchar
		If Not IsEmpty(Request.Form("txtVendorDescH").Item) Then
			rs("VendorName") = Trim(Mid(Request.Form("txtVendorDescH").Item,1,50))	' Nullable: YES Type: nvarchar
		End If
	Else
		rs("VendorPK") = Null
		rs("VendorID") = Null
		rs("VendorName") = Null
	End If
	If Not IsEmpty(Request.Form("txtRepairCenterPK").Item) Then
		If Len(Trim(Request.Form("txtRepairCenterPK").Item)) > 0 Then
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
	If Not IsEmpty(Request.Form("txtPriority").Item) Then
		If Len(Trim(Request.Form("txtPriority").Item)) > 0 Then
			rs("Priority") = Trim(Mid(Request.Form("txtPriority").Item,1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtPriorityDescH").Item) Then
				rs("PriorityDesc") = Trim(Mid(Request.Form("txtPriorityDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("Priority") = Null
			rs("PriorityDesc") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtInvoiceNumber").Item) Then
		If Len(Trim(Request.Form("txtInvoiceNumber").Item)) > 0 Then
			rs("InvoiceNumber") = Trim(Mid(Request.Form("txtInvoiceNumber").Item,1,30))	' Nullable: YES Type: nvarchar
		Else
			rs("InvoiceNumber") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtFreightTerms").Item) Then
		If Len(Trim(Request.Form("txtFreightTerms").Item)) > 0 Then
			rs("FreightTerms") = Trim(Mid(Request.Form("txtFreightTerms").Item,1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtFreightTermsDescH").Item) Then
				rs("FreightTermsDesc") = Trim(Mid(Request.Form("txtFreightTermsDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("FreightTerms") = Null
			rs("FreightTermsDesc") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtShippingMethod").Item) Then
		If Len(Trim(Request.Form("txtShippingMethod").Item)) > 0 Then
			rs("ShippingMethod") = Trim(Mid(Request.Form("txtShippingMethod").Item,1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtShippingMethodDescH").Item) Then
				rs("ShippingMethodDesc") = Trim(Mid(Request.Form("txtShippingMethodDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("ShippingMethod") = Null
			rs("ShippingMethodDesc") = Null
		End If
	End If
	If Len(Trim(Request.Form("txtSubtotal").Item)) > 0 Then
		rs("Subtotal") = FixInternationalNumber(Request.Form("txtSubtotal").Item)	' Nullable: No Type: money
		rs("SubtotalPreDiscount") = FixInternationalNumber(Request.Form("txtSubtotal").Item)	' Nullable: No Type: money
	End If
	If Len(Trim(Request.Form("txtSubtotalPreDiscount").Item)) > 0 Then
		rs("SubtotalPostDiscount") = FixInternationalNumber(Request.Form("txtSubtotalPreDiscount").Item)	' Nullable: No Type: money
	End If
	If Len(Trim(Request.Form("txtFreightCharge").Item)) > 0 Then
		rs("FreightCharge") = FixInternationalNumber(Request.Form("txtFreightCharge").Item)	' Nullable: No Type: money
	End If
	If Len(Trim(Request.Form("txtTaxAmount").Item)) > 0 Then
		rs("TaxAmount") = FixInternationalNumber(Request.Form("txtTaxAmount").Item)	' Nullable: No Type: real
	End If
	If Len(Trim(Request.Form("txtDiscountPercentage").Item)) > 0 Then
		rs("DiscountPercentage") = FixInternationalNumber(Request.Form("txtDiscountPercentage").Item)	' Nullable: No Type: money
	End If
	If Len(Trim(Request.Form("txtDiscount").Item)) > 0 Then
		rs("Discount") = FixInternationalNumber(Request.Form("txtDiscount").Item)	' Nullable: No Type: money
	End If
	If Len(Trim(Request.Form("txtTotal").Item)) > 0 Then
		rs("Total") = FixInternationalNumber(Request.Form("txtTotal").Item)	' Nullable: No Type: money
	End If
	If Not IsEmpty(Request.Form("txtTerms").Item) Then
		If Len(Trim(Request.Form("txtTerms").Item)) > 0 Then
			rs("Terms") = Trim(Mid(Request.Form("txtTerms").Item,1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form("txtTermsDescH").Item) Then
				rs("TermsDesc") = Trim(Mid(Request.Form("txtTermsDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("Terms") = Null
			rs("TermsDesc") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtTermsInfo").Item) Then
		If Len(Trim(Request.Form("txtTermsInfo").Item)) > 0 Then
			rs("TermsInfo") = Trim(Mid(Request.Form("txtTermsInfo").Item,1,2000))	' Nullable: YES Type: nvarchar
		Else
			rs("TermsInfo") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtShipToPK").Item) Then
		If Len(Trim(Request.Form("txtShipToPK").Item)) > 0 Then
			rs("ShipToPK") = Request.Form("txtShipToPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtShipTo").Item) Then
				rs("ShipToID") = Trim(Mid(Request.Form("txtShipTo").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtShipToDescH").Item) Then
				rs("ShipToName") = Trim(Mid(Request.Form("txtShipToDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("ShipToPK") = Null
			rs("ShipToID") = Null
			rs("ShipToName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtShipToAttention").Item) Then
		If Len(Trim(Request.Form("txtShipToAttention").Item)) > 0 Then
			rs("ShipToAttention") = Trim(Mid(Request.Form("txtShipToAttention").Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("ShipToAttention") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtShipToAddress1").Item) Then
		If Len(Trim(Request.Form("txtShipToAddress1").Item)) > 0 Then
			rs("ShipToAddress1") = Trim(Mid(Request.Form("txtShipToAddress1").Item,1,80))	' Nullable: YES Type: nvarchar
		Else
			rs("ShipToAddress1") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtShipToAddress2").Item) Then
		If Len(Trim(Request.Form("txtShipToAddress2").Item)) > 0 Then
			rs("ShipToAddress2") = Trim(Mid(Request.Form("txtShipToAddress2").Item,1,80))	' Nullable: YES Type: nvarchar
		Else
			rs("ShipToAddress2") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtShipToAddress3").Item) Then
		If Len(Trim(Request.Form("txtShipToAddress3").Item)) > 0 Then
			rs("ShipToAddress3") = Trim(Mid(Request.Form("txtShipToAddress3").Item,1,80))	' Nullable: YES Type: nvarchar
		Else
			rs("ShipToAddress3") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtBillToPK").Item) Then
		If Len(Trim(Request.Form("txtBillToPK").Item)) > 0 Then
			rs("BillToPK") = Request.Form("txtBillToPK").Item	' Nullable: YES Type: int
			If Not IsEmpty(Request.Form("txtBillTo").Item) Then
				rs("BillToID") = Trim(Mid(Request.Form("txtBillTo").Item,1,25))	' Nullable: YES Type: nvarchar
			End If
			If Not IsEmpty(Request.Form("txtBillToDescH").Item) Then
				rs("BillToName") = Trim(Mid(Request.Form("txtBillToDescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("BillToPK") = Null
			rs("BillToID") = Null
			rs("BillToName") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtBillToAttention").Item) Then
		If Len(Trim(Request.Form("txtBillToAttention").Item)) > 0 Then
			rs("BillToAttention") = Trim(Mid(Request.Form("txtBillToAttention").Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("BillToAttention") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtBillToAddress1").Item) Then
		If Len(Trim(Request.Form("txtBillToAddress1").Item)) > 0 Then
			rs("BillToAddress1") = Trim(Mid(Request.Form("txtBillToAddress1").Item,1,80))	' Nullable: YES Type: nvarchar
		Else
			rs("BillToAddress1") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtBillToAddress2").Item) Then
		If Len(Trim(Request.Form("txtBillToAddress2").Item)) > 0 Then
			rs("BillToAddress2") = Trim(Mid(Request.Form("txtBillToAddress2").Item,1,80))	' Nullable: YES Type: nvarchar
		Else
			rs("BillToAddress2") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtBillToAddress3").Item) Then
		If Len(Trim(Request.Form("txtBillToAddress3").Item)) > 0 Then
			rs("BillToAddress3") = Trim(Mid(Request.Form("txtBillToAddress3").Item,1,80))	' Nullable: YES Type: nvarchar
		Else
			rs("BillToAddress3") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtTakenByInitials").Item) Then
		If Len(Trim(Request.Form("txtTakenByInitials").Item)) > 0 Then
			rs("TakenByInitials") = Trim(Mid(Request.Form("txtTakenByInitials").Item,1,25))	' Nullable: YES Type: nvarchar
		Else
			rs("TakenByInitials") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtAuthStatusUserPK").Item) Then
		If Len(Trim(Request.Form("txtAuthStatusUserPK").Item)) > 0 Then
			rs("AuthStatusUserPK") = Request.Form("txtAuthStatusUserPK").Item	' Nullable: YES Type: int
		Else
			rs("AuthStatusUserPK") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtAuthStatusUserInitials").Item) Then
		If Len(Trim(Request.Form("txttxtAuthStatusUserInitialsTerms").Item)) > 0 Then
			rs("AuthStatusUserInitials") = Trim(Mid(Request.Form("txtAuthStatusUserInitials").Item,1,5))	' Nullable: YES Type: nvarchar
		Else
			rs("AuthStatusUserInitials") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtAuthStatusDate").Item) Then
		If Not Request.Form("txtAuthStatusDate").Item = "" Then
			rs("AuthStatusDate") = SQLdatetimeADO(Request.Form("txtAuthStatusDate").Item)	' Nullable: YES Type: datetime
		Else
			rs("AuthStatusDate") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtRequested").Item) Then
		If Not Request.Form("txtRequested").Item = "" Then
			rs("Requested") = SQLdatetimeADO(Request.Form("txtRequested").Item)	' Nullable: YES Type: datetime
		Else
			rs("Requested") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtIssued").Item) Then
		If Not Request.Form("txtIssued").Item = "" Then
			rs("Issued") = SQLdatetimeADO(Request.Form("txtIssued").Item)	' Nullable: YES Type: datetime
		Else
			rs("Issued") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtOnHold").Item) Then
		If Not Request.Form("txtOnHold").Item = "" Then
			rs("OnHold") = SQLdatetimeADO(Request.Form("txtOnHold").Item)	' Nullable: YES Type: datetime
		Else
			rs("OnHold") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtClosed").Item) Then
		If Not Request.Form("txtClosed").Item = "" Then
			rs("Closed") = SQLdatetimeADO(Request.Form("txtClosed").Item)	' Nullable: YES Type: datetime
		Else
			rs("Closed") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtCanceled").Item) Then
		If Not Request.Form("txtCanceled").Item = "" Then
			rs("Canceled") = SQLdatetimeADO(Request.Form("txtCanceled").Item)	' Nullable: YES Type: datetime
		Else
			rs("Canceled") = Null
		End If
	End If
	If Not IsEmpty(Request.Form("txtDenied").Item) Then
		If Not Request.Form("txtDenied").Item = "" Then
			rs("Denied") = SQLdatetimeADO(Request.Form("txtDenied").Item)	' Nullable: YES Type: datetime
		Else
			rs("Denied") = Null
		End If
	End If

	' CUSTOMIZED
	'--------------------------------------------------------------------------------------------------------------
	If Not IsEmpty(Request.Form("txtComments").Item) Then
		If Len(Trim(Request.Form("txtComments").Item)) > 0 Then
			rs("Comments") = Trim(Mid(Replace(Request.Form("txtComments").Item,Chr(13)+Chr(10),"%0D%0A"),1,2000))	' Nullable: YES Type: nvarchar
		Else
			rs("Comments") = Null
		End If
	End If
	'--------------------------------------------------------------------------------------------------------------

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
		keyvalue = rs("POPK")
	End If

    Set rs.ActiveConnection = Nothing
	rs.close
	Set rs = Nothing

End Sub

Sub db_li(rs,isinsert,suffix,htmltable,theindex,db)

	Dim fp
	fp = "txt" & LCase(htmltable) & suffix

	' -- Start Table Fields ------------------------------------------------

	rs("POPK") = keyvalue

    'Response.Write UCase(Trim(Mid(Request.Form(fp & "Part" & theindex).Item,1,25)))
    'Response.End

    If UCase(Trim(Mid(Request.Form(fp & "Part" & theindex).Item,1,25))) = "MANUAL ENTRY" Then

        If isinsert Then

            ' Ensure the Manual Entry Part Exists
            Call db.RunSQL("EXEC MC_CheckForManualEntryPart","")
            ' Get the Manual Entry  Part #
            Dim rsme,sql,partpk
            partpk = "-1"
            sql = "SELECT PartPK FROM Part WITH (NOLOCK) WHERE PartID = 'Manual Entry'"
	        'Response.Write sql
	        'Response.End
            Set rsme = db.RunSQLReturnRS(sql,"")
            If db.dok Then
                If Not rsme.eof Then
                    partpk = rsme("PartPK")
                End If
            End If
            rs("PartPK") = PartPK	' Nullable: No Type: int
            rs("PartID") = "Manual Entry"	' Nullable: No Type: nvarchar

            rs("LocationPK") = Null
            rs("LocationID") = Null
            rs("LocationName") = Null

    	    rs("DirectIssue") = True ' Nullable: No Type: bit

	        rsme.close
	        Set rsme = Nothing

        End If

        If Not IsEmpty(Request.Form(fp & "Part" & theindex & "DescH").Item) Then
	        rs("PartName") = Trim(Mid(Request.Form(fp & "Part" & theindex & "DescH").Item,1,500))	' Nullable: YES Type: nvarchar
        End If

    Else

	    If Not IsEmpty(Request.Form(fp & "Part" & theindex & "PK").Item) Then
		    If Len(Trim(Request.Form(fp & "Part" & theindex & "PK").Item)) > 0 Then
			    rs("PartPK") = Request.Form(fp & "Part" & theindex & "PK").Item	' Nullable: No Type: int
			    If Not IsEmpty(Request.Form(fp & "Part" & theindex).Item) Then
				    rs("PartID") = Trim(Mid(Request.Form(fp & "Part" & theindex).Item,1,25))	' Nullable: No Type: nvarchar
			    End If
			    If Not IsEmpty(Request.Form(fp & "Part" & theindex & "DescH").Item) Then
				    rs("PartName") = Trim(Mid(Request.Form(fp & "Part" & theindex & "DescH").Item,1,50))	' Nullable: YES Type: nvarchar
			    End If
		    End If
	    End If

	    If Not IsEmpty(Request.Form(fp & "DirectIssue" & theindex).Item) Then
		    rs("DirectIssue") = Request.Form(fp & "DirectIssue" & theindex).Item ' Nullable: No Type: bit
	    End If

        If Not IsEmpty(Request.Form(fp & "Location" & theindex & "PK").Item) Then
	        If Len(Trim(Request.Form(fp & "Location" & theindex & "PK").Item)) > 0 Then
		        rs("LocationPK") = Request.Form(fp & "Location" & theindex & "PK").Item	' Nullable: YES Type: int
		        If Not IsEmpty(Request.Form(fp & "Location" & theindex).Item) Then
			        rs("LocationID") = Trim(Mid(Request.Form(fp & "Location" & theindex).Item,1,25))	' Nullable: YES Type: nvarchar
		        End If
		        If Not IsEmpty(Request.Form(fp & "Location" & theindex & "DescH").Item) Then
			        rs("LocationName") = Trim(Mid(Request.Form(fp & "Location" & theindex & "DescH").Item,1,50))	' Nullable: YES Type: nvarchar
		        End If
	        End If
			'@$CUSTOMISED
			If Not IsEmpty(Request.Form(fp & "Bin" & theindex)) Then
				rs("Bin") = Trim(Mid(Request.Form(fp & "Bin" & theindex),1,25)) ' Nullable: YES Type: nvarchar
			End If
			'@$END
        End If

    End If

	If Len(Trim(Request.Form(fp & "LineItemNo" & theindex).Item)) > 0 Then
		rs("LineItemNo") = Request.Form(fp & "LineItemNo" & theindex).Item	' Nullable: No Type: int
	End If

   	If Not IsEmpty(Request.Form(fp & "VendorPartNumber" & theindex).Item) Then
		If Len(Request.Form(fp & "VendorPartNumber" & theindex).Item) > 0 Then
			rs("VendorPartNumber") = Trim(Mid(Request.Form(fp & "VendorPartNumber" & theindex).Item,1,50))	' Nullable: YES Type: nvarchar
		Else
			rs("VendorPartNumber") = Null
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "WO" & theindex & "PK").Item) Then
		If Not Request.Form(fp & "WO" & theindex & "PK").Item = "UNCHANGED" Then
			If Len(Trim(Request.Form(fp & "WO" & theindex & "PK").Item)) > 0 Then
				rs("WOPK") = Request.Form(fp & "WO" & theindex & "PK").Item	' Nullable: YES Type: int
				If Not IsEmpty(Request.Form(fp & "WO" & theindex).Item) Then
					rs("WOID") = Trim(Mid(Request.Form(fp & "WO" & theindex).Item,1,25))	' Nullable: YES Type: nvarchar
				End If
			Else
				rs("WOPK") = Null
				rs("WOID") = Null
			End If
		End If
	End If
	If Not IsEmpty(Request.Form(fp & "Asset" & theindex & "PK").Item) Then
		If Not Request.Form(fp & "Asset" & theindex & "PK").Item = "UNCHANGED" Then
			If Len(Trim(Request.Form(fp & "Asset" & theindex & "PK").Item)) > 0 Then
				rs("AssetPK") = Request.Form(fp & "Asset" & theindex & "PK").Item	' Nullable: YES Type: int
				If Not IsEmpty(Request.Form(fp & "Asset" & theindex).Item) Then
					rs("AssetID") = Trim(Mid(Request.Form(fp & "Asset" & theindex).Item,1,25))	' Nullable: YES Type: nvarchar
				End If
				If Not IsEmpty(Request.Form(fp & "Asset" & theindex & "DescH").Item) Then
					rs("AssetName") = Trim(Mid(Request.Form(fp & "Asset" & theindex & "DescH").Item,1,50))	' Nullable: YES Type: nvarchar
				End If
			Else
				rs("AssetPK") = Null
				rs("AssetID") = Null
				rs("AssetName") = Null
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
	If Not IsEmpty(Request.Form(fp & "SubAccount" & theindex & "PK").Item) Then
		If Not Request.Form(fp & "SubAccount" & theindex & "PK").Item = "UNCHANGED" Then
			If Len(Trim(Request.Form(fp & "SubAccount" & theindex & "PK").Item)) > 0 Then
				rs("SubAccountPK") = Request.Form(fp & "SubAccount" & theindex & "PK").Item	' Nullable: YES Type: int
				If Not IsEmpty(Request.Form(fp & "SubAccount" & theindex).Item) Then
					rs("SubAccountID") = Trim(Mid(Request.Form(fp & "SubAccount" & theindex).Item,1,25))	' Nullable: YES Type: nvarchar
				End If
				If Not IsEmpty(Request.Form(fp & "SubAccount" & theindex & "DescH").Item) Then
					rs("SubAccountName") = Trim(Mid(Request.Form(fp & "SubAccount" & theindex & "DescH").Item,1,50))	' Nullable: YES Type: nvarchar
				End If
			Else
				rs("SubAccountPK") = Null
				rs("SubAccountID") = Null
				rs("SubAccountName") = Null
			End If
		End If
	End If
	rs("OrderUnits") = Trim(Mid(Request.Form(fp & "OrderUnits" & theindex).Item,1,25))	' Nullable: No Type: nvarchar
	If Not IsEmpty(Request.Form(fp & "OrderUnits" & theindex & "DescH").Item) Then
		rs("OrderUnitsDesc") = Trim(Mid(Request.Form(fp & "OrderUnits" & theindex & "DescH").Item,1,50))	' Nullable: YES Type: nvarchar
	End If
	If Len(Trim(Request.Form(fp & "ConversionToIssueUnits" & theindex).Item)) > 0 Then
		rs("ConversionToIssueUnits") = FixInternationalNumber(Request.Form(fp & "ConversionToIssueUnits" & theindex).Item)	' Nullable: No Type: real
	End If
	If Not IsEmpty(Request.Form(fp & "IssueUnits" & theindex).Item) Then
		If Len(Request.Form(fp & "IssueUnits" & theindex).Item) > 0 Then
			rs("IssueUnits") = Trim(Mid(Request.Form(fp & "IssueUnits" & theindex).Item,1,25))	' Nullable: YES Type: nvarchar
			If Not IsEmpty(Request.Form(fp & "IssueUnits" & theindex & "DescH").Item) Then
				rs("IssueUnitsDesc") = Trim(Mid(Request.Form(fp & "IssueUnits" & theindex & "DescH").Item,1,50))	' Nullable: YES Type: nvarchar
			End If
		Else
			rs("IssueUnits") = Null
			rs("IssueUnitsDesc") = Null
		End If
	End If
	If Len(Trim(Request.Form(fp & "OrderUnitPrice" & theindex).Item)) > 0 Then
		rs("OrderUnitPrice") = FixInternationalNumber(Request.Form(fp & "OrderUnitPrice" & theindex).Item)	' Nullable: No Type: money
	End If
	If Len(Trim(Request.Form(fp & "OrderUnitQty" & theindex).Item)) > 0 Then
		rs("OrderUnitQty") = Request.Form(fp & "OrderUnitQty" & theindex).Item	' Nullable: No Type: int
	End If
	If Len(Trim(Request.Form(fp & "OrderUnitQtyBackOrdered" & theindex).Item)) > 0 Then
		rs("OrderUnitQtyBackOrdered") = Request.Form(fp & "OrderUnitQtyBackOrdered" & theindex).Item	' Nullable: No Type: int
	End If
	If Len(Trim(Request.Form(fp & "OrderUnitQtyReceived" & theindex).Item)) > 0 Then
		rs("OrderUnitQtyReceived") = Request.Form(fp & "OrderUnitQtyReceived" & theindex).Item	' Nullable: No Type: int
	End If
	If IsDate(Request.Form(fp & "ReceiveDate" & theindex)) Then
		If Not IsEmpty(Request.Form(fp & "ReceiveTime" & theindex)) Then
			rs("ReceiveDate") = SQLdatetimeADOAT(Request.Form(fp & "ReceiveDate" & theindex) & " " & Request.Form(fp & "ReceiveTime" & theindex))	' Nullable: No Type: datetime
		Else
			rs("ReceiveDate") = SQLdatetimeADO(Request.Form(fp & "ReceiveDate" & theindex) & " " & Time())	' Nullable: No Type: datetime
		End If
	Else
		' Default to Now if at least 1 part has been received
		If rs("OrderUnitQtyReceived") > 0 Then
			rs("ReceiveDate") = SQLdatetimeADO(DateTimeNullCheck(Now()))	' Nullable: Yes Type: datetime
		Else
			rs("ReceiveDate") = Null
		End If
	End If
	If Len(Trim(Request.Form(fp & "Discount" & theindex).Item)) > 0 Then
		rs("Discount") = FixInternationalNumber(Request.Form(fp & "Discount" & theindex).Item)	' Nullable: No Type: float
	End If
	If Len(Trim(Request.Form(fp & "Subtotal" & theindex).Item)) > 0 Then
		rs("Subtotal") = FixInternationalNumber(Request.Form(fp & "Subtotal" & theindex).Item)	' Nullable: No Type: money
	End If
	' --------------- Customized ------------------
	rs("IsTax") = (InStr(Request.Form(fp & "IsTax" & theindex).Item,"taskchecked.gif") > 0) ' Nullable: No Type: bit
	' ---------------------------------------------
	If Len(Trim(Request.Form(fp & "TaxRate" & theindex).Item)) > 0 Then
		rs("TaxRate") = FixInternationalNumber(Request.Form(fp & "TaxRate" & theindex).Item)	' Nullable: No Type: real
	End If
	If Len(Trim(Request.Form(fp & "TaxAmount" & theindex).Item)) > 0 Then
		rs("TaxAmount") = FixInternationalNumber(Request.Form(fp & "TaxAmount" & theindex).Item)	' Nullable: No Type: real
	End If
	If Len(Trim(Request.Form(fp & "LineItemTotal" & theindex).Item)) > 0 Then
		rs("LineItemTotal") = FixInternationalNumber(Request.Form(fp & "LineItemTotal" & theindex).Item)	' Nullable: No Type: money
	End If
	If Not IsEmpty(Request.Form(fp & "Comments" & theindex).Item) Then
		If Len(Request.Form(fp & "Comments" & theindex).Item) > 0 Then
			rs("Comments") = Trim(Mid(Request.Form(fp & "Comments" & theindex).Item,1,2000))	' Nullable: YES Type: nvarchar
		Else
			rs("Comments") = Null
		End If
	End If
	If IsDate(Request.Form(fp & "DueDate" & theindex)) Then
		If Not IsEmpty(Request.Form(fp & "ReceiveTime" & theindex)) Then
			rs("DueDate") = SQLdatetimeADOAT(Request.Form(fp & "DueDate" & theindex) & " " & Request.Form(fp & "ReceiveTime" & theindex))	' Nullable: No Type: datetime
		Else
			rs("DueDate") = SQLdatetimeADO(Request.Form(fp & "DueDate" & theindex) & " " & Time())	' Nullable: No Type: datetime
		End If
	Else
		' Default to Now if at least 1 part has been received
		If rs("OrderUnitQtyReceived") > 0 Then
			rs("DueDate") = SQLdatetimeADO(DateTimeNullCheck(Now()))	' Nullable: Yes Type: datetime
		Else
			rs("DueDate") = Null
		End If
	End If

	' -- End Table Fields ------------------------------------------------

	Call db_version(rs)
End Sub

Sub db_at(rs,isinsert,suffix,htmltable,theindex)

	Dim fp
	fp = "txt" & LCase(htmltable) & suffix

	' -- Start Table Fields ------------------------------------------------

	rs("POPK") = keyvalue

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

	rs("POPK") = keyvalue

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



