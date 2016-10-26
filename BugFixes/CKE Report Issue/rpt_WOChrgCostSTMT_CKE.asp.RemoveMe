<%@ EnableSessionState=False Language=VBScript %>
<% Option Explicit %>
<!--#INCLUDE FILE="../common/mc_all.asp" -->
<!--#INCLUDE FILE="includes/mcReport_common.asp" -->
<!--#INCLUDE FILE="includes/rpt_WO_common.asp" -->
<%
'Response.Write("URL QueryString: <br>" & Request.QueryString)

Dim WOPK, WOsql, WOGroupPK
Dim RS_WO, RS_WOassign, RS_Asset, RS_WOtask, RS_WOpart, RS_WOmiscCost, RS_WOTool, RS_WODocument, RS_WOPref
Dim AssetName, IconFile, NoIconFile_Location, NoIconFile_Asset, skipfirstrow, AssetOutput
Dim LaborTotal, MaterialsTotal, OtherCostsTotal, isCharge, WOStatus
isCharge = false

'Buffer code begin
Dim BufferCount,BufferCountB
'Buffer end
	
'Call AspDebug()
'Response.End

Call SetPrintedFlag()
Call SetupWOBarcode()
Call SetupWOGroupData()
If InWoModule Then
	Call GetRecordStatusForClickedWO()
End If
Call DoOutput()

If SubReport or FromAgent or (Trim(UCase(Request("EmailReport"))) = "Y") or (Not Request.QueryString("ExportReportOutputType") = "") Then
' ----- No Script -----
Else
'BEGIN Buffer code - End of report output
Response.Write "<script type='text/javascript'>try{ShowHideLoading('0');} catch(e){} </script> "
'End buffer
End If

Call CloseDown()

Sub DoOutput()

    Dim LaborRS, PartRS, ToolRS, OtherCostRS, DocumentRS, Field, IsGroup
    
	Call OutputHeader()
	Call OutputStandardBodyTag()		
	Call OutputEmailMessage()

	If len(errormessage) or len(uf_errormessage) Then
	%>
		<font face="Arial" size="2" color="red"><% =uf_errormessage %></font><br>
	<%
	Else
		Call OutputToolbar()

		' Do not rw the form tag
		If Not FromAgent Then
			Response.Write("<form id=""mcform"" name=""mcform"" method=""post"">")			
		End If

        If SubReport or FromAgent or (Trim(UCase(Request("EmailReport"))) = "Y") or (Not Request.QueryString("ExportReportOutputType") = "") Then
    	' ----- No Script -----
        Else
        'BEGIN Buffer code
        Response.Write "<script type='text/javascript'> "
        Response.Write "function ShowHideLoading(inShow){ "
        Response.Write "if (inShow=='1'){try{document.getElementById('loadingArea').style.display='';} catch(e){}} "
        Response.Write "else{try{document.getElementById('loadingArea').style.display='none';} catch(e){}} "
        Response.Write "} "
        Response.Write "</script> "
        'Div tag above Header code
        Response.Write "<div id='loadingArea' name='loadingArea' style='position:absolute; width:100%; height:100%; overflow:hidden; background-color:#ffffff; display:none; z-index:1000;'><div style='text-align:center;position:relative; top:25%;'><img src='logo.gif' alt='' title='' style='border:none;' /><br /><br /><br /><img src='progress.gif' alt='' title='' style='border:none;' /></div></div> "
        'End buffer
        End If
        		
		If Not RS_WO.EOF Then

            'Begin Buffer Code: 
            Response.Flush
            BufferCount = 0
            If SubReport or FromAgent or (Trim(UCase(Request("EmailReport"))) = "Y") or (Not Request.QueryString("ExportReportOutputType") = "") Then
    	    ' ----- No Script -----
            Else
            Response.Write "<script type='text/javascript'>try{ShowHideLoading('1');} catch(e){} </script> "
            End If
            'End buffer

		'loop through all WO
		Do While Not RS_WO.EOF
		    isCharge = false
            If ( (Not RS_WO("IsOpen")) Or (NullCheck(RS_WO("Canceled")) <> "") OR (NullCheck(RS_WO("Denied")) <> "") ) Then
                WOStatus = "CLOSED"
            Else
                WOStatus = "OPEN"
            End If
            'Response.write "WOStatus: " & WOStatus & "<br>"
            If InStr(RS_WO("RepairCenterID"),"R99") > 0 Then
                isCharge = true
            End If
            'Begin Buffer Code: Buffer flush before and during report output loop
            BufferCount = BufferCount + 1
            If BufferCount > 20 Then
                BufferCount = 0
                Response.Flush
            End If
            'End buffer

			'set work order PK
			WOPK = RS_WO("WOPK")
			WOGroupPK = RS_WO("WOGroupPK")
						
			rw "<table border=""0"" width=""100%"">"			
				rw "<tr>"
					rw "<td valign=""top"">"
						Call OutputLogoOrName()
					rw "</td>"
					rw "<td class=""no-print"" align=""center"" valign=""bottom""><nobr><span style=""font-family:Arial;font-size:20px;color:#333333;font-weight:bold"">"
                    If isCharge Then
                        rw "Charge Statement"
                    Else
                        rw "Cost Statement"
                    End If
                    rw "</span></nobr><br></td>"					
					rw "<td valign=""top"" align=""right"">"
						If NullCheck(RS_WO("WOGroupPK")) = "" or _
						   NullCheck(RS_WO("WOGroupPK")) = "-1" Then
								Call OutputWOHeaderRight("WO")
						 Else 
								Call OutputWOHeaderRight("WOGROUP")
						End If
					rw "</td>"
				rw "</tr>"
			rw "</table>"
									
			' ====================================================================
			' INDIVIDUAL WORK ORDER
			' ====================================================================
			'If NullCheck(RS_WO("WOGroupPK")) = "" or _
			'   NullCheck(RS_WO("WOGroupPK")) = "-1" Then

			'	IsGroup = False
									
				' Output Main Details, Tasks, Labor, Materials, Tools, Other Costs,
				' & Documents for each Work Order
				' ================================================================
				If reporthasfields Then
					Call OutputReportBox()
				End If
				Call OutputMainDetailsBox()
				Call NewOutputTaskBox(RS_WOTask)
				Call NewOutputLaborBox(RS_WOAssign,False)
				Call NewOutputMaterialsToolsBox(RS_WOPart,RS_WOTool,False)
				Call NewOutputOtherCostsBox(RS_WOMiscCost,False)
				Call NewOutputGrandTotalBox(False)
				Call OutputDocumentsBoxCustom(RS_WODocument,False)			
				If Not reporthasfields Then
					Call OutputReportBox()
				End If
			

						
			RS_WO.MoveNext
			
			' do not output a page break on the last record
			If Not RS_WO.EOF Then
				rw "<P style='page-break-before: always'>"
			End If			
		Loop
		Else
			rw "<div style=""padding-top:5px; padding-left:5px; font-family:arial; font-size:10pt; color:gray; font-weight:bold;"">"
			rw "(No Records)"
			rw "</div>"
		End If
	End If

	Call OutputSQL()
	Call OutputFormHelper()	
	
	If wostate = "CC" Then
		Call DoMCDivs()
		Call DoIFrame
	End If

	Call OutputFooter()
	Call EndFile()

	CloseObj RS_WO
	CloseObj RS_WOassign
	CloseObj RS_Asset
	CloseObj RS_WOtask
	CloseObj RS_WOpart
	CloseObj RS_WOmiscCost
	CloseObj RS_WOTool
	CloseObj RS_WODocument

End Sub

Sub SetupWOGroupData()

	Dim RecordType
	
	
	
	If errormessage = "" Then
		' Set RecordType	
        'If WOStatus = "OPEN" Then
        '    RecordType = "1"
        'Else
            RecordType = "2"		
		'End If

		If InStr(UCase(sql_where),"LEFT OUTER JOIN ") > 0 or InStr(UCase(sql_where),"INNER JOIN ") > 0 Then
			WOsql = "SELECT DISTINCT WO.WOPK FROM WO WITH (NOLOCK) LEFT OUTER JOIN Asset WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON AssetHierarchy.AssetPK = WO.AssetPK " & sql_where
		Else
			WOsql = "SELECT WOPK FROM WO WITH (NOLOCK) LEFT OUTER JOIN Asset WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON AssetHierarchy.AssetPK = WO.AssetPK " & sql_where
		End If
		'Response.Write WOsql
		'Response.End
		
		' GET WORK ORDER
		If InStr(UCase(sql_where),"LEFT OUTER JOIN ") > 0 or InStr(UCase(sql_where),"INNER JOIN ") > 0 Then
			sql = "SELECT DISTINCT WO.*, Address = dbo.MCUDF_GetAssetAddressWalkingUpTree(Asset.AssetPK), Asset.LeaseNumber, Asset.IsLocation, Asset.Meter1UnitsDesc, Asset.Meter2UnitsDesc, InstructionsAsset = Case When Asset.InstructionsToWO = 1 Then Asset.Instructions Else Null End FROM WO WITH (NOLOCK) LEFT OUTER JOIN Asset WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " & sql_where & " ORDER BY WO.WOGroupPK Desc, WO.WOPK"
		Else
			sql = "SELECT WO.*,Address = dbo.MCUDF_GetAssetAddressWalkingUpTree(Asset.AssetPK), Asset.LeaseNumber, Asset.IsLocation, Asset.Meter1UnitsDesc, Asset.Meter2UnitsDesc, InstructionsAsset = Case When Asset.InstructionsToWO = 1 Then Asset.Instructions Else Null End FROM WO WITH (NOLOCK) LEFT OUTER JOIN Asset WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " & sql_where & " ORDER BY WO.WOGroupPK Desc, WO.WOPK"
		End If
		'Response.Write sql
		'Response.End
		
		Set RS_WO = db.runSQLReturnRS(sql,"")
		If Trim(Request.QueryString("sqlwhere")) = "" Then
			Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	
		Else
			Call dok_check_afterflush_noinfo(db,"Report Message","You have chosen to run a report that is not compatible with the current module criteria. You can either choose a different report, or you can run the selected report from the Maintenance Reporter applicaiton by clicking the Reports button on the toolbar.")		
		End If


        ' GET ACTUAL LABOR
		sql = "SELECT   WO.ChargeLaborActual, WO.CostLaborActual, WOlabor.PK, WOlabor.WOPK, WOlabor.LaborPK, lt.moduleid, WOlabor.LaborID, WOlabor.LaborName, WOlabor.EstimatedHours, WOlabor.RegularHours, WOlabor.OvertimeHours, WOlabor.OtherHours, WOlabor.WorkDate, WOlabor.TimeIn, " & _
		"			  WOlabor.TimeOut, WOlabor.AccountID, WOlabor.AccountName, WOlabor.CategoryID, WOlabor.CategoryName, WOlabor.TotalCost, " & _ 
		"			  WOlabor.TotalCharge, WOlabor.CostRegular, WOlabor.CostOvertime, WOlabor.CostOther, WOlabor.ChargeRate, WOlabor.ChargePercentage, WOlabor.RowVersionDate, Labor.Photo, wolabor.comments " & _
		"FROM WOlabor WITH (NOLOCK) INNER JOIN WO WITH (NOLOCK) ON WO.WOPK = WOlabor.WOPK LEFT OUTER JOIN " & _
		"			  Labor WITH (NOLOCK) ON WOlabor.LaborPK = Labor.LaborPK INNER JOIN " & _
		"           LaborTypes lt WITH (NOLOCK) ON lt.LaborType = Labor.LaborType " & _
		"WHERE (WOlabor.RecordType = " & RecordType & ") "
		If Not sql_where = "" Then
		sql = sql & "AND (WOlabor.WOPK in (" & WOsql & ")) "
		End If
		sql = sql & _
		"ORDER BY WO.WOGroupPK Desc, WOlabor.WOPK, Labor.LaborName"				  		

		Set RS_WOAssign = db.runSQLReturnRS(sql,"")
		Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

		' Get Work Order Asset Heirarchy
		sql = "SELECT WO.WOPK, Asset.Icon, Asset.AssetPK, Asset.AssetID, Asset.AssetName, Asset.IsLocation, Asset.IsUp " +_
			  "FROM AssetAncestor WITH (NOLOCK) INNER JOIN " +_
	          "Asset WITH (NOLOCK) ON AssetAncestor.AncestorPK = Asset.AssetPK INNER JOIN " +_
			  "WO WITH (NOLOCK) ON AssetAncestor.AssetPK = WO.AssetPK " +_
			  "WHERE     (AssetAncestor.System = N'MC') "
			  If Not sql_where = "" Then
				sql = sql & "AND (WO.WOPK IN (" & WOsql & ")) "
			  End If
			  sql = sql & _
			  "ORDER BY WO.WOGroupPK Desc, WO.WOPK, AssetAncestor.AncestorLevel"

		'Response.Write sql
		'Response.End
		
		Set RS_Asset = db.runSQLReturnRS(sql,"")
		Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

		'Response.Write db.derror
		'Response.End
				
		' Get Work Order Tasks
		sql = "SELECT WOtask.PK, WOtask.WOPK, WOtask.TaskNo, TaskAction = CASE WHEN WOTask.AssetPK Is Not Null Then '<b>'+Asset.AssetName + ' [' + Asset.AssetID + ']</b> ' + WOtask.TaskAction Else WOtask.TaskAction END, WOtask.Rate, WOtask.Measurement, WOtask.Initials, WOtask.Fail, WOtask.Complete, WOtask.Header, WOtask.LineStyle, WOtask.Comments " +_
			  "FROM WOtask WITH (NOLOCK) INNER JOIN WO WITH (NOLOCK) ON WO.WOPK = WOtask.WOPK LEFT OUTER JOIN Asset WITH (NOLOCK) ON Asset.AssetPK = WOTask.AssetPK "
			  If Not sql_where = "" Then
				sql = sql & "WHERE (WOtask.WOPK in (" & WOsql & ")) "
			  End If
			  sql = sql & _			  
			  "ORDER BY WO.WOGroupPK Desc, WOtask.WOPK, WOtask.TaskNo"

		Set RS_WOtask = db.runSQLReturnRS(sql,"")
		Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	
		
		' Get Work Order Material
		sql = "SELECT WO.ChargePartActual, WO.CostPartActual, WOpart.PK, WOpart.WOPK, WOpart.PartID, WOpart.PartName, WOpart.LocationID, WOpart.QuantityEstimated, WOpart.QuantityActual, WOpart.OtherCost, WOPart.TotalCharge, WOPart.TotalCost, WOpart.AccountID, Part.PartDescription " +_
			  "FROM WOpart WITH (NOLOCK) LEFT OUTER JOIN " +_
			  "Part WITH (NOLOCK) ON WOpart.PartPK = Part.PartPK  INNER JOIN WO WITH (NOLOCK) ON WO.WOPK = WOpart.WOPK " +_
			  "WHERE (WOpart.RecordType = " & RecordType & ") "
			  If Not sql_where = "" Then
				sql = sql & "AND (WOpart.WOPK in (" & WOsql & ")) "
			  End If
			  sql = sql & _			  			  
			  "ORDER BY WO.WOGroupPK Desc, WOpart.WOPK, WOpart.LocationID, WOpart.PartID"

		Set RS_WOpart = db.runSQLReturnRS(sql,"")
		Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

		' Get Work Order Other Costs
		sql = "SELECT WO.ChargeMiscActual, WO.CostMiscActual, WOmiscCost.PK, WOmiscCost.WOPK, WOmiscCost.MiscCostName, WOmiscCost.MiscCostDesc, WOmiscCost.InvoiceNumber, WomiscCost.MiscCostDate, WOmiscCost.AccountID, WOmiscCost.AccountName, WOmiscCost.QuantityEstimated, WOmiscCost.EstimatedCost, WOmiscCost.ActualCost, WOmiscCost.TotalCharge " +_
			  "FROM WOmiscCost WITH (NOLOCK) INNER JOIN WO WITH (NOLOCK) ON WO.WOPK = WOmiscCost.WOPK " +_
			  "WHERE (WOmiscCost.RecordType = " & RecordType & ") "
			  If Not sql_where = "" Then
				sql = sql & "AND (WOmiscCost.WOPK in (" & WOsql & ")) "
			  End If
			  sql = sql & _			  			  			  
			  "ORDER BY WO.WOGroupPK Desc, WOmiscCost.WOPK, WOmiscCost.MiscCostName"
		Set RS_WOmiscCost = db.runSQLReturnRS(sql,"")
		Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

		' Get Work Order Estimated Tools
		sql = "SELECT WOtool.WOPK, WOtool.ToolID, WOtool.ToolName, WOtool.LocationID, WOtool.LocationName, WOtool.QuantityEstimated " +_
			  "FROM WOtool WITH (NOLOCK) LEFT OUTER JOIN " +_
	          "Tool WITH (NOLOCK) ON WOtool.ToolPK = Tool.ToolPK INNER JOIN WO WITH (NOLOCK) ON WO.WOPK = WOtool.WOPK " +_
			  "WHERE (WOtool.RecordType = " & RecordType & ") "
			  If Not sql_where = "" Then
				sql = sql & "AND (WOtool.WOPK in (" & WOsql & ")) "
			  End If
			  sql = sql & _			  			  			  			  
			  "ORDER BY WO.WOGroupPK Desc, WOtool.WOPK, WOtool.LocationName, WOtool.ToolName"

		Set RS_WOtool = db.runSQLReturnRS(sql,"")
		Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

		' Get Work Order Document Attachments */
		
		' NOTE: Since we are including a text field (DocumentText) we must do a 
		' UNION ALL - since not doing a UNION ALL does an implicit DISTINCT which
		' is ILLEGAL with TEXT FIELDS....
		
		sql = _
		"SELECT     WO.WOGroupPK, WO.WOPK, md.PK, d.LocationType, d.DocumentID, d.DocumentName, md.ModuleID, d.DocumentTypeDesc, " & _
		"                      d.Location, md.PrintWithWO, md.SendWithEmail, md.RowVersionDate, d.Photo, " & _
		"                      MCModule.TitleforDocumentList, d.DocumentText " & _
		"FROM         AssetDocument md WITH (NOLOCK) LEFT OUTER JOIN " & _
		"                      Document d WITH (NOLOCK) ON md.DocumentPK = d.DocumentPK INNER JOIN " & _
		"                      MCModule WITH (NOLOCK) ON md.ModuleID = MCModule.ModuleID INNER JOIN " & _
		"                      WO WITH (NOLOCK) ON WO.AssetPK = md.AssetPK " + _
		"WHERE "
		If Not sql_where = "" Then
		  sql = sql & "(WO.WOPK in (" & WOsql & ")) AND "
		End If
		sql = sql & _			  			  			  			  		
		"(d.Active = 1) " & _
		"UNION ALL " & _
		"SELECT     WO.WOGroupPK, WO.WOPK, md.PK, d.LocationType, d.DocumentID, d.DocumentName, md.ModuleID, d.DocumentTypeDesc, " & _
		"                      d.Location, md.PrintWithWO, md.SendWithEmail, md.RowVersionDate, d.Photo, " & _
		"                      MCModule.TitleforDocumentList, d.DocumentText " & _
		"FROM         LaborDocument md WITH (NOLOCK) LEFT OUTER JOIN " & _
		"                      Document d WITH (NOLOCK) ON md.DocumentPK = d.DocumentPK INNER JOIN " & _
		"                      MCModule WITH (NOLOCK) ON md.ModuleID = MCModule.ModuleID INNER JOIN " & _
		"                      WO WITH (NOLOCK) ON WO.RequesterPK = md.LaborPK " & _
		"WHERE "
		If Not sql_where = "" Then
		  sql = sql & "(WO.WOPK in (" & WOsql & ")) AND "
		End If
		sql = sql & _			  			  			  			  		
		"(d.Active = 1) " & _
		"UNION ALL " & _		
		"SELECT     WO.WOGroupPK, WO.WOPK, md.PK, d.LocationType, d.DocumentID, d.DocumentName, md.ModuleID, d.DocumentTypeDesc, " & _ 
		"                      d.Location, md.PrintWithWO, md.SendWithEmail, md.RowVersionDate, d.Photo, " & _
		"                      MCModule.TitleforDocumentList, d.DocumentText " & _
		"FROM         RepairCenterDocument md WITH (NOLOCK) LEFT OUTER JOIN " & _
		"                      Document d WITH (NOLOCK) ON md.DocumentPK = d.DocumentPK INNER JOIN " & _
		"                      MCModule WITH (NOLOCK) ON md.ModuleID = MCModule.ModuleID INNER JOIN " & _
		"                      WO WITH (NOLOCK) ON WO.RepairCenterPK = md.RepairCenterPK " & _
		"WHERE "
		If Not sql_where = "" Then
		  sql = sql & "(WO.WOPK in (" & WOsql & ")) AND "
		End If
		sql = sql & _			  			  			  			  		
		"(d.Active = 1) " & _
		"UNION ALL " & _
		"SELECT     WO.WOGroupPK, WO.WOPK, md.PK, d.LocationType, d.DocumentID, d.DocumentName, md.ModuleID, d.DocumentTypeDesc, " & _ 
		"					   d.Location, md.PrintWithWO, md.SendWithEmail, md.RowVersionDate, d.Photo, " & _
		"                      MCModule.TitleforDocumentList, d.DocumentText " & _
		"FROM ProjectDocument md WITH (NOLOCK) LEFT OUTER JOIN " & _
		"                      Document d WITH (NOLOCK) ON md.DocumentPK = d.DocumentPK INNER JOIN " & _
		"                      MCModule WITH (NOLOCK) ON md.ModuleID = MCModule.ModuleID INNER JOIN " & _
		"                      WO WITH (NOLOCK) ON WO.ProjectPK = md.ProjectPK " & _
		"WHERE "
		If Not sql_where = "" Then
		  sql = sql & "(WO.WOPK in (" & WOsql & ")) AND "
		End If
		sql = sql & _			  			  			  			  		
		"(d.Active = 1) " & _
		"UNION ALL " & _
		"SELECT     WO.WOGroupPK, WO.WOPK, md.PK, d.LocationType, d.DocumentID, d.DocumentName, md.ModuleID, d.DocumentTypeDesc, " & _
		"                      d.Location, md.PrintWithWO, md.SendWithEmail, md.RowVersionDate, d.Photo, " & _
		"                      MCModule.TitleforDocumentList, d.DocumentText " & _
		"FROM         WOdocument md WITH (NOLOCK) LEFT OUTER JOIN " & _
		"                      Document d WITH (NOLOCK) ON md.DocumentPK = d.DocumentPK INNER JOIN " & _
		"                      MCModule WITH (NOLOCK) ON md.ModuleID = MCModule.ModuleID INNER JOIN " & _
		"                      WO WITH (NOLOCK) ON WO.WOPK = md.WOPK " & _
		"WHERE "
		If Not sql_where = "" Then
		  sql = sql & "(WO.WOPK in (" & WOsql & ")) AND "
		End If
		sql = sql & _			  			  			  			  		
		"(d.Active = 1) " & _
		"ORDER BY WO.WOGroupPK Desc, WO.WOPK, md.ModuleID, d.DocumentID "
		
		'Response.Write "<TextArea rows=30 cols=100>"
		'Response.Write sql
		'Response.Write "</TextArea>"
		'Response.End
				
		Set RS_WOdocument = db.runSQLReturnRS(sql,"")				
		Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

		Set RS_WOPref = db.runSPReturnRS("MC_GetWorkOrderPrefs",Array(Array("@LaborPK", adInteger, adParamInput, 4, GetSession("USERPK")),Array("@RepairCenterPK", adInteger, adParamInput, 4, GetSession("RCPK"))),"")
		Call dok_check_afterflush(db,"Report Message","There was a problem retrieving the Work Order Preferences. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

		If Not RS_WOPref.Eof Then
			WO_LABORSECTION = RS_WOPref("WO_LABORSECTION")
			WO_MATERIALSECTION = RS_WOPref("WO_MATERIALSECTION")
			WO_OTHERCOSTSECTION = RS_WOPref("WO_OTHERCOSTSECTION")
			WO_DOCUMENTSECTION = RS_WOPref("WO_DOCUMENTSECTION")
			WO_REPORTSECTION = RS_WOPref("WO_REPORTSECTION")
			WO_LABORSECTION_B = RS_WOPref("WO_LABORSECTION_B")
			WO_MATERIALSECTION_B = RS_WOPref("WO_MATERIALSECTION_B")
			WO_OTHERCOSTSECTION_B = RS_WOPref("WO_OTHERCOSTSECTION_B")
			WO_TASKSECTION = RS_WOPref("WO_TASKSECTION")
			WO_TASKSECTION_B = RS_WOPref("WO_TASKSECTION_B")
		End If

		CloseObj RS_WOPref

	End If

End Sub

Sub NewOutputOtherCostsHeader()
	rw "<tr>"
		rw "<td class=""labels"">Name</td>"
		rw "<td class=""labels"">Description</td>"
		'rw "<td class=""labels"" width=""100"">Invoice&nbsp;#</td>"
		'rw "<td class=""labels"">Date</td>"
		rw "<td class=""labels"" width=""80"" align=""right"">Charge</td>"
	rw "</tr>"
End Sub

Sub NewOutputOtherCosts(RS_WOmiscCost)
	rw "<tr>"
		If CInt(NullCheck(RS_WOmiscCost("QuantityEstimated"))) > 1 Then
			rw "<td class=""data_underline"">" & NullCheckNBSP(RS_WOmiscCost("MiscCostName")) & "&nbsp;(" & CInt(NullCheck(RS_WOmiscCost("QuantityEstimated"))) & ")</td>"
		Else
			rw "<td class=""data_underline"">" & NullCheckNBSP(RS_WOmiscCost("MiscCostName")) & "&nbsp;</td>"
		End If
		rw "<td class=""data_underline"">" & NullCheckNBSP(RS_WOmiscCost("MiscCostDesc")) & "</td>"
		'rw "<td class=""data_underline"">" & NullCheckNBSP(RS_WOmiscCost("InvoiceNumber")) & "</td>"
		'rw "<td class=""data_underline"">" & NullCheckNBSP(DateNullCheck(RS_WOmiscCost("MiscCostDate"))) & "</td>"
        If isCharge Then
		    rw "<td class=""data_underline"" align=""right"">" & FormatCurrency(NullCheck(RS_WOmiscCost("TotalCharge"))) & "</td>"
        Else
		    rw "<td class=""data_underline"" align=""right"">" & FormatCurrency(NullCheck(RS_WOmiscCost("ActualCost"))) & "</td>"
        End If
	rw "</tr>"
End Sub

Sub NewOutputOtherCostsTotals(total)
	OtherCostsTotal = Total
	rw "<tr>"
		rw "<td class=""labels"" colspan=""2"">Other Cost Totals:</td>"
		'rw "<td class=""data_underline""></td>"
		'rw "<td class=""data_underline""></td>"
		'rw "<td class=""data_underline""></td>"
		rw "<td class=""data_underline"" align=""right"">" & FormatCurrency(OtherCostsTotal) & "</td>"
	rw "</tr>"
End Sub

Sub NewOutputOtherCostsBox(rs,nowocheck)
	Dim BlankRowNum, GrandTotal
	GrandTotal = 0
	rw "<fieldset style=""padding-top:14px"">"
		If nowocheck Then
			rw "<legend class=""legendHeader"">Other Costs (Summary)</legend>"
		Else
			rw "<legend class=""legendHeader"">Other Costs</legend>"
		End If
		rw "<table style=""margin-top:5px;"" border=""0"" cellspacing=""3"" cellpadding=""0"" width=""98%"" align=""center"">"

			Call NewOutputOtherCostsHeader()
					
			If Not rs.eof and (NullCheck(rs("WOPK")) = NullCheck(WOPK) or nowocheck) Then
				Do While Not rs.eof and (NullCheck(rs("WOPK")) = NullCheck(WOPK) or nowocheck)
					Call NewOutputOtherCosts(rs)
                    If ischarge Then
					    GrandTotal = GrandTotal + rs("TotalCharge")
                    Else
					    GrandTotal = GrandTotal + rs("ActualCost")
                    End If
					rs.MoveNext
				Loop
			Else
				Call NewOutputOtherCostBlankRow()
			End If
			NewOutputOtherCostsTotals(GrandTotal)

		rw "</table><br>"
	rw "</fieldset>"
End Sub

Sub NewOutputLaborHeader(LaborFormat)
	rw "<tr>"	
		rw "<td class=""labels"">Labor</td>"
		rw "<td class=""labels"" align=""center"" width=""60"" nowrap>&nbsp;Reg&nbsp;Hrs&nbsp;</td>"
		rw "<td class=""labels"" align=""center"" width=""60"" nowrap>&nbsp;OT&nbsp;Hrs&nbsp;</td>"
		rw "<td class=""labels"" align=""center"" width=""60"" nowrap>&nbsp;Other&nbsp;Hrs&nbsp;</td>"
		rw "<td class=""labels"" align=""left"" width=""100"" nowrap>Date</td>"
		rw "<td class=""labels"" align=""right"" width=""80"" nowrap>Charge</td>"
	rw "</tr>"
End Sub

Sub NewOutputLaborBox(rs,nowocheck)

	Dim BlankRowNum,LaborFormat,LaborCols, GrandTotal
	BlankRowNum = 0	
	GrandTotal = 0

	LaborFormat = "NONE"
	
	rw "<fieldset style=""padding-top:14px"">"
		If nowocheck Then
			rw "<legend class=""legendHeader"">Labor (Summary)</legend>"
		Else
			rw "<legend class=""legendHeader"">Labor</legend>"
		End If
		
		rw "<table style=""margin-top:5px;"" border=""0"" cellspacing=""3"" cellpadding=""0"" width=""98%"" align=""center"">"

			Call NewOutputLaborHeader(LaborFormat)
						
			' labor data 	
			If Not rs.eof and (NullCheck(rs("WOPK")) = NullCheck(WOPK) or nowocheck) Then
				Do While Not rs.eof and (NullCheck(rs("WOPK")) = NullCheck(WOPK) or nowocheck)
					Call NewOutputLabor(LaborFormat,rs)
					If ischarge Then
                        GrandTotal = GrandTotal + rs("TotalCharge")
                    Else
                        GrandTotal = GrandTotal + rs("TotalCost")
                    End If
					rs.MoveNext
				Loop
			Else
				Call NewOutputLaborBlankRow()
			End If
			
			Call NewOutputLaborTotals(GrandTotal)
			
		rw "</table><br>"
	rw "</fieldset>"

End Sub

Sub NewOutputMaterialsToolsBox(matRS,tlsRS,nowocheck)
	Dim BlankRowNum, OtherCostTotal, ChargeTotal
	ChargeTotal = 0
	OtherCostTotal = 0 

	rw "<fieldset style=""padding-top:14px"">"
		If nowocheck Then
			rw "<legend class=""legendHeader"">Materials (Summary)</legend>"
		Else
			rw "<legend class=""legendHeader"">Materials</legend>"
		End If
		rw "<table style=""margin-top:5px;"" border=""0"" cellspacing=""3"" cellpadding=""0"" width=""98%"" align=""center"">"

			Call NewOutputMaterialsHeader
						
			If Not matRS.eof and (NullCheck(matRS("WOPK")) = NullCheck(WOPK) or nowocheck) Then
				Do While Not matRS.eof and (NullCheck(matRS("WOPK")) = NullCheck(WOPK) or nowocheck)
					Call NewOutputMaterials(matRS)
					If isCharge Then
                    ChargeTotal = ChargeTotal + matRS("TotalCharge")
					OtherCostTotal = OtherCostTotal + matRS("OtherCost")
                    Else
                    ChargeTotal = ChargeTotal + matRS("TotalCost")
					OtherCostTotal = OtherCostTotal + matRS("OtherCost")
                    End If
					matRS.MoveNext
				Loop
			else
				Call NewOutputMaterialsBlankRow()
			End If
			
			Call NewOutputMaterialsTotals(OtherCostTotal, ChargeTotal)
			
		rw "</table><br>"
	rw "</fieldset>"
End Sub

Sub NewOutputGrandTotalBox(nowocheck)
	Dim GrandTotal
	GrandTotal = LaborTotal + MaterialsTotal + OtherCostsTotal
	rw "<fieldset style=""padding-top:14px"">"
		If nowocheck Then
			rw "<legend class=""legendHeader"">Totals (Summary)</legend>"
		Else
			rw "<legend class=""legendHeader"">Totals</legend>"
		End If
		rw "<table style=""margin-top:5px;"" border=""0"" cellspacing=""3"" cellpadding=""0"" width=""98%"" align=""center"">"
			rw "<tr>"
				rw "<td valign=""bottom"" class=""labels"" align=""left"" nowrap>Section</td>"
				rw "<td class=""labels"" width=""80"" align=""right"">Charge</td>"
			rw "</tr>"
			rw "<tr>"
				rw "<td class=""data_underline"" align=""left"">Labor&nbsp;Total</td>"
				rw "<td class=""data_underline"" align=""right"">" & FormatCurrency(LaborTotal) & "</td>"
			rw "</tr>"
			rw "<tr>"
				rw "<td class=""data_underline"" align=""left"">Materials&nbsp;Total</td>"
				rw "<td class=""data_underline"" align=""right"">" & FormatCurrency(MaterialsTotal) & "</td>"
			rw "</tr>"
			rw "<tr>"
				rw "<td class=""data_underline"" align=""left"">Other&nbsp;Costs&nbsp;Total</td>"
				rw "<td class=""data_underline"" align=""right"">" & FormatCurrency(OtherCostsTotal) & "</td>"
			rw "</tr>"
			rw "<tr>"
				rw "<td class=""labels"" align=""left"">Grand&nbsp;Total:</td>"
				rw "<td class=""data_underline"" align=""right"">" & FormatCurrency(GrandTotal) & "</td>"
			rw "</tr>"
		rw "</table><br>"
		
	rw "</fieldset>"
End Sub

Sub NewOutputMaterialsHeader()
	rw "<tr>"
		If WO_Barcode_PartID Then
			rw "<td valign=""bottom"" width=""1%"" class=""labels"">Barcode</td>"
		End If
		rw "<td valign=""bottom"" class=""labels"" align=""left"" nowrap>Item Name</td>"
		rw "<td valign=""bottom"" class=""labels"" width=""200"" align=""left"">Location</td>"
		rw "<td class=""labels"" width=""60"" align=""center"" nowrap>Quantity</td>"
		'rw "<td class=""labels"" width=""80"" align=""right"">Other&nbsp;Cost</td>"
		rw "<td class=""labels"" width=""80"" align=""right"">Charge</td>"
		
	rw "</tr>"
End Sub

Sub NewOutputMaterials(RS_WOPart)
	rw "<tr>"
		If WO_Barcode_PartID Then
			rw "<td style=""padding-right:10px;"" class=""data_underline"">"
				Call OutputBarCode(WO_Barcode_PartID,IN_BarcodeFormat_PartID,NullCheck(RS_WOPart("PartID")),"White")				
			rw "</td>"
		End If
		rw "<td class=""data_underline"">" & NullCheck(RS_WOPart("PartName")) & " (" & NullCheck(RS_WOPart("PartID")) & ")&nbsp;"
		If Not NullCheck(RS_WOPart("PartDescription")) = "" Then
			rw "<br/>" & NullCheck(RS_WOPart("PartDescription"))
		End If
		rw "</td>"				
		rw "<td class=""data_underline"">" & NullCheck(RS_WOPart("LocationID")) & "&nbsp;</td>"
		rw "<td class=""data_underline"" align=""center"">" & NullCheck(RS_WOPart("QuantityActual")) & "</td>"
		'rw "<td class=""data_underline"" align=""right"">" & FormatCurrency(NullCheck(RS_WOPart("OtherCost"))) & "</td>"
        If isCharge Then
		    rw "<td class=""data_underline"" align=""right"">" & FormatCurrency(NullCheck(RS_WOPart("TotalCharge"))) & "</td>"
        Else
		    rw "<td class=""data_underline"" align=""right"">" & FormatCurrency(NullCheck(RS_WOPart("TotalCost"))) & "</td>"
        End If
	rw "</tr>"
End Sub

Sub NewOutputMaterialsTotals(OtherTotal, ChargeTotal)
	MaterialsTotal = ChargeTotal
	rw "<tr>"
		rw "<td class=""labels"" colspan=""3"">Material Totals:</td>"
		If WO_Barcode_PartID Then
			rw "<td style=""padding-right:10px;"" class=""data_underline""></td>"
		End If
		'rw "<td class=""data_underline""></td>"
		'rw "<td class=""data_underline"" align=""center""></td>"
		'rw "<td class=""data_underline"" align=""right"">" & FormatCurrency(OtherTotal) & "</td>"
		rw "<td class=""data_underline"" align=""right"">" & FormatCurrency(MaterialsTotal) & "</td>"
	rw "</tr>"
End Sub

Sub NewOutputLaborBlankRow()	
	rw "<tr>"
		rw "<td class=""data_underline"">&nbsp;</td>"
		rw "<td class=""data_underline"" align=""center"">0</td>"
		rw "<td class=""data_underline"" align=""center"">0</td>"
		rw "<td class=""data_underline"" align=""center"">0</td>"				
		rw "<td class=""data_underline"" align=""left"">&nbsp;</td>"
		rw "<td class=""data_underline"" align=""right"">$0.00</td>"
	rw "</tr>"
End Sub

Sub NewOutputMaterialsBlankRow()	
	rw "<tr>"
		If WO_Barcode_PartID Then
			rw "<td style=""padding-right:10px;"" class=""data_underline"">&nbsp;</td>"
		End If
		rw "<td class=""data_underline"">&nbsp;</td>"
		rw "<td class=""data_underline"">&nbsp;</td>"
		rw "<td class=""data_underline"" align=""center"">0</td>"
		'rw "<td class=""data_underline"" align=""right"">$0.00</td>"
		rw "<td class=""data_underline"" align=""right"">$0.00</td>"
	rw "</tr>"
End Sub

Sub NewOutputOtherCostBlankRow()	
	rw "<tr>"
		rw "<td class=""data_underline"">&nbsp;</td>"
		rw "<td class=""data_underline"">&nbsp;</td>"
		rw "<td class=""data_underline"">&nbsp;</td>"
		'rw "<td class=""data_underline"">&nbsp;</td>"  --> Removed extra column 10/25/2016 Remi
		rw "<td class=""data_underline"" align=""right"">$0.00</td>"
	rw "</tr>"
End Sub

Sub NewOutputLabor(LaborFormat,RS_WOAssign)
	rw "<tr>"
		rw "<td class=""data_underline"">"
		rw NullCheck(RS_WOAssign("LaborName"))
		If Not NullCheck(RS_WOAssign("Comments")) = "" Then
		rw "<br/>" & NullCheck(RS_WOAssign("Comments"))
		End If
		rw "</td>"
		rw "<td class=""data_underline"" align=""center"">" & NullCheckNBSP(RS_WOAssign("RegularHours")) & "</td>"
		rw "<td class=""data_underline"" align=""center"">" & NullCheckNBSP(RS_WOAssign("OvertimeHours")) & "</td>"
		rw "<td class=""data_underline"" align=""center"">" & NullCheckNBSP(RS_WOAssign("OtherHours")) & "</td>"				
		rw "<td class=""data_underline"" align=""left"">" & DateNullCheck(RS_WOAssign("WorkDate")) & "</td>"
        If isCharge then
		    rw "<td class=""data_underline"" align=""right"">" & FormatCurrency(NullCheck(RS_WOAssign("TotalCharge"))) & "</td>"
        Else
		    rw "<td class=""data_underline"" align=""right"">" & FormatCurrency(NullCheck(RS_WOAssign("TotalCost"))) & "</td>"
        End If
	rw "</tr>"
End Sub

Sub NewOutputLaborTotals(Total)
	LaborTotal = Total
	rw "<tr>"
		rw "<td colspan=""5"" class=""labels"">Labor Totals:</td>"
		'rw "<td class=""data_underline"" align=""center""></td>"
		'rw "<td class=""data_underline"" align=""center""></td>"
		'rw "<td class=""data_underline"" align=""center""></td>"				
		'rw "<td class=""data_underline"" align=""left""></td>"
		rw "<td class=""data_underline"" align=""right"">" & FormatCurrency(LaborTotal) & "</td>"
	rw "</tr>"
End Sub

Sub NewOutputTaskBox(rs)

	Dim BlankRowNum

	If Not WO_TASKSECTION and Not wostate = "WOC" Then
		Exit Sub
	End If

	If rs.Eof Then
		If wostate = "WOC" Then
			Exit Sub		
		End If
		If Not WO_TASKSECTION_B Then
			Exit Sub
		End If
	End If		

	If rs.EOF Then 
		BlankRowNum = 3
	Else
		BlankRowNum = 1	
	End If
	rw "<fieldset style=""padding-top:14px"">"
		rw "<legend class=""legendHeader"">Tasks</legend>"
		rw "<table style=""margin-top:5px;"" border=""0"" cellspacing=""3"" cellpadding=""0"" width=""98%"" align=""center"">"

			Call OutputTasksHeader() 
			' task data 	
			If Not rs.EOF Then
				Do While NullCheck(rs("WOPK")) = NullCheck(WOPK)
					Call NewOutputTasks(rs)		
					rs.MoveNext
				Loop
			End If
										
			Call OutputBlankTaskRow(BlankRowNum,7)

		rw "</table>"
		rw "<br>"	
	rw "</fieldset>"

End Sub

Sub NewOutputTasks(RS_WOTask)
	rw "<tr>"
		If NullCheck(RS_WOTask("Header")) Then		
		rw "<td valign=""top"" class=""data"" colspan=""8"" style=""padding-top:8px""><b>" & NullCheck(RS_WOTask("TaskAction")) & "</b></td>"
		Else
		rw "<td valign=""top"" class=""data_underline"">" & NullCheck(RS_WOTask("TaskNo")) & "&nbsp;</td>"
		Dim LineStyle
		Select Case NullCheck(RS_WOTask("LineStyle"))
			Case "I1"
				LineStyle=" style=""padding-left:10px;"""
			Case "I2"
				LineStyle=" style=""padding-left:20px;"""
			Case "I3"
				LineStyle=" style=""padding-left:30px;"""
			Case "B"
				LineStyle=" style=""font-weight:bold;"""
			Case "BR"
				LineStyle=" style=""font-weight:bold;color:red;"""
			Case "I"
				LineStyle=" style=""font-style:italic;"""
			Case Else
				LineStyle=""
		End Select		
		rw "<td" & LineStyle & " valign=""top"" class=""data_underline"">" & Replace(NullCheck(RS_WOTask("TaskAction")),vbCrLf,"<br/>") & "&nbsp;"

        If Not NullCheck(RS_WOTask("Comments")) = "" Then
            If BitNullCheck(RS_WOTask("Fail")) Then
                rw "<br/><span style=""color:red; font-weight:bold;"">Comments: " & NullCheck(RS_WOTask("Comments")) & "</span>"
            Else
                rw "<br/><span style=""color:green; font-weight:bold;"">Comments: " & NullCheck(RS_WOTask("Comments")) & "</span>"
            End If
        End If

        rw "</td>"
		Select Case wostate

		Case "WO","WOC"
			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">" & NullCheck(RS_WOTask("Rate")) & "&nbsp;</td>"
			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">" & NullCheck(RS_WOTask("Measurement")) & "&nbsp;</td>"
			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">" & NullCheck(RS_WOTask("Initials")) & "&nbsp;</td>"
			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">"
			Call OutputBitFailure(RS_WOTask("Fail").Value)
			rw "</td>"
			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">"
			Call OutputBit(RS_WOTask("Complete").Value)			
			rw "</td>"

		Case "CC"		
			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">"
			rw "<input name=""TA_Rate_" & RS_WOTask("PK") & """ value=""" & NullCheck(RS_WOTask("Rate")) & """ class=""normalright"" mcType=""N"" maxlength=""12"" size=""1"" type=""text"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);"">"
			rw "</td>"
			
			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">"
			rw "<input name=""TA_Measurement_" & RS_WOTask("PK") & """ value=""" & NullCheck(RS_WOTask("Measurement")) & """ class=""normalright"" mcType=""N"" maxlength=""12"" size=""1"" type=""text"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);"">"
			rw "</td>"

			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">"
			rw "<input name=""TA_Initials_" & RS_WOTask("PK") & """ value=""" & NullCheck(RS_WOTask("Initials")) & """ class=""normal"" mcType=""N"" maxlength=""12"" size=""1"" type=""text"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);"">"
			rw "</td>"

			rw "<td valign=""bottom"" class=""data_underline_checkbox"" align=""center"">"
			rw "<input name=""TA_Fail_" & RS_WOTask("PK") & """ " & checkboxstate(RS_WOTask("Fail")) & "type=""checkbox"" class=""mccheckbox"" mcType=""B"" value=""ON"" onclick=""top.isdirty(this);"">"
			rw "</td>"
			
			rw "<td valign=""bottom"" class=""data_underline_checkbox"" align=""center"">"
			rw "<input name=""TA_Fail_" & RS_WOTask("PK") & """ " & checkboxstate(RS_WOTask("Complete")) & "type=""checkbox"" class=""mccheckbox"" mcType=""B"" value=""ON"" onclick=""top.isdirty(this);"">"
			rw "</td>"
		
		End Select
			
	rw "</tr>"
	End If
End Sub

Sub OutputMaintDetails()

    Dim RS_AssetTempValue
    RS_AssetTempValue = NullCheck(RS_WO("AssetPK"))
    If RS_AssetTempValue = "" Then
        RS_AssetTempValue = "-1"
    End If
    
	sql = "SELECT Asset.Icon, Asset.AssetPK, Asset.AssetID, Asset.AssetName, Asset.IsLocation, Asset.IsUp, Asset.Address, Asset.City, Asset.State, Asset.Zip " +_
		  "FROM Asset WITH (NOLOCK) INNER JOIN MC_GetAssetParentPK('" & RS_AssetTempValue & "') b On b.AssetPK = Asset.AssetPK " &_
		  "ORDER BY b.lvl DESC"
    
	Set RS_Asset = db.runSQLReturnRS(sql,"")
	Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

	rw "<table cellpadding=""3"" width=""100%"" border=""0"" style=""font-family:Arial;font-size:12px;color:#333333"">"
		rw "<tr>"
		rw "<td valign=""top"" width=""33%"" rowspan=""3"">"
			rw "<table border=""0"" cellspacing=""1"" cellpadding=""1"">"
				rw "<tr><td class=""labels"">Bill To:</td></tr>"									
					' asset hierarchy								
					NoIconFile_Location = "images/icons/facility_g.gif"
					NoIconFile_Asset = "images/icons/gearsxp_g.gif"

					skipfirstrow = True											
					If Not RS_Asset.EOF Then
					rw "<tr>"
					rw "<td valign=""top"" colspan=""2"">"
					rw "<table border=""0"" cellspacing=""0"" cellpadding=""0"">"
					'Do While (NullCheck(RS_Asset("WOPK")) = NullCheck(WOPK)) and (Not RS_Asset.Eof)
                    Do While (Not RS_Asset.Eof)							
						If RS_Asset("IsUp") Then
							stclass = "assetUp"												
						Else
							stclass = "assetDown"
						End If
								
						If RS_Asset("IsLocation") Then
							AssetName = RS_Asset("AssetName")
						Else
							AssetName = RS_Asset("AssetName") & " (" & RS_Asset("AssetID") & ")"
						End If			
						If NullCheck(RS_Asset("Icon")) = "" Then
							If RS_Asset("IsLocation") Then
								  IconFile = NoIconFile_Location
							Else
								  IconFile = NoIconFile_Asset				
							End If
						Else
							  IconFile = RS_Asset("Icon")
						End If 
												
						AssetOutput="<img src='" & GetSession("webHTTP") & GetWebServer() & Application("Web_Path") & GetPathFromSchema("mapp_path") & IconFile & "' border=0 style='margin-right:4'>" & AssetName & "<br>"
																								
						If skipfirstrow Then   
							RS_Asset.MoveNext
							If RS_Asset.Eof Then
								rw AssetOutput
								Exit Do			
							End If
							skipfirstrow = False
						Else
							RS_Asset.MoveNext
							rw "<tr>"
							'If (Not RS_Asset.Eof) and (NullCheck(RS_Asset("WOPK")) = NullCheck(WOPK)) Then
							If (Not RS_Asset.Eof) Then
								stclass = "asset"
							Else
                                AssetOutput = Left(AssetOutput,(Len(AssetOutput)-4)) & "-" & RS_WO("LeaseNumber")
                            End If
							rw "<td class=""" & stclass & """ valign=""top"">"
							rw AssetOutput
							rw "</td>"
							rw "</tr>"									
						End If											
					Loop 
					rw "</table>"
					rw "<br>"
					rw "</td>"
					rw "</tr>"
					End If
					
					'rw "<tr>"
					'  rw "<td valign=""top"">"
				    '  rw "<table cellspacing=""0"" cellpadding=""0"" border=""0"">"
				      'If WO_REPORT_MAINTDETAIL_SHOWASADDRESS = "Yes" Then
				      '  If NullCheck(RS_WO("Address")) <> "" Then
				      '    rw "<tr>"
					    '      rw "<td valign=""top"" class=""labels"">Address:&nbsp;</td>"
					    '      rw "<td valign=""top"" class=""data"">" 
					    '          rw NullCheck(RS_WO("Address")) 
					    '      rw "</td>"
				      '    rw "</tr>"	
			        '  End If
				      'End If
				      'If Not IsLocation Then
				      '  If WO_REPORT_MAINTDETAIL_SHOWASDETAILS = "Yes" Then
			            'If NullCheck(RS_WO("Model")) <> "" Then
			        '      rw "<tr>"
				      '        rw "<td valign=""top"" class=""labels"">Model:&nbsp;</td>"
				      '        rw "<td valign=""top"" class=""data"">" & NullCheck(RS_WO("Model")) & "</td>"
			        '      rw "</tr>"
			            'End If
			            'If NullCheck(RS_WO("Serial")) <> "" Then
			        '      rw "<tr>"
				      '        rw "<td valign=""top"" class=""labels"">Serial:</td>"
				      '        rw "<td valign=""top"" class=""data"">" & NullCheck(RS_WO("Serial")) & "</td>"
			        '      rw "</tr>"
			            'End If
			            'If NullCheck(RS_WO("ManufacturerName")) <> "" Then
			        '      rw "<tr>"
				      '        rw "<td valign=""top"" class=""labels"">Manufacturer:&nbsp;</td>"
				      '        rw "<td valign=""top"" class=""data"">" & NullCheck(RS_WO("ManufacturerName")) & "</td>"
			        '      rw "</tr>"
			            'End If
			            'If NullCheck(RS_WO("Vicinity")) <> "" Then							
			        '      rw "<tr>"
				      '        rw "<td valign=""top"" class=""labels"">Vicinity:</td>"
				      '        rw "<td valign=""top"" class=""data"">" & NullCheck(RS_WO("Vicinity")) & "</td>"
			        '      rw "</tr>"
			            'End If	
  			          			      
  				        
				      '  End If
				      'End If
				      'If WO_REPORT_MAINTDETAIL_SHOWREQUESTBYALLIN1PLACE = "No" Then
			          'rw "<tr>"
				      '    rw "<td valign=""top"" class=""labels"">Contact:&nbsp;</td>"
				      '    rw "<td valign=""top"" class=""data"">" & NullCheck(RS_WO("RequesterName")) & "</td>"
			          'rw "</tr>"							
			          'rw "<tr>"
				      '    rw "<td valign=""top"" class=""labels"">Phone:</td>"
				      '    rw "<td valign=""top"" class=""data"">" & NullCheck(RS_WO("RequesterPhone")) & "</td>"
			          'rw "</tr>"
				      'End If  						
				      'rw "</table>"
					  'rw "</td>"
					  'If WO_REPORT_MAINTDETAIL_SHOWASDETAILS = "Yes" Then
					  '  rw "<td style=""padding-left:10px;"">"
					    'If NullCheck(RS_WO("assetphoto")) <> "" Then
				      '  rw "<img src="""&assetphoto&""">"
				      'End If
					  '  rw "</td>"
				    'End If
					'rw "</tr>"
			  rw "</table>"					
		  rw "</td>"
		rw "<td valign=""top"" width=""33%"">"		
			rw "<table border=""0"" cellspacing=""1"" cellpadding=""1"">"
			    rw "<tr><td class=""labels"">Address:</td></tr>"
                rw "<tr>"
			      
			      rw "<td valign=""top"" class=""data"" colspan=""2"">" 
			          rw NullCheck(RS_WO("Address")) 
			      rw "</td>"
			    rw "</tr>"
				'If Not NullCheck(RS_WO("RequesterName")) = "" Then			
				'rw "<tr>"
				'	rw "<td nowrap class=""labels"" valign=""top"">Requested By:&nbsp;</td>"
				'	rw "<td valign=""top"" class=""data"">" & NullCheck(RS_WO("RequesterName"))
				'	If Not NullCheck(RS_WO("Requested")) = "" Then			
				'		rw " on " & DateTimeNullCheckAT(RS_WO("Requested")) & "&nbsp;"
				'	End If
				'	rw "</td>"
				'rw "</tr>"				
				'Else
				'rw "<tr>"
				'	rw "<td class=""labels"" nowrap valign=""top"">Requested:</td>"
				'	rw "<td class=""data"" valign=""top"">"
				'		rw DateTimeNullCheckAT(RS_WO("Requested")) & "&nbsp;"
				'	rw "</td>"
				'rw "</tr>"
				'End If
        'If WO_REPORT_MAINTDETAIL_SHOWREQUESTBYALLIN1PLACE = "Yes" Then
        '  rw "<tr>"
	      '    rw "<td valign=""top"" class=""labels"">Phone:</td>"
	      '    rw "<td valign=""top"" class=""data"">" & NullCheck(RS_WO("RequesterPhone")) & "</td>"
        '  rw "</tr>"  
        '  rw "<tr>"
	      '    rw "<td valign=""top"" class=""labels"">Email:</td>"
	      '    rw "<td valign=""top"" class=""data"">" & NullCheck(RS_WO("RequesterEmail")) & "</td>"
        '  rw "</tr>"  
        'End If

        'If CheckIfFieldExists(RS_WO,"TakenByName") Then
				'If Not NullCheck(RS_WO("TakenByName")) = "" Then	
				'  If WO_REPORT_MAINTDETAIL_SHOWTAKENBY = "Yes" then
				'    If Trim(NullCheck(RS_WO("RequesterName"))) <> Trim(NullCheck(RS_WO("TakenByName"))) Then
				'      rw "<tr>"
				'	      rw "<td nowrap class=""labels"" valign=""top"">Taken By:&nbsp;</td>"
				'	      rw "<td valign=""top"" class=""data"">" & NullCheck(RS_WO("TakenByName"))
				'	      If NullCheck(RS_WO("TakenByPhone")) <> "" Then
				'	        rw " (Work: " & NullCheck(RS_WO("TakenByPhone")) & ")"
				'	      End If
				'	      rw "</td>"
				'      rw "</tr>"	
				'    End If			  
				'  Else		
				'    rw "<tr>"
				'	    rw "<td nowrap class=""labels"" valign=""top"">Taken By:&nbsp;</td>"
				'	    rw "<td valign=""top"" class=""data"">" & NullCheck(RS_WO("TakenByName"))
				'	    rw "</td>"
				'    rw "</tr>"				
				'  End If
				'End If
				'End If
				
				'If Not NullCheck(RS_WO("ProblemID")) = "" Then			
				'rw "<tr>"
				'	rw "<td class=""labels"" valign=""top"">Problem:</td>"
				'	rw "<td class=""data"" valign=""top"">" & NullCheck(RS_WO("ProblemName"))
				'	If Not NullCheck(RS_WO("ProblemID")) = "" Then
				'		rw " (" & NullCheck(RS_WO("ProblemID")) & ")</td>"
				'	Else
				'		rw "</td>"
				'	End If
				'rw "</tr>"		
				'End If
				
                'If WO_REPORT_TASK_SHOWPMINFO = "No" Then
                '  If Not NullCheck(RS_WO("ProcedureID")) = "" Then													
				'  rw "<tr>"
				'	  rw "<td class=""labels"" valign=""top"">Procedure:</td>"
				'	  rw "<td valign=""top"" class=""data"">" & NullCheck(RS_WO("ProcedureName"))
				'	  If Not NullCheck(RS_WO("ProcedureID")) = "" Then
				'		  rw " (" & NullCheck(RS_WO("ProcedureID")) & ")"
				'	  End If
				'	  rw "</td>"
				'  rw "</tr>"
				'  End If
				'End If
				
				'If Not NullCheck(RS_WO("ProjectName")) = "" Then
				'rw "<tr>"
				'	rw "<td class=""labels"" valign=""top"">Project:</td>"
				'	rw "<td class=""data"" valign=""top"">" & NullCheck(RS_WO("ProjectName"))
				'	If Not NullCheck(RS_WO("ProjectID")) = "" Then
				'		rw " (" & NullCheck(RS_WO("ProjectID")) & ")"
				'	End If
				'	rw "</td>"
				'rw "</tr>"								
				'End If				
			rw "</table>"		
		rw "</td>"
		rw "<td valign=""top"" width=""33%"">"
			rw "<table border=""0"" cellspacing=""1"" cellpadding=""1"">"
				rw "<tr>"
					rw "<td nowrap class=""labels"" valign=""top"">Requested By:&nbsp;</td>"
					rw "<td valign=""top"" class=""data"">" & NullCheck(RS_WO("RequesterName"))
					If Not NullCheck(RS_WO("Requested")) = "" Then			
						rw " on " & DateTimeNullCheckAT(RS_WO("Requested")) & "&nbsp;"
					End If
					rw "</td>"
				rw "</tr>"				
				rw "<tr>"
					rw "<td class=""labels"" valign=""top"">Problem:</td>"
					rw "<td class=""data"" valign=""top"">" & NullCheck(RS_WO("ProblemName"))
					If Not NullCheck(RS_WO("ProblemID")) = "" Then
						rw " (" & NullCheck(RS_WO("ProblemID")) & ")</td>"
					Else
						rw "</td>"
					End If
				rw "</tr>"
		        rw "<tr>"
			        rw "<td class=""labels"" valign=""top"">Reason:</td>"
                    rw "<td class=""data"">" & Replace(NullCheck(RS_WO("Reason")),"%0D%0A"," ") & "</td>"
		        rw "</tr>"
				'rw "<tr>"
				'	rw "<td class=""labels"" valign=""top"">Target:</td>"
				'	rw "<td class=""data"" valign=""top"">" & DateNullCheck(RS_WO("TargetDate")) & "&nbsp;"
					'If WO_REPORT_MAINTDETAIL_SHOWTARGETHOURS = "Yes" Then
				'	  If Not NullCheck(RS_WO("TargetHours")) = "" Then
				'	    If CLng(RS_WO("TargetHours")) > 1 Then
				'	    rw "(" & CStr(RS_WO("TargetHours")) & ") hrs"
				'	    Else
				'	    rw "(" & CStr(RS_WO("TargetHours")) & ") hr"
				'	    End If					
				'	  End If
					'End If
				'	rw "</td>"
				'rw "</tr>"
				rw "<tr>"
					rw "<td class=""labels"" valign=""top"">Priority/Type:&nbsp;</td>"
					rw "<td class=""data"" valign=""top"">" & NullCheck(RS_WO("PriorityDesc")) 
					    If Not NullCheck(RS_WO("TypeDesc")) = "" Then
					        rw "/ " & NullCheck(RS_WO("TypeDesc")) 
					    End If
				        rw "&nbsp;&nbsp;&nbsp;"
                        Call OutputBit(NullCheck(RS_WO("Chargeable")))
				        rw "Credit"
					rw "</td>"
				rw "</tr>"			
				'If Not NullCheck(RS_WO("SupervisorName")) = "" Then					
				'rw "<tr>"
				'	rw "<td class=""labels"" valign=""top"">Supervisor:&nbsp;</td>"
				'	rw "<td class=""data"" valign=""top"">" & NullCheck(RS_WO("SupervisorName")) & "</td>"
				'rw "</tr>"								
				'End If
				'If Not NullCheck(RS_WO("ShopName")) = "" Then
				'rw "<tr>"
				'	rw "<td class=""labels"" valign=""top"">Shop:</td>"
				'	rw "<td class=""data"" valign=""top"">" & NullCheck(RS_WO("ShopID"))
				'	rw "</td>"
				'rw "</tr>"								
				'End If	
				'If WO_REPORT_MAINTDETAIL_SHOWDEPARTMENT = "Yes" Then
				'  If Not NullCheck(RS_WO("DepartmentName")) = "" Then
				'    rw "<tr>"
				'	    rw "<td class=""labels"" valign=""top"">Department:</td>"
				'	    rw "<td class=""data"" valign=""top"">" & NullCheck(RS_WO("DepartmentName"))
				'	    rw "</td>"
				'    rw "</tr>"								
				'  End If
				'End If	
				'If WO_REPORT_MAINTDETAIL_SHOWACCOUNT = "Yes" Then
				'  If NullCheck(RS_WO("AccountName")) <> "" Then
				'    rw "<tr>"
				'	    rw "<td class=""labels"" valign=""top"">Account:</td>"
				'	    rw "<td class=""data"" valign=""top"">" & NullCheck(RS_WO("AccountName"))
				'	    rw "</td>"
				'    rw "</tr>"								
				'  End If
				'End If						
			rw "</table>"						
		rw "</td>"

		rw "</tr>"
		'rw "<tr>"
		'	rw "<td valign=""top"" colspan=""2""><span class=""labels"">Reason:&nbsp;</span><span class=""data"">" & Replace(NullCheck(RS_WO("Reason")),"%0D%0A"," ") & "</span>"
		'rw "</tr>"
		'Dim CombinedInstructions
		'CombinedInstructions=""
		'If NullCheck(RS_WO("Instructions")) <> "" Then
		'  CombinedInstructions = Trim(NullCheck(RS_WO("Instructions")))
		'End If
		'If NullCheck(RS_WO("InstructionsAsset")) <> "" Then
		'  CombinedInstructions = CombinedInstructions + "&nbsp;&nbsp;" + Trim(NullCheck(RS_WO("InstructionsAsset")))
	    'End If
		'If NullCheck(CombinedInstructions) <> "" Then
		'  rw "<tr>"
		'    rw "<td valign=""top"" colspan=""2""><span class=""labels"">Special Instructions:&nbsp;</span><span class=""data"">" & Replace(CombinedInstructions,"%0D%0A"," ") & "</span>"
		'  rw "</tr>"						
		'End If
		'rw "<tr>"
		'	rw "<td valign=""top"" colspan=""2"" class=""data"">"
		'		rw "<br>"
				'If WO_REPORT_MAINTDETAIL_SHOWONLYCHECKED = "Yes" Then
				'  If BitNullCheck(RS_WO("WarrantyBox")) Then
				'    Call OutputBit(NullCheck(RS_WO("WarrantyBox")))
				'    rw "Warranty&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				'  End If
				'  If BitNullCheck(RS_WO("ShutdownBox")) Then
				'    Call OutputBit(NullCheck(RS_WO("ShutdownBox")))
				'    rw "Shutdown&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				'  End If
				'  If BitNullCheck(RS_WO("LockoutTagoutBox")) Then
				'    Call OutputBit(NullCheck(RS_WO("LockoutTagoutBox")))
				'    rw "Lockout&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				'  End If
				'  If BitNullCheck(RS_WO("AttachmentsBox")) Then
				'    Call OutputBit(NullCheck(RS_WO("AttachmentsBox")))
				'    rw "Attach&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				'  End If
				'  If BitNullCheck(RS_WO("Chargeable")) Then
				'    Call OutputBit(NullCheck(RS_WO("Chargeable")))
				'    rw "Charge&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				'  End If
				'Else
				'  Call OutputBit(NullCheck(RS_WO("WarrantyBox")))
				'  rw "Warranty&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				'  Call OutputBit(NullCheck(RS_WO("ShutdownBox")))
				'  rw "Shutdown&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				'  Call OutputBit(NullCheck(RS_WO("LockoutTagoutBox")))
				'  rw "Lockout&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				'  Call OutputBit(NullCheck(RS_WO("AttachmentsBox")))
				'  rw "Attach&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				'  Call OutputBit(NullCheck(RS_WO("Chargeable")))
				'  rw "Credit&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"				
				'End If
				'Call OutputBit(NullCheck(RS_WO("FollowupWork")))
				'rw "Follow-up Work&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		'	rw "</td>"
		'rw "</tr>"						
	rw "</table>"
End Sub

Sub OutputWOReport()

	Dim compdate,closestate
	
	compdate = DateNullCheck(RS_WO("Closed"))
	If compdate = "" Then
		compdate = CDate(Date())
	End If

	closestate = Not RS_WO("IsOpen")

	rw "<table style=""margin-top:5px;"" border=""0"" cellspacing=""0"" cellpadding=""0"" width=""98%"" align=""center"">"
		rw "<tr>"
			rw "<td class=""data"" width=""20%"">"
				rw "<table border=""0"" cellspacing=""3"" cellpadding=""0"">"
					rw "<tr>"
						rw "<td nowrap class=""labels"" valign=""bottom"">Completed:&nbsp;</td>"
						Select Case wostate
						Case "WO"						
						rw "<td class=""data_underline"" width=""100%"">&nbsp;</td>"
						Case "CC"
						
						rw "<td class=""data_underline"" width=""100%"">"
							rw "<input class=""normal"" mcType=""D"" maxlength=""10"" mcRequired=""N"" type=""text"" name=""WO_Completed_" & RS_WO("WOPK") & """ value=""" & compdate & """ size=""8"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.checkKey(this,self);""><img src=""../../images/lookupiconxp3.gif"" border=""0"" onclick=""top.showpopup('calendar','Calendar',172,160,this,WO_Completed_" & RS_WO("WOPK") & ",self)"" align=""absbottom"" class=""lookupicon"" WIDTH=""16"" HEIGHT=""20"">"
							rw "<span style=""display:none;"" id=""WO_Completed_" & RS_WO("WOPK") & "Err"" class=""mc_lookupdesc""></span>"
						rw "</td>"												
						Case "WOC"
						rw "<td class=""data_underline"" width=""100%"">" & NullCheckNBSP(DateTimeNullCheckAT(RS_WO("Complete"))) & "</td>"						
						End Select						
					rw "</tr>"
				rw "</table>"
			rw "</td>"
			'rw "<td class=""data"" width=""30%"">"
			'	rw "<table border=""0"" cellspacing=""3"" cellpadding=""0"">"
			'		rw "<tr>"
			'			rw "<td nowrap class=""labels"" valign=""bottom"">Failure:&nbsp;</td>"
			'			Select Case wostate
			'			Case "WO"						
			'			rw "<td class=""data_underline"" width=""100%"">&nbsp;</td>"
			'			Case "CC"
			'			rw "<td class=""data_underline"" width=""100%"">"
			'				rw "<input class=""normal"" mcType=""C"" maxlength=""25"" mcRequired=""N"" type=""text"" name=""WO_Failure_" & RS_WO("WOPK") & """ value=""" & NullCheck(RS_WO("FailureID")) & """ size=""10"" onChange=""top.dovalid('FA',this,'WO');"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.checkKey(this,self);""><img src=""../../images/lookupiconxp3fk.gif"" border=""0"" align=""absbottom"" onclick=""top.dolookup('FA',WO_Failure_" & RS_WO("WOPK") & ",'WO')"" class=""lookupicon"" width=""16"" height=""20"">"
			'				rw "<span style=""display:none;"" id=""WO_Failure_" & RS_WO("WOPK") & "Desc"" class=""mc_lookupdesc"">" & NullCheck(RS_WO("FailureName")) & "</span>"
			'				rw "<input type=""hidden"" name=""WO_Failure_" & RS_WO("WOPK") & "PK"" class=""mc_pluggedvalue"">"				
			'			rw "</td>"												
			'			Case "WOC"
			'			rw "<td class=""data_underline"" width=""100%"">" & NullCheckNBSP(RS_WO("FailureID")) & " / " & NullCheckNBSP(RS_WO("FailureName")) & "</td>"						
			'			End Select						
			'		rw "</tr>"
			'	rw "</table>"
			'rw "</td>"
			'rw "<td style=""padding-left:7px;"" class=""data"" width=""35%"">"
			  'If IsMetered Then
			'	  rw "<table border=""0"" cellspacing=""3"" cellpadding=""0"" width=""100%"">"
			'		  rw "<tr width=""100%"">"
						  
			'			  Select Case wostate
			'			  Case "WO"	
			'          rw "<td class=""labels"" width=""5%"" nowrap>Meter 1: </td>"
			'          rw "<td class=""data_underline"" width=""45%"">&nbsp;</td>"
			'          rw "<td width=""4%"">&nbsp;</td>"
			'          rw "<td class=""labels"" width=""5%"" nowrap>Meter 2: </td>"
			'          rw "<td class=""data_underline"" width=""45%"">&nbsp;</td>"						          
			'			  Case "CC"
			'			    rw "<td nowrap class=""labels"" valign=""bottom"">Meter(s):&nbsp;</td>"
			'			    rw "<td nowrap class=""data_underline"" width=""100%"">"
			'			    rw "<input name=""WO_Meter1Reading_" & RS_WO("WOPK") & """ value=""" & NullCheck(RS_WO("Meter1Reading")) & """ class=""normalright"" mcType=""N"" maxlength=""12"" mcRequired=""N"" type=""text"" size=""6"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);""><img src=""../../images/lookupiconxp3.gif"" border=""0"" onclick=""top.showpopup('calculator','Calculator',125,100,this,WO_Meter1Reading_" & RS_WO("WOPK") & ",self)"" align=""absbottom"" class=""lookupicon"" WIDTH=""16"" HEIGHT=""20"">"
			'			    rw "&nbsp;/&nbsp;<input name=""WO_Meter2Reading_" & RS_WO("WOPK") & """ value=""" & NullCheck(RS_WO("Meter2Reading")) & """ class=""normalright"" mcType=""N"" maxlength=""12"" mcRequired=""N"" type=""text"" size=""6"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);""><img src=""../../images/lookupiconxp3.gif"" border=""0"" onclick=""top.showpopup('calculator','Calculator',125,100,this,WO_Meter1Reading_" & RS_WO("WOPK") & ",self)"" align=""absbottom"" class=""lookupicon"" WIDTH=""16"" HEIGHT=""20"">"
			'			    rw "</td>"						
			'			  Case "WOC"
			'			    rw "<td nowrap class=""labels"" valign=""bottom"">Meter(s):&nbsp;</td>"
			'			    If RS_WO("Meter1Reading") > "0" and RS_WO("Meter2Reading") > "0" Then
			'				    rw "<td class=""data_underline"" width=""100%"">" & NullCheckNBSP(RS_WO("Meter1Reading")) & " " & NullCheckNBSP(RS_WO("Meter1UnitsDesc")) & "&nbsp;/&nbsp;" & NullCheckNBSP(RS_WO("Meter2Reading")) & " " & NullCheckNBSP(RS_WO("Meter2UnitsDesc")) &  "</td>"						
			'			    ElseIf RS_WO("Meter1Reading") > "0" Then
			'				    rw "<td class=""data_underline"" width=""100%"">" & NullCheckNBSP(RS_WO("Meter1Reading")) & " " & NullCheckNBSP(RS_WO("Meter1UnitsDesc")) &  "</td>"						
			'			    ElseIf RS_WO("Meter2Reading") > "0" Then
			'				    rw "<td class=""data_underline"" width=""100%"">" & NullCheckNBSP(RS_WO("Meter2Reading")) & " " & NullCheckNBSP(RS_WO("Meter2UnitsDesc")) &  "</td>"						
			'			    Else
			'			    rw "<td class=""data_underline"" width=""100%"">&nbsp;</td>"
			'			    End If
			'			  End Select						
			'		  rw "</tr>"																		
			'	  rw "</table>"		
				'End If				
			'rw "</td>"
		rw "</tr>"

		rw "<tr style=""margin-bottom:20px;"">"
			rw "<td class=""labels"">"
				rw "<table style=""margin-top:10px;"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
					If wostate = "CC" Then
						rw "<tr>"
						rw "<td colspan=""3"" width=""100%"" align=""right"">"
						rw "<img border=""0"" src=""../../images/button_addarrow.gif"" onclick=""top.showpopup('actions','Actions',266,100,this,WO_LaborReport_" & RS_WO("WOPK") & ",self)"" WIDTH=""80"" HEIGHT=""15"">"
						rw "</td>"
						rw "</tr>"
					End If
					rw "<tr>"
					rw "<td valign=""top"" style=""padding-left:3px;"" style=""width:10%;"" nowrap class=""labels"">"
						rw "Report:&nbsp;&nbsp;"
					rw "</td>"
                    if WOStatus = "OPEN" Then
                        rw "<td style=""width:100%"">"
                            rw "<table style=""width:100%"">"
		                        rw "<tr class=""blank_row"">"
			                        rw "<td class='data_underline'>&nbsp;</td>"
		                        rw "</tr>"
		                        rw "<tr class=""blank_row"">"
			                        rw "<td class='data_underline'>&nbsp;</td>"
		                        rw "</tr>"
                            rw "</table>"
                        rw "</td>"
                    Else
                    	rw "<td class=""data_underline"" width=""100%"">" & Replace(NullCheckNBSP(RS_WO("LaborReport")),"%0D%0A","<br>") & "</td>"						
                    End If
					'Select Case wostate
					'Case "WO"	
					  'If WO_REPORT_LABORRPT_SHOWBLANKLINES > 0 Then
					'    Call OutputBlankRow(WO_REPORT_LABORRPT_SHOWBLANKLINES,1)
					  'End If
					'Case "CC"
					'rw "<td width=""100%"">"						
					'rw "<textarea mcType=""C"" name=""WO_LaborReport_" & RS_WO("WOPK") & """ wrap=""hard"" style=""width: 100%; height:35px;"" class=""normal"" onfocus=""top.fieldfocus(this);"" onkeypress=""return top.checkKey(this,self);"" onblur=""top.fieldblur(this);"" onChange=""top.fieldvalid(this);"" rows=""1"" cols=""20"">" & Replace(NullCheckNBSP(RS_WO("LaborReport")),"%0D%0A",chr(13)+chr(10)) & "</textarea>"
					'rw "</td>"
					'Case "WOC"
					'rw "<td class=""data_underline"" width=""100%"">" & Replace(NullCheckNBSP(RS_WO("LaborReport")),"%0D%0A","<br>") & "</td>"						
					'End Select											
					'rw "</tr>"																		
					'If wostate = "CC" Then
					'rw "<tr>"
					'rw "<td colspan=""4"">"

					'rw "<table cellspacing=""0"" cellpadding=""0"" width=""100%"" class=""normaltext"" style=""margin-top:10px;"">"
					'rw "<tr>"
					
					'rw "<td>Set Tasks: Completed / Failed</td>"
					'rw "<td>Set Actuals = Estimates: All / Labor / Materials / Other Costs</td>"
					'rw "<td>Set Labor Hours...</td>"
					
					'rw "</tr>"
					'rw "</table>"
					
					'rw "</td>"
					rw "</tr>"
					'End If
				rw "</table>"						
			rw "</td>"					
		rw "</tr>"
		'If WO_REPORT_LABORRPT_SHOWSIGNATURE = "Yes" Then
		rw "<tr style=""padding-top:40px;"">"
			rw "<td colspan=""4"" class=""labels"">"
				rw "<table style=""margin-top:10px;"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
				  rw "<tr>"
		            rw "<td class=""data_underline"" width=""60%"">&nbsp;</td>"
		            rw "<td width=""2%"">&nbsp;</td>"
		            rw "<td class=""data_underline"" width=""38%"">&nbsp;</td>"
		            'rw "<td width=""4%"">&nbsp;</td>"
		            'rw "<td class=""data_underline"" width=""30%"">&nbsp;</td>"
		            'rw "<td width=""2%"">&nbsp;</td>"
		            'rw "<td class=""data_underline"" width=""16%"">&nbsp;</td>"
		          rw "</tr>"
		          rw "<tr>"
		            rw "<td valigh=""top"" class=""labels"" align=""center"">Signature / Name</td>"
		            rw "<td valigh=""top"" style=""font-size:xx-small;"" align=""center""></td>"
		            rw "<td valigh=""top"" class=""labels"" align=""center"">Date</td>"
		            'rw "<td valigh=""top"" style=""font-size:xx-small;"" align=""center""></td>"
		            'rw "<td valigh=""top"" style=""font-size:xx-small;"" align=""center"">Signature / Name</td>"
		            'rw "<td valigh=""top"" style=""font-size:xx-small;"" align=""center""></td>"
		            'rw "<td valigh=""top"" style=""font-size:xx-small;"" align=""center"">Date</td>"
		          rw "</tr>"
		        rw "</table>"
		      rw "</td>"
		    rw "</tr>"
		    'End If
	rw "</table>"
End Sub

Sub OutputHeader()

	If SubReport Then
		Exit Sub
	End If

	' BRAD STARTSPAN
	If BatchReportMode Then
		'Exit Sub
	Else
        If UCase(GetSession("WEBHTTP")) = "HTTPS://" and _
           Not Request.QueryString("ExportReportOutputType") = "" and _
           Not Request.QueryString("ExportReportOutputType") = "HTM" Then
            ' Do not add the cache-control as IE has a bug with SSL and downloads
            ' http://support.microsoft.com/default.aspx?kbid=323308
        Else
    		Response.AddHeader "cache-control","no-store"
        End If
	End If
	' BRAD ENDSPAN

    rw "<!DOCTYPE html>"
    rw "<html dir=""ltr"" lang=""en-us"">"
    rw "<head>"
    rw "<meta content=""IE=edge"" http-equiv=""X-UA-Compatible"" />"
    rw "<meta content=""text/html;charset=UTF-8"" http-equiv=""Content-type"" />"
    rw "<meta content=""text/javascript"" http-equiv=""content-script-type"" />"
    rw "<meta content=""text/css"" http-equiv=""content-style-type"" />"
    rw "<meta content=""favorite"" name=""save"">"
    rw "<meta content=""noindex"" name=""robots"">"
    rw "<meta content=""no"" name=""allow-search"" />"
    rw "<meta content=""yes"" name=""apple-mobile-web-app-capable"" />"
    rw "<meta content=""black"" name=""apple-mobile-web-app-status-bar-style"" />"
	rw "<title>" & reportName & "</title>"
	If Not FromAgent Then
		If Not Request.QueryString("ExportReportOutputType") = "HTM" Then %>
			<script type="text/javascript" SRC="../../javascript/normal/mc_rpt_common.js"></script>
            <script type="text/javascript" SRC="../../javascript/normal/mc_flashobject.js"></script>
            <% If SCSearchAny Then %>
			<script type="text/javascript" SRC="../../javascript/normal/mc_tablefilter_all.js"></script>
			<% End If %>

			<script type="text/javascript">
				//window.onerror = top.doError;

			    var reportid = "<% =reportid %>";
			    var reportidc = "<% =reportidc %>";
				var reportname = "<% =reportName %>";
				var custom = <% =LCase(custom) %>;
				var reportcopy = <% =LCase(reportcopy) %>;
				var showSQL = "<% =Request.QueryString("showSQL") %>";
				var reporthasfields = <% =LCase(reporthasfields) %>;

                try{
                var sql_where = '<% =JSEncode(SQL_Where) %>';
                //alert(sql_where);
                } catch(e) {};

	            function psr(ReportID,extracrit)
	            {
	                try {
	                //alert(extracrit);
	            	event.stopPropagation();
	            	url = 'modules/reports/rpt_check.asp?rptID='+ReportID+'&noprompt=true&showSQL=false&popupreport=Y&sqlwhere='+self.sql_where+' '+extracrit;
                    //alert(self.location.href);
                    //alert(self.sql_where);
	            	var param = new Object();
	            	param.caller = mcmain;
	            	param.mcmain = mcmain;
	            	aW = screen.availWidth;
	            	aH = screen.availHeight;
	            	if (aW >= ((1024*2)-0) || aH >= ((768*2)-0))
	            	{
	            	 	if (aW >= ((1024*2)-0))
	            	 	{aW = aW / 2;}
	            	 	else
	            	 	{aH = aH / 2;}
	            	}
	            	aW -= 0;
	            	aH -= 0;
                    var finalurl = mcmain.path + 'modules/common/mc_dialogrefreshable.asp?scroll=auto&title=Report+Preview&url='+escape(mcmain.path) + url;
                    //alert(finalurl.length);
                    //alert(finalurl);
	                    //mcmain.showModelessDialog(finalurl, param,'dialogHeight: ' + aH + 'px; dialogWidth: ' + aW + 'px; dialogTop: 0px; dialogLeft: 0px; center: No; help: No; resizable: No; status: No; scroll: Auto' );
                    //alert(top.path + url);
                    top.opendhtmlwin(top.path + url, param, 0, 0, aW, aH, true, false);
	            	} catch(e) {};
	            }

                function fixedEncodeURIComponent(str){
                     return encodeURIComponent(str).replace(/[!'()]/g, escape).replace(/\*/g, "%2A");
                }

				function checkhandle()
				{
				    try {
		            var x = window.event.clientX;
		            var y = window.event.clientY;

                    if (x < 8)
                    {
                        parent.togglesmartcritpane(reportid,'ON',false);
                    }
                    window.event.returnValue = true;
                    } catch(e) {};
				    return true;
				}

				function doonload()
				{
				    // the try must be in there for Smart Criteria Left Pane Reports
				    // as togglesmartcritpane would not exist...etc.
				    if (document.getElementById('tbl-container-div'))
				    {
				        if (self.innerWidth > parseFloat(document.getElementById('tbl-container-div').style.width))
				        {
				            document.getElementById('tbl-container-div').style.width = self.innerWidth - 30 + 'px';
				        }
				    }

				    try {
				    <% If (SCDefault = "S" or SCDefault = "F") and SCAvailable Then %>
		            parent.togglesmartcritpane(reportid,'ON',true);
				    <% Else %>
				    <% If (SCDefault = "H" or SCDefault = "C") and SCAvailable Then %>
		            parent.togglesmartcritpane(reportid,'OFF',true);
				    <% Else %>
		            parent.togglesmartcritpane(reportid,'DISABLED',true);
			        <% End If %>
				    <% End If %>
				    } catch(e) {};
				    try {
				    if (self.fixscroll)
				    {
				        fixscroll();
				    }
				    } catch(e) {};

				    <% If wtf Then %>
					if (top.topwin.sendend)
					{
					    top.topwin.sendend();
					}
				    <% Else %>
                    <% If Not InlineCriteria > 0 Then %>
					self.focus();
                    <% End If %>
					<% End If %>

					<% If InWOModule and Not wtf Then %>
					top.endprocessgeneric();
					if (parent && parent.current_ra)
					{
						parent.current_ra.style.display = 'none';
					}
					if (parent && parent.ra_<% =LCase(recordstatus) %>)
					{
						parent.ra_<% =LCase(recordstatus) %>.style.display = '';
						parent.current_ra = parent.ra_<% =LCase(recordstatus) %>;
					}
					parent.document.mcform.txtIsApproved.value = '<% =recordauthstatus %>';
					<% End If %>

					<% If reporthasfields Then %>
					document.onclick = rptdoc_click;
					document.onkeydown = top.doc_keydown;

					// set focus to first textarea if it exists
					try {
					var textareatags = document.getElementsByTagName( "textarea" );
					var slength = textareatags.length;
					if (slength > 0)
					{
						textareatags[0].focus();
					}
					}
					catch(e) {}
					<% End If %>
            		<% If (Not SubReport) and (SLDefault) and (Not SDCode = "") Then %>
                    try {
                    var ctall=document.body.all;
                    for(i=0;i<ctall.length;i++){
                    if(((ctall(i).tagName.toUpperCase()=="INPUT"&&ctall(i).type.toUpperCase()=="TEXT")||ctall(i).tagName.toUpperCase()=="TEXTAREA")&&(ctall(i).disabled==false)&&(ctall(i).style.display==''))
                    {
                        try {
                            //alert(ctall(i).outerHTML);                            
                            ctall(i).focus();
                            break;
                        }
                        catch(e) { break; };
                    }
                    }
                    } catch(e) {};
                    <% End If %>

                    <% If SCSearchAny and Not recordCount = 0 Then %>
                    //try {
                    var tfConfig = {highlight_keywords: true, 
                                        alternate_rows: true,  
                                        selectable: false,  
                                        <% If SCSearch = "YS" or SCSearch = "YSS" Then %>
                                        single_search_filter: true,
                                        <% End If %>
                                        enable_empty_option: true,
                                        empty_text: '(No Value)',
                                        enable_non_empty_option: false,
                                        <% If GetSession("dff") = "Y" Then %>
                                        default_date_type: 'DMY',
                                        <% End If %>
                                        non_empty_text: '(Any Value)',
                                        filters_row_index: 1,
                                        <% If SCRpt Then %>
                                        odd_row_css_class: "ReportRowCrit1",
                                        even_row_css_class: "ReportRowCrit2",
                                        <% Else %>
                                        odd_row_css_class: "ReportRow1",
                                        even_row_css_class: "ReportRow2",
                                        <% End If %>
                                        <% If RecordCount <= 2500 Then %>
                                        on_keyup: true,  
                                        on_keyup_delay: 100,
                                        <% End If %>
                                        loader: false<%
                                        Dim scsearchrows, scsearchrow,scsearchids,scsearchcols,scsearchops,scsearchdecs
                                        If Not RecordCount = 0 Then
                                            scsearchrows = UBound(ReportFields,2)
                                            For scsearchrow = 0 to scsearchrows
                                                If ReportFields(4,scsearchrow) Then
                                                    If scsearchids = "" Then
                                                        scsearchids = """gtot" & CStr(scsearchrow) & """"
                                                        scsearchcols = CStr(scsearchrow)
                                                        scsearchops = """sum"""
                                                        scsearchdecs = "0"
                                                    Else
                                                        scsearchids = scsearchids & "," & """gtot" & CStr(scsearchrow) & """"
                                                        scsearchcols = scsearchcols & "," & CStr(scsearchrow)
                                                        scsearchops = scsearchops & "," & """sum"""
                                                        scsearchdecs = scsearchdecs & "," & "0"
                                                    End If
                                                End If
		                                        Select Case NullCheck(ReportFields(26,scsearchrow))
                                                    Case "1"
                                                    Response.Write ",col_" & CStr(scsearchrow) & ": ""select"""
                                                    Case "2"
                                                    Response.Write ",col_" & CStr(scsearchrow) & ": ""multiple"""
                                                    Case "3"
                                                    Response.Write ",col_" & CStr(scsearchrow) & ": ""checklist"""
                                                    Case "4"
                                                    Response.Write ",col_" & CStr(scsearchrow) & ": ""none"""
                                                    End Select
                                            Next
                                        End If %>
                                        <% If Not scsearchids = "" Then %>,
                                        col_operation: {   
                                            id: [<% =scsearchids %>],  
                                            col: [<% =scsearchcols %>],  
                                            operation: [<% =scsearchops %>],  
                                            exclude_row: ["totRowIndex"],  
                                            tot_row_index: ["totRowIndex"],
                                            decimal_precision: [<% =scsearchdecs %>]  
                                        }
                                        <% End If %>
                    };
                    var tf1 = setFilterGrid("tbl",tfConfig);                

                    for(var w = 0; w < document.getElementById('tbl').rows[1].cells.length; w++) 
                    {
                        //console.log(document.getElementById('tbl').rows[1].cells[w]);
                        document.getElementById('tbl').rows[0].cells[w].className = document.getElementById('tbl').rows[1].cells[w].className + ' headingfixed'; // + tbl.rows[1].cells[w].className;
                    }

                    //vartotrow = document.getElementById('totRowIndex');
                    //if (vartotrow) {
                        //for(var w = 0; w < vartotrow.cells.length; w++) 
                        //{
                            //vartotrow.cells[w].className = vartotrow.cells[w].className + ' footerfixed'; // + tbl.rows[1].cells[w].className;
                        //}
                    //}
                    //alert(vartotrow.outerHTML);

                    //} catch(e) {alert('Problem with Smart Search');}
                    <% End If %>

                    try {
                    //new way to handle print events for chrome and other modern browsers
                    if (window.matchMedia) {
                        var mediaQueryList = window.matchMedia('print');
                        mediaQueryList.addListener(function(mql) {
                            if (mql.matches) {
                                rpt_onbeforeprint();
                            } else {
                                rpt_onafterprint();
                            }
                        });
                    }
                    } catch(e) {}

			        try {
			            //top.walk(self.document.body, null, self);
			            top.walk(self.document.getElementById('actionbar'), null, self);
                    } catch(e) { }
				}
			</script>

            <% If SCSearchAny Then %>
			<link REL="stylesheet" TYPE="text/css" HREF="../../css/mc_filtergrid.css">
            <% End If %>

            <% If ReportStyleCSS = "" Then %>
			<link REL="stylesheet" TYPE="text/css" HREF="../../css/mc_report.css">
			<%
		    Response.Write "<style>"
		    Response.Write ReportStyleUDF
		    Response.Write "</style>"
            Else
			    Response.Write "<style>"
			    Response.Write ReportStyleCSS
			    Response.Write ReportStyleUDF
			    Response.Write "</style>"
			End If %>
			<link REL="stylesheet" TYPE="text/css" HREF="../../css/mc_report_edit.css">
		<%
		Else
			outputreportcssforexport()
			'OutputStandAloneJavascript()
		End If
	End If

	If wtf Then
		outputreportcss()
		'OutputStandAloneJavascript()
	End If

	If PO_REPORT_LINEHEIGHT <> "SMALL" Then
	  rw "<style type=""text/css"">"
	  rw "  .blank_row {"
	  If PO_REPORT_LINEHEIGHT = "MEDIUM" Then
	    rw "    height:24px"
	  ElseIf PO_REPORT_LINEHEIGHT = "LARGE" Then
	    rw "    height:30px"
	  End If
	  rw "  }"
	  rw "</style>"
	End If
    rw "<style type=""text/css"">"
    rw "@media print "
    rw "{"
    rw "    .no-print, .no-print * "
    rw "    {"
    rw "        display: none !important;"
    rw "    }"
    rw "}"
    rw "</style>"
	rw "</HEAD>"
End Sub

Sub OutputWOHeaderRight(htype)
	Dim uploadDate, dateNumber

	Select Case htype
	
	Case "WO"
		rw "<div style=""font-family:Arial;font-size:16px;color:#333333;font-weight:bold;margin-bottom:4px;"">"
        If WOStatus = "OPEN" Then
            rw "ESTIMATE"
        Else
            rw "INVOICE"
        End If
        rw "</div>"
		rw "<div style=""font-family:Arial;font-size:16px;color:#333333;font-weight:bold;margin-bottom:4px;"">" & _
		   NullCheck(RS_WO("WOID")) & _
			"</div>"
		rw "<div style=""font-family:Arial;font-size:11px;font-weight:normal"">" & _
			NullCheck(RS_WO("RepairCenterName")) & _
			"</div>"
        rw "<div style=""font-family:Arial;font-size:11px;color:#333333;font-weight:bold;margin-bottom:4px;"">"
        If WOStatus = "OPEN" Then
            rw "Upload Date: "
        Else
            uploadDate = RS_WO("Closed")
            dateNumber = Weekday(uploadDate)
            If dateNumber = 1 Then
                 rw "Upload Date: " & DateNullCheck(DateAdd("d",1,uploadDate))
            ElseIf dateNumber = 2 Then
                rw "Upload Date: " & DateNullCheck(uploadDate)
            Else
               dateNumber = 9 - dateNumber
               rw "Upload Date: " & DateNullCheck(DateAdd("d",dateNumber,uploadDate))
            End If
        End If
        rw "</div>"
		'If ((WO_REPORT_HEADER_SHOWSTATUS = "Yes") Or (WO_REPORT_HEADER_SHOWSUBSTATUS = "Yes")) Then
		'  rw "<div style=""font-family:Arial;font-size:11px;font-weight:normal"">" 
		'  If WO_REPORT_HEADER_SHOWSTATUS = "Yes" Then
		'    If NullCheck(RS_WO("WOStatus")) <> "" Then
		'      rw NullCheck(RS_WO("WOStatus")) 
		'      If WO_REPORT_HEADER_SHOWSUBSTATUS = "Yes" Then
		'        If NullCheck(RS_WO("SubStatusDesc")) <> "" Then
		'          rw " (" & NullCheck(RS_WO("SubStatusDesc")) & ")"
		'        End If
		'      End If	
		'    End If	    
		'  Else
		'    If WO_REPORT_HEADER_SHOWSUBSTATUS = "Yes" Then
		'      If NullCheck(RS_WO("SubStatusDesc")) <> "" Then
		'        rw NullCheck(RS_WO("SubStatusDesc"))
		'      End If
		'    End If
		'  End If
		'  rw "</div>"
		'End If
		'If WO_REPORT_HEADER_SHOWSUBSTATUS = "Yes" Then
		'  rw "<div style=""font-family:Arial;font-size:10px;font-weight:normal"">" & NullCheck(RS_WO("SubStatusDesc")) & "</div>"
		'End If
		'If reporthasfields Then
		'	Call OutputStatusControl()
		'Else
		'	rw_fileonly "<div style=""font-family:Arial;font-size:11px;font-weight:normal"">" & _
		'	   "Sent " & CStr(Date()) & "&nbsp;-&nbsp;" & CStr(TimeNullCheckAT(Time())) & "&nbsp;"
		'		If RS_WO("PrintedBox") Then
		'		rw_fileonly "(Duplicate Copy)"
		'		 End If
		'		rw_fileonly "</div>"
		'	If Not FromAgent Then				
		'		Response.Write "<div style=""font-family:Arial;font-size:11px;font-weight:normal"">" & _
		'		   "Printed " & CStr(Date()) & "&nbsp;-&nbsp;" & CStr(TimeNullCheckAT(Time())) & "&nbsp;"
		'			If RS_WO("PrintedBox") Then
		'			Response.Write "(Duplicate Copy)"
		'			End If
		'			Response.Write "</div>"			
		'	End If
		'End If
	
	Case "WOGROUP"

		rw "<div style=""font-family:Arial;font-size:16px;color:#333333;font-weight:bold;margin-bottom:4px;"">" & _
		   "Work Order Group " & NullCheck(RS_WO("WOGroupPK")) & _
		   "</div>"
		rw "<div style=""font-family:Arial;font-size:11px;font-weight:normal;"">" & _
		   NullCheck(RS_WO("RepairCenterName")) & _
		   "</div>"		   		
		If reporthasfields Then		      
			Call OutputStatusControlGroup("Close All Work Order in Group")			
		Else
			rw_fileonly "<div style=""font-family:Arial;font-size:11px;font-weight:normal"">" & _
			   "Sent " & CStr(Date()) & "&nbsp;-&nbsp;" & CStr(TimeNullCheckAT(Time())) & "&nbsp;" & _
			   "</div>"
			If Not FromAgent Then				
				Response.Write "<div style=""font-family:Arial;font-size:11px;font-weight:normal"">" & _
				   "Printed " & CStr(Date()) & "&nbsp;-&nbsp;" & CStr(TimeNullCheckAT(Time())) & "&nbsp;" & _
				   "</div>"
			End If
		End If

	Case "PROJECT"

		rw "<div style=""font-family:Arial;font-size:16px;color:#333333;font-weight:bold;margin-bottom:4px;"">" & _
		   "Project " & NullCheck(RS_WO("ProjectID")) & _
		   "</div>"
		rw "<div style=""font-family:Arial;font-size:11px;font-weight:normal;"">"
		If (Not NullCheck(RS_WO("PJSupervisorName")) = "") Then
		rw "Supervisor: " & NullCheck(RS_WO("PJSupervisorName"))
		End If
		If (Not NullCheck(RS_WO("RepairCenterName")) = "") or (Not NullCheck(RS_WO("PJSupervisorName")) = "") Then
		rw " / "
		End If
		If (Not NullCheck(RS_WO("RepairCenterName")) = "") Then
		rw NullCheck(RS_WO("RepairCenterName")) 
		End If
		rw "</div>"		   
		If reporthasfields Then
			Call OutputStatusControlGroup("Close All Work Order in Project")			
		Else
			rw_fileonly "<div style=""font-family:Arial;font-size:11px;font-weight:normal"">" & _
			   "Sent " & CStr(Date()) & "&nbsp;-&nbsp;" & CStr(TimeNullCheckAT(Time())) & "&nbsp;" & _
			   "</div>"
			If Not FromAgent Then				
				Response.Write "<div style=""font-family:Arial;font-size:11px;font-weight:normal"">" & _
				   "Printed " & CStr(Date()) & "&nbsp;-&nbsp;" & CStr(TimeNullCheckAT(Time())) & "&nbsp;" & _
				   "</div>"	
			End If
		End If
				
	End Select

End Sub	

Sub OutputDocumentsBoxCustom(rs,nowocheck)   'Moved function from Core common page to this report added WOCCS check for third tab
		
	Dim DocCounter
	If NOT wostate = "WOCCS" Then    	
		Exit Sub					
	End If							

	If Not WO_DOCUMENTSECTION Then
		Exit Sub
	End If
	
	If Not rs.EOF Then 
		If NullCheck(rs("WOPK")) = NullCheck(WOPK) or nowocheck Then
		rw "<fieldset style=""padding-top:14px"">"
			If nowocheck Then
				'rw "<legend class=""legendHeader"">Documents (for all Work Orders in Group " & NullCheck(WOGroupPK) & ")</legend>"
				rw "<legend class=""legendHeader"">Documents(Summary)</legend>"
			Else
				rw "<legend class=""legendHeader"">Documents</legend>"
			End If
			rw "<table style=""margin-top:5px;"" border=""0"" cellspacing=""3"" cellpadding=""0"" width=""98%"" align=""center"">"

				Call OutputDocumentsHeader()

				DocCounter = 0

				Do While Not rs.eof and (NullCheck(rs("WOPK")) = NullCheck(WOPK) or nowocheck)
					Call OutputDocuments(rs)
					rs.MoveNext
					DocCounter = DocCounter + 1
				Loop
				
				' This is to make sure we output any documentation				
				If DocCounter > 0 Then
					rs.Move (DocCounter * -1)
				End If

			rw "</table><br>"
		rw "</fieldset>"
		End If 
	End If 
End Sub
							
%>