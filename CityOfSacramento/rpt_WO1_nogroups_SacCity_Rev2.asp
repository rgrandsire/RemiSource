<%@ EnableSessionState=False Language=VBScript %>
<% Option Explicit %>
<!--#INCLUDE FILE="../common/mc_all.asp" -->
<!--#INCLUDE FILE="includes/mcReport_common.asp" -->
<!--#INCLUDE FILE="includes/rpt_WO_common.asp" -->
<%
'WO87403: Add comments 10/2016 to the labor box

'Response.Write("URL QueryString: <br>" & Request.QueryString)

Dim WOPK, WOsql,AssetPK, RS_Parent, WellSump, IsCalibration, RS_PNR
Dim RS_WO, RS_WOassign, RS_Asset, RS_WOtask, RS_WOpart, RS_WOmiscCost, RS_WOTool, RS_WODocument, RS_WOPref, RS_L5WO, RS_EQ, RS_PM, RS_METER, RS_Spec, RS_Spec_Test
Dim AssetName, IconFile, NoIconFile_Location, NoIconFile_Asset, skipfirstrow, AssetOutput, reporthasasset
Dim CalibratedRange, Accuracy, CalibrationSchedule, RS_Spec1, RS_Spec2, RS_Spec3
reporthasasset = false
IsCalibration = false
Call SetPrintedFlag()
Call SetupWOBarcode()
Call SetupWOData()
If InWoModule Then
	Call GetRecordStatusForClickedWO()
End If
Call DoOutput()
Call CloseDown()

Sub DoOutput()
	
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

		If Not RS_WO.EOF Then		
		
			Do While Not RS_WO.EOF
				'set work order PK
				WOPK = RS_WO("WOPK")
				AssetPK = RS_WO("AssetPK")
		    If Not IsNull(AssetPK) Then
	        Call GetReportAsset 
	        reporthasasset = true
		    End IF
        If NullCheck(RS_WO("ProcedureID")) = "CAL-A" Then
          IsCalibration = true
        End If
				'rw IsCalibration
				rw "<table border=""0"" width=""100%"">"			
					rw "<tr>"
						rw "<td valign=""top"">"
							Call OutputLogoOrName()
						rw "</td>"
						rw "<td align=""right"" valign=""bottom"">"
							Call OutputBarCode(WO_Barcode_WOID,WO_BarcodeFormat_WOID,NullCheck(RS_WO("WOID")),"White")
						rw "</td>"
						rw "<td valign=""top"" align=""right"">"
							Call OutputWOHeaderRight("WO")
						rw "</td>"
					rw "</tr>"
				rw "</table>"

		        If WellSump = true then	
                rw "<table align=""center"" border=""0"" cellpadding='1' cellspacing='1' width=""180"" bgColor=""#F5F5F5"" style=""border:1px solid"" height=""40"">"
				    rw "<tr>"
					    rw "<td valign=""middle"" align=""center"" style=""padding-left:10;padding-bottom:2;padding-top:2; font-size:18pt;"" class=""data""><b>" & NullCheck(RS_Parent("AssetName")) & "</b></td>"		        
				    rw "</tr>"
			    rw "</table>"				
		            
		        End If	

				' Output Main Details, Tasks, Labor, Materials, Tools, Other Costs,
				' & Documents for each Work Order
				' ================================================================										
				If reporthasfields Then
					Call OutputReportBox()
				End If
				
        If IsCalibration then
          rw "<table width=""100%"" style=""border: 1px solid grey;""><tr style=""border: 1px solid grey;""><td style=""FONT-WEIGHT: bold; FONT-SIZE: 16px; COLOR: royalblue; FONT-FAMILY: Arial"" align=""center"">Calibration Record Form</td></tr>"
          rw "<tr><td>"
          OutputCalibrationInstrumentInfo
          rw "</td></tr><tr><td>"
          OutputCalibrationTaskBox
          rw "</td></tr><tr><td>"
          OutputCalibrationPNR
          rw "</td></tr>"
          rw "</table>"
		      rw "<div style=""padding-top:5; padding-left:5; font-family:arial; font-size:10pt; color:gray; font-weight:bold;"" align=""center"">"
		      rw "Calibration Record Form - Doc. No. 1044, " & NullCheck(RS_WO("Revised"))
		      rw "</div>"
				Else
				
				  OutputMainDetailsBox
          'BEGIN***SacCity***********************************************************************************
          If reporthasasset Then
    			'Call OutputAssetDetailsBox() ' New Section for Equipment Details and Customer Info
            If RS_Spec_Test("EqSpecCount") > 0 Then
				      OutputAssetSpecificationsBox() ' New Section for Equipment Specifications
            End If    				
    	    End If
          'END***SacCity*************************************************************************************				
				  OutputTaskBox RS_WOTask
				  OutputLaborBoxCustom RS_WOAssign, False   'WO87403: Add comments 10/2016
				  OutputMaterialsToolsBox RS_WOpart, RS_WOtool, False
				  OutputOtherCostsBox RS_WOmiscCost, False
          'BEGIN***SacCity***********************************************************************************
          If reporthasasset Then
    				  OutputLastFiveWOBox							
    	    End IF
          'END***SacCity*************************************************************************************				
				  OutputDocumentsBox RS_WOdocument, False
				  If Not reporthasfields Then
					  Call OutputReportBox()
				  End If
				  OutputDocumentText RS_WOdocument,False, "WO"
			  
        End If
				RS_WO.MoveNext

				' do not output a page break on the last record
				If Not RS_WO.EOF Then
					rw "<P style='page-break-before: always'>"
				End If
			Loop
		Else
			rw "<div style=""padding-top:5; padding-left:5; font-family:arial; font-size:10pt; color:gray; font-weight:bold;"">"
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
	CloseObj RS_L5WO
	CloseObj RS_EQ
	CloseObj RS_PM
	CloseObj RS_METER
	CloseObj RS_Spec
	CloseObj RS_Spec_Test
	CloseObj RS_Spec1
	CloseObj RS_Spec2
	CloseObj RS_Spec3
	CloseObj RS_PNR

End Sub

Sub SetupWOData()

	Dim RecordType

	If wostate = "WOC" or wostate = "CC" Then
		RecordType = "2"
	Else
		RecordType = "1"
	End If

	If errormessage = "" Then

		If InStr(UCase(sql_where),"LEFT OUTER JOIN ") > 0 or InStr(UCase(sql_where),"INNER JOIN ") > 0 Then
			WOsql = "SELECT DISTINCT WO.WOPK FROM WO WITH (NOLOCK) LEFT OUTER JOIN Asset WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON AssetHierarchy.AssetPK = WO.AssetPK " & sql_where
		Else		
			WOsql = "SELECT WOPK FROM WO WITH (NOLOCK) LEFT OUTER JOIN Asset WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK LEFT OUTER JOIN AssetHierarchy WITH (NOLOCK) ON AssetHierarchy.AssetPK = WO.AssetPK " & sql_where
		End If

		'Response.Write "<textarea id=textarea1 name=textarea1>" & WOsql & "</textarea>"
		'Response.End
		
		' GET WORK ORDER
		If InStr(UCase(sql_where),"LEFT OUTER JOIN ") > 0 or InStr(UCase(sql_where),"INNER JOIN ") > 0 Then
			sql = "SELECT DISTINCT WO.*, Asset.IsLocation, Asset.Meter1UnitsDesc, Asset.Meter2UnitsDesc, InstructionsAsset = Case When Asset.InstructionsToWO = 1 Then Asset.Instructions Else Null End, Revised = ProcedureLibrary.UDFChar1 FROM WO WITH (NOLOCK) LEFT OUTER JOIN Asset WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK LEFT OUTER JOIN ProcedureLibrary WITH (NOLOCK) ON ProcedureLibrary.ProcedurePK = WO.ProcedurePK " & sql_where & " ORDER BY WO.WOPK"
		Else
			sql = "SELECT WO.*, Asset.IsLocation, Asset.Meter1UnitsDesc, Asset.Meter2UnitsDesc, InstructionsAsset = Case When Asset.InstructionsToWO = 1 Then Asset.Instructions Else Null End, Revised = ProcedureLibrary.UDFChar1 FROM WO WITH (NOLOCK) LEFT OUTER JOIN Asset WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK LEFT OUTER JOIN ProcedureLibrary WITH (NOLOCK) ON ProcedureLibrary.ProcedurePK = WO.ProcedurePK " & sql_where & " ORDER BY WO.WOPK"
		End If

		'Response.Write "<textarea id=textarea1 name=textarea1>" & sql & "</textarea>"
		'Response.End
		'Call OutputInTextArea(sql,"ReportSQLOutput",True)

		Set RS_WO = db.runSQLReturnRS(sql,"")
	
		If Trim(Request.QueryString("sqlwhere")) = "" Then
			Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	
		Else
			Call dok_check_afterflush_noinfo(db,"Report Message","You have chosen to run a report that is not compatible with the current module criteria. You can either choose a different report, or you can run the selected report from the Maintenance Reporter applicaiton by clicking the Reports button on the toolbar.")		
		End If

		If wostate = "WOC" Then
			' GET ACTUAL LABOR		WO87403: Add comments 10/2016	
			sql = "SELECT     WOlabor.PK, WOlabor.WOPK, WOlabor.LaborPK, WOlabor.LaborID, WOlabor.LaborName, WOLabor.Comments, WOlabor.EstimatedHours, WOlabor.RegularHours, WOlabor.OvertimeHours, WOlabor.OtherHours, WOlabor.WorkDate, WOlabor.TimeIn, " & _
				  "			  WOlabor.TimeOut, WOlabor.AccountID, WOlabor.AccountName, WOlabor.CategoryID, WOlabor.CategoryName, WOlabor.TotalCost, " & _ 
				  "			  WOlabor.TotalCharge, WOlabor.CostRegular, WOlabor.CostOvertime, WOlabor.CostOther, WOlabor.ChargeRate, WOlabor.ChargePercentage, WOlabor.RowVersionDate, Labor.Photo " & _
				  "FROM WOlabor WITH (NOLOCK) INNER JOIN WO WITH (NOLOCK) ON WO.WOPK = WOlabor.WOPK LEFT OUTER JOIN " & _
				  "			  Labor WITH (NOLOCK) ON WOlabor.LaborPK = Labor.LaborPK INNER JOIN " & _
				  "           LaborTypes lt WITH (NOLOCK) ON lt.LaborType = Labor.LaborType " & _
				  "WHERE (WOlabor.RecordType = " & RecordType & ") "
				  If Not sql_where = "" Then
					sql = sql & "AND (WOlabor.WOPK in (" & WOsql & ")) "
				  End If
				  sql = sql & _
				  "ORDER BY WOlabor.WOPK, Labor.LaborName"				  		
		Else								
			' GET ASSIGNED LABOR
			sql = "SELECT WOassign.PK, WOassign.WOPK, WOassign.IsAssigned, Labor.LaborID, " +_
			      "Labor.LaborName, Labor.LaborType, WOassign.AssignedHours, WOassign.AssignedDate " +_
				  "FROM WOassign WITH (NOLOCK) " +_
				  "INNER JOIN Labor WITH (NOLOCK) ON Labor.LaborPK = WOassign.LaborPK "
				  If Not sql_where = "" Then
					sql = sql & "WHERE (WOassign.WOPK in (" & WOsql & ")) AND (WOassign.Active = 1 or WOassign.Active Is Null) "
				  Else
				    sql = sql & "WHERE (WOassign.Active = 1 or WOassign.Active Is Null) "
				  End If
				  sql = sql & _
				  "ORDER BY WOassign.WOPK, WOassign.IsAssigned Desc, Labor.LaborName"
		End If
		
		Set RS_WOAssign = db.runSQLReturnRS(sql,"")
		Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

			
		' Get Work Order Tasks
		sql = "SELECT PK, WOPK, TaskNo, TaskAction = CASE WHEN WOTask.AssetPK Is Not Null Then '<b>'+Asset.AssetName + ' [' + Asset.AssetID + ']</b> ' + CASE WHEN Asset.Vicinity Is Not Null AND Asset.Vicinity <> '' Then RTrim(Asset.Vicinity) + ': ' Else '' END + WOtask.TaskAction Else WOtask.TaskAction END, Rate, Measurement, MeasurementInitial, WOTask.Comments, Initials, Fail, Complete, Header, LineStyle, ToolName  " +_
			  "FROM WOtask WITH (NOLOCK) LEFT OUTER JOIN Asset WITH (NOLOCK) ON Asset.AssetPK = WOTask.AssetPK LEFT OUTER JOIN Tool  WITH (NOLOCK) ON Tool.ToolPK = WOTask.ToolPK "
			  If Not sql_where = "" Then
				sql = sql & "WHERE (WOPK in (" & WOsql & ")) "
			  End If
			  sql = sql & _
			  "ORDER BY WOPK, TaskNo"
    Call OutputInTextArea(sql,"ReportSQLOutput",True)
		'Response.Write "<textarea cols=100 rows=6>" & sql & "</textarea>"
		'Response.End
    
		Set RS_WOtask = db.runSQLReturnRS(sql,"")
		Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

    sql = "SELECT PK, WOPK, TaskNo, TaskAction = CASE WHEN WOTask.AssetPK Is Not Null Then '<b>'+Asset.AssetName + ' [' + Asset.AssetID + ']</b> ' + CASE WHEN Asset.Vicinity Is Not Null AND Asset.Vicinity <> '' Then RTrim(Asset.Vicinity) + ': ' Else '' END + WOtask.TaskAction Else WOtask.TaskAction END, Rate, Measurement, MeasurementInitial, WOTask.Comments, Initials, Fail, Complete, Header, LineStyle, ToolName  " +_
			  "FROM WOtask WITH (NOLOCK) LEFT OUTER JOIN Asset WITH (NOLOCK) ON Asset.AssetPK = WOTask.AssetPK LEFT OUTER JOIN Tool  WITH (NOLOCK) ON Tool.ToolPK = WOTask.ToolPK "
			  If Not sql_where = "" Then
				sql = sql & "WHERE (WOPK in (" & WOsql & ")) AND TaskNo IN (340,350) "
				Else
				sql = sql & sql_where & " AND TaskNo IN (340,350) "
			  End If
			  sql = sql & _
			  "ORDER BY WOPK, TaskNo"
		Set RS_PNR = db.runSQLReturnRS(sql,"")
		Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	
		
		' Get Work Order Material
		sql = "SELECT WOpart.PK, WOpart.WOPK, WOpart.PartID, WOpart.PartName, WOpart.LocationID, WOpart.QuantityEstimated, WOpart.QuantityActual, WOpart.OtherCost, WOpart.AccountID, Part.PartDescription, PartLocation.Bin, PartLocation.Lot  " +_
			  "FROM WOpart WITH (NOLOCK) " +_
			  "LEFT OUTER JOIN Part WITH (NOLOCK) ON WOpart.PartPK = Part.PartPK " +_
			  "LEFT OUTER JOIN PartLocation WITH (NOLOCK) ON WOpart.PartPK = PartLocation.PartPK and WOpart.LocationPK = PartLocation.LocationPK " +_
			  "WHERE (WOpart.RecordType = " & RecordType & ") "
			  If Not sql_where = "" Then
				sql = sql & "AND (WOpart.WOPK in (" & WOsql & ")) "
			  End If
			  sql = sql & _
			  "ORDER BY WOpart.WOPK, WOpart.LocationID, WOpart.PartID"
		
		Set RS_WOpart = db.runSQLReturnRS(sql,"")
		Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

		' Get Work Order Other Costs
		sql = "SELECT PK, WOPK, MiscCostName, MiscCostDesc, WOmiscCost.InvoiceNumber, MiscCostDate, AccountID, AccountName, QuantityEstimated, EstimatedCost, ActualCost " +_
			  "FROM WOmiscCost WITH (NOLOCK) " +_
			  "WHERE (RecordType = " & RecordType & ") "
			  If Not sql_where = "" Then
				sql = sql & "AND (WOPK in (" & WOsql & ")) "
			  End If
			  sql = sql & _
			  "ORDER BY WOPK, MiscCostName"
		Set RS_WOmiscCost = db.runSQLReturnRS(sql,"")
		Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

		' Get Work Order Tools
		sql = "SELECT WOtool.WOPK, WOtool.ToolID, WOtool.ToolName, WOtool.LocationID, WOtool.LocationName, WOtool.QuantityEstimated " +_
			  "FROM WOtool WITH (NOLOCK) LEFT OUTER JOIN " +_
	          "Tool WITH (NOLOCK) ON WOtool.ToolPK = Tool.ToolPK " +_
			  "WHERE (RecordType = 1) "
			  If Not sql_where = "" Then
				sql = sql & "AND (WOPK in (" & WOsql & ")) "
			  End If
			  sql = sql & _
			  "ORDER BY WOtool.WOPK, WOtool.LocationName, WOtool.ToolName"

		Set RS_WOtool = db.runSQLReturnRS(sql,"")
		Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

		' Get Work Order Document Attachments */
		
		sql = _
		"SELECT     WO.WOPK, md.PK, d.LocationType, d.DocumentID, d.DocumentName, md.ModuleID, d.DocumentTypeDesc, " & _
		"                      d.Location, md.PrintWithWO, md.SendWithEmail, md.RowVersionDate, d.Photo, " & _
		"                      MCModule.TitleforDocumentList, d.DocumentText " & _
		"FROM         AssetDocument md WITH (NOLOCK) LEFT OUTER JOIN " & _
		"                      Document d WITH (NOLOCK) ON md.DocumentPK = d.DocumentPK INNER JOIN " & _
		"                      MCModule WITH (NOLOCK) ON md.ModuleID = MCModule.ModuleID INNER JOIN " & _
		"                      WO WITH (NOLOCK) ON WO.AssetPK = md.AssetPK " & _
		"WHERE "
		If Not sql_where = "" Then
		  sql = sql & "(WO.WOPK in (" & WOsql & ")) AND "
		End If
		sql = sql & _			  			  			  			  		
		"(md.PrintWithWO = 1 OR md.SendWithEmail = 1 OR d.LocationType = 'LIBRARY') " & _
		"UNION ALL " & _
		"SELECT     WO.WOPK, md.PK, d.LocationType, d.DocumentID, d.DocumentName, md.ModuleID, d.DocumentTypeDesc, " & _
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
		"(md.PrintWithWO = 1 OR md.SendWithEmail = 1 OR d.LocationType = 'LIBRARY') " & _
		"UNION ALL " & _		
		"SELECT     WO.WOPK, md.PK, d.LocationType, d.DocumentID, d.DocumentName, md.ModuleID, d.DocumentTypeDesc, " & _ 
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
		"(md.PrintWithWO = 1 OR md.SendWithEmail = 1 OR d.LocationType = 'LIBRARY') " & _
		"UNION ALL " & _
		"SELECT     WO.WOPK, md.PK, d.LocationType, d.DocumentID, d.DocumentName, md.ModuleID, d.DocumentTypeDesc, " & _ 
		"         d.Location, md.PrintWithWO, md.SendWithEmail, md.RowVersionDate, d.Photo, " & _
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
		"(md.PrintWithWO = 1 OR md.SendWithEmail = 1 OR d.LocationType = 'LIBRARY') " & _
		"UNION ALL " & _
		"SELECT     WO.WOPK, md.PK, d.LocationType, d.DocumentID, d.DocumentName, md.ModuleID, d.DocumentTypeDesc, " & _
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
		"(md.PrintWithWO = 1 OR md.SendWithEmail = 1 OR d.LocationType = 'LIBRARY') " & _
		"ORDER BY WO.WOPK, md.ModuleID, d.DocumentID "
		
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

Sub GetReportAsset
' *************************************************************************
  'Get Asset Hierarchy
  Dim RS_AssetTempValue
  RS_AssetTempValue = NullCheck(AssetPK)
  If RS_AssetTempValue = "" Then
      RS_AssetTempValue = "-1"
  End If

  sql = "SELECT Asset.Icon, Asset.AssetPK, Asset.AssetID, Asset.AssetName, Asset.IsLocation, Asset.IsUp " +_
    "FROM Asset WITH (NOLOCK) INNER JOIN MC_GetAssetParentPK('" & RS_AssetTempValue & "') b On b.AssetPK = Asset.AssetPK " &_
    "ORDER BY b.lvl DESC"

  Set RS_Asset = db.runSQLReturnRS(sql,"")
  Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

  ' Get Work Order Equipment Details
  sql = "SELECT Asset.* FROM Asset WITH (NOLOCK) WHERE AssetPK = " & AssetPK & " "
  'Response.Write "<textarea id=textarea1 name=textarea1 cols=100 rows=6>" & sql & "</textarea>"
  'Response.End

  Set RS_EQ = db.runSQLReturnRS(sql,"")
  Call dok_check_afterflush(db,"Report Message","EQ: There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	


  ' Get Parent
  sql = "SELECT AssetName FROM Asset WITH (NOLOCK) WHERE AssetPK = (SELECT ParentPK FROM AssetHierarchy  WITH (NOLOCK) WHERE AssetPK = " & RS_AssetTempValue & ")"
  'Response.Write "<textarea id=textarea1 name=textarea1 cols=100 rows=6>" & sql & "</textarea>"
  'Response.End
    
  Set RS_Parent = db.runSQLReturnRS(sql,"")
  Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	
      
  If NullCheck(RS_Parent("AssetName")) <> "" Then
      If ((MID(RS_Parent("AssetName"),1,4)  = "Well") OR (MID(RS_Parent("AssetName"),1,4) = "Sump") OR (MID(RS_Parent("AssetName"),1,4) = "WELL") OR (MID(RS_Parent("AssetName"),1,4) = "SUMP")) Then
          WellSump = true
      Else
          WellSump = false
      End If        
  Else
      WellSump = false
  End If
  'Response.Write MID(RS_Parent("AssetName"),1,4)
  'Response.Write WellSump
  'Response.End

  ' Test for Asset Specifications
  Dim sqlspec
  sqlspec = "Select Count(*) AS EqSpecCount FROM Asset WITH (NOLOCK) " +_
      "INNER JOIN WO WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " +_
      "INNER JOIN AssetSpecification WITH (NOLOCK) ON Asset.AssetPK = AssetSpecification.AssetPK "
  sqlspec = sqlspec & Request("sqlwhere") 

  Set RS_Spec_Test = db.runSQLReturnRS(sqlspec,"")
  Call dok_check_afterflush(db,"Report Message","EQSpecTest: There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

  ' Get next PM Date
  sql = "Select PM.NextScheduledDate FROM Asset WITH (NOLOCK) " +_
          "INNER JOIN WO WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " +_
          "INNER JOIN PMAsset WITH (NOLOCK) ON PMAsset.AssetPK = Asset.AssetPK " +_
          "INNER JOIN PM WITH (NOLOCK) ON PM.PMPK = PMAsset.PMPK " +_
          "INNER JOIN PMProcedure WITH (NOLOCK) ON PMProcedure.PMPK = PM.PMPK "
  sql = sql & "where (WO.WOPK in (" & WOsql & ")) "

  Set RS_PM = db.runSQLReturnRS(sql,"")
  Call dok_check_afterflush(db,"Report Message","EQPM: There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

  ' Get Work Order Asset Specification
  sql = "Select AssetSpecification.* FROM Asset WITH (NOLOCK) " +_
	    "INNER JOIN WO WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " +_
	    "INNER JOIN AssetSpecification WITH (NOLOCK) ON Asset.AssetPK = AssetSpecification.AssetPK " 
  sql = sql & "where (WO.WOPK = (" & WOPK & ")) "

  Set RS_Spec = db.runSQLReturnRS(sql,"")
  Call dok_check_afterflush(db,"Report Message","EqSpec: There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	


  ' Get Work Order Asset Specification Calibrated Range
  sql = "Select AssetSpecification.ValueText FROM Asset WITH (NOLOCK) " +_
	    "INNER JOIN WO WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " +_
	    "INNER JOIN AssetSpecification WITH (NOLOCK) ON Asset.AssetPK = AssetSpecification.AssetPK " 
  sql = sql & "where (WO.WOPK = (" & WOPK & ")) AND SpecificationName = 'Calibrated Range'"

  Set RS_Spec1 = db.runSQLReturnRS(sql,"")
  Call dok_check_afterflush(db,"Report Message","EqSpec: There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

  If Not RS_Spec1.EOF Then
    CalibratedRange = RS_Spec1("ValueText")
  End If  

  ' Get Work Order Asset Specification ASccuracy
  sql = "Select AssetSpecification.ValueText FROM Asset WITH (NOLOCK) " +_
	    "INNER JOIN WO WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " +_
	    "INNER JOIN AssetSpecification WITH (NOLOCK) ON Asset.AssetPK = AssetSpecification.AssetPK " 
  sql = sql & "where (WO.WOPK = (" & WOPK & ")) AND SpecificationName = 'Accuracy'"

  Set RS_Spec2 = db.runSQLReturnRS(sql,"")
  Call dok_check_afterflush(db,"Report Message","EqSpec: There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

  If Not RS_Spec2.EOF Then
    Accuracy = RS_Spec2("ValueText")
  End If  

  ' Get Work Order Asset Specification Calibration Schedule
  sql = "Select AssetSpecification.ValueText FROM Asset WITH (NOLOCK) " +_
	    "INNER JOIN WO WITH (NOLOCK) ON Asset.AssetPK = WO.AssetPK " +_
	    "INNER JOIN AssetSpecification WITH (NOLOCK) ON Asset.AssetPK = AssetSpecification.AssetPK " 
  sql = sql & "where (WO.WOPK = (" & WOPK & ")) AND SpecificationName = 'Calibration Schedule'"

  Set RS_Spec3 = db.runSQLReturnRS(sql,"")
  Call dok_check_afterflush(db,"Report Message","EqSpec: There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

  If Not RS_Spec3.EOF Then
    CalibrationSchedule = RS_Spec3("ValueText")
  End If  


  ' GET last 5 WORK ORDERs
  sql = "SELECT TOP 5 WOID, Reason, StatusDesc, TargetDate, LaborReport FROM WO WITH (NOLOCK) "_
          & "WHERE AssetPK = " & AssetPK & " AND WOPK <> " & WOPK & " ORDER BY WO.Closed DESC "
  'Response.Write "<textarea id=textarea1 name=textarea1 cols=100 rows=6>" & sql & "</textarea>"
  'Response.End
  Set RS_L5WO = db.runSQLReturnRS(sql,"")

  Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

  ' Get Meter History Details
  sql = "Select top 1 AssetMeterHistory.* " +_
	    "FROM AssetMeterHistory WITH (NOLOCK) INNER JOIN " +_
	    "WO WITH (NOLOCK) ON AssetMeterHistory.AssetPK = WO.AssetPK " 
  	 
	    If Not sql_where = "" Then
		  sql = sql & "where (WO.WOPK = (" & WOPK & ")) "
	    End If
	    sql = sql & _
	    "ORDER BY ReadingDate Desc "

  Set RS_METER = db.runSQLReturnRS(sql,"")
  Call dok_check_afterflush(db,"Report Message","Meter: There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")	

' *************************************************************************

End Sub

Sub OutputAssetDetailsBox()
	rw "<table cellpadding=""0"" cellspacing=""0"" width=""100%"" border=""0"" style=""font-family:Arial;font-size:12px;color:#333333"">"
		rw "<tr>"
		    rw "<td valign=""top"" width=""100%"">"		
	        rw "<fieldset style=""padding-top:5"">"
		        rw "<legend class=""legendHeader"">Asset Details</legend>"
    		        Call OutputAssetDetails()
	        rw "</fieldset>"				
	        rw "</td>"				
        rw "</tr>"
    rw "</table>"
End Sub
Sub OutputAssetDetails()
	rw "<table cellpadding=""3"" width=""100%"" height="""" border=""0"" style=""font-family:Arial;font-size:12px;color:#333333"">"
		rw "<tr>"
		rw "<td valign=""top"" width=""100%"">"		
			rw "<table border=""0"" cellspacing=""1"" cellpadding=""1"">"        
    			rw "<tr>"
					rw "<td class=""labels"" nowrap valign=""top"">Asset Name/ID:</td>"
					rw "<td class=""data"" valign=""top"">"
					if NOT RS_EQ.EOF then
						rw NullCheck(RS_EQ("AssetName")) & "&nbsp;(" & NullCheck(RS_EQ("AssetID")) & ")"
					end if
					rw "</td>"
				rw "</tr>"
' BEGIN 07/17/2006 Addition --> Asset Address
                IF 	NullCheck(RS_EQ("Address")) <> "" then
			    rw "<tr>"
				    rw "<td class=""labels"" nowrap valign=""top"">Street</td>"
				    rw "<td class=""data"" valign=""top"">"
					    rw NullCheck(RS_EQ("Address")) 
				    rw "</td>"
			    rw "</tr>"  
			    If ((NullCheck(RS_EQ("City")) <> "") AND (NullCheck(RS_EQ("State")) <> "") AND (NullCheck(RS_EQ("ZIP")) <> "")) Then          		
			    rw "<tr>"
				    rw "<td class=""labels"" nowrap valign=""top"">City/State/Zip:</td>"
				    rw "<td class=""data"">"
				        If NullCheck(RS_EQ("City")) <> "" Then
				            rw NullCheck(RS_EQ("City"))& "&nbsp;" 
				        End If
				        If NullCheck(RS_EQ("State")) <> "" Then
    				        rw NullCheck(RS_EQ("State"))& ",&nbsp;" 
    				    End If
    				    If NullCheck(RS_EQ("Zip")) <> "" Then
    				        rw NullCheck(RS_EQ("Zip"))
    				    End If
				    rw "</td>"
			    rw "</tr>"
			    End If
	            End IF
' END 07/17/2006 Addition							
    			rw "<tr>"
					rw "<td class=""labels"" nowrap valign=""top"">Model #:</td>"
					rw "<td class=""data"" valign=""top"">"
					if NOT RS_EQ.EOF then
						rw NullCheck(RS_EQ("Model")) & "&nbsp;"
					end if
					rw "</td>"
				rw "</tr>"
    			rw "<tr>"
					rw "<td class=""labels"" nowrap valign=""top"">Serial #:</td>"
					rw "<td class=""data"" valign=""top"">"
					if NOT RS_EQ.EOF then
						rw NullCheck(RS_EQ("Serial")) & "&nbsp;"
					end if
					rw "</td>"
				rw "</tr>"
    			rw "<tr>"
					rw "<td class=""labels"" nowrap valign=""top"">Vicinity:</td>"
					rw "<td class=""data"" valign=""top"">"
					if NOT RS_EQ.EOF then
						rw NullCheck(RS_EQ("Vicinity")) & "&nbsp;"
					end if
					rw "</td>"
				rw "</tr>"
				if NOT RS_EQ.EOF then
                    if NOT RS_PM.EOF then
                        if NOT IsNull(RS_PM("NextScheduledDate")) or RS_PM("NextScheduledDate") <> "" then
                            rw "<tr>"
					            rw "<td class=""labels"" nowrap valign=""top"">Next PM Date:</td>"
					            rw "<td class=""data"" valign=""top"">"
						            rw NullCheck(RS_PM("NextScheduledDate")) & "&nbsp;"
					            rw "</td>"
				            rw "</tr>"
				        end if
				    End If
				    if NOT IsNull(RS_EQ("InstallDate")) or RS_EQ("InstallDate") <> "" then
    			        rw "<tr>"
					        rw "<td class=""labels"" nowrap valign=""top"">Install Date:</td>"
					        rw "<td class=""data"" valign=""top"">"
						        rw NullCheck(RS_EQ("InstallDate")) & "&nbsp;"
					        rw "</td>"
				        rw "</tr>"
                    end if
                    if NOT IsNull(RS_EQ("WarrantyExpire")) or RS_EQ("WarrantyExpire") <> "" then
    			    rw "<tr>"
					    rw "<td class=""labels"" nowrap valign=""top"">Warranty Expires:</td>"
					    rw "<td class=""data"" valign=""top"">"
						    rw NullCheck(RS_EQ("WarrantyExpire")) & "&nbsp;"
					    rw "</td>"
				    rw "</tr>"
				    end if
                    IF 	RS_EQ("IsMeter") = true then
    			    rw "<tr>"
					    rw "<td class=""labels"" nowrap valign=""top"">Last Reading:</td>"
					    rw "<td class=""data"" valign=""top"">"
						    rw NullCheck(RS_METER("Meter1Reading")) & "&nbsp;on&nbsp;" & NullCheck(RS_Meter("ReadingDate"))
					    rw "</td>"
				    rw "</tr>"            		
    			    rw "<tr>"
					    rw "<td class=""labels"" nowrap valign=""top"">Current Reading:</td>"
					    rw "<td class=""data_underline"" align=""center"">&nbsp;</td>"
				    rw "</tr>"
		            End IF
	            End If
	        rw "</table>"			
	    rw "</td>"	
		rw "</tr>"					
	rw "</table>"
End Sub

Sub OutputAssetSpecificationsBox()
    rw "<fieldset style=""padding-top:5"">"
        rw "<legend class=""legendHeader"">Equipment Specifications</legend>"
        rw "<table width='98%' align='center'>" 
        Call OutputAssetSpecificationsHeader()
        Call OutputAssetSpecifications(RS_Spec)
        rw "</table>"
    rw "</fieldset>"
End Sub

Sub OutputAssetSpecificationsHeader()
	rw "<tr>"
		rw "<td class=""labels"" align=""left"">Specification</td>"
		rw "<td class=""labels"" align=""center"" width=""55"">#</td>"
		rw "<td class=""labels"" align=""center"" width=""55"">High</td>"
		rw "<td class=""labels"" align=""center"" width=""55"">Low</td>"
		rw "<td class=""labels"" align=""center"" width=""100"">Text</td>"
		rw "<td class=""labels"" align=""center"" width=""80"">Date</td>"
	rw "</tr>"		
End Sub

Sub OutputAssetSpecifications(RS_Spec)
    Do while not RS_Spec.EOF
	rw "<tr>"
		    rw "<td valign=""bottom"" class=""data_underline"" align=""left"">" & NullCheck(RS_Spec("SpecificationName")) & "&nbsp;</td>"
			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">" 
			if IsNull(RS_Spec("ValueNumeric")) or RS_Spec("ValueNumeric") = "" then
			    response.Write "&#8722 &nbsp;</td>"
			else
			    Response.Write NullCheck(RS_Spec("ValueNumeric")) & "&nbsp;</td>"
			end if 
			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">"
			if IsNull(RS_Spec("ValueHi")) or RS_Spec("ValueHi") = "" then
			    Response.Write "&#8722 &nbsp;</td>"
			else
			    Response.Write NullCheck(RS_Spec("ValueHi")) & "&nbsp;</td>"
			end if
			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">"
			if IsNull(RS_Spec("ValueLow")) or RS_Spec("ValueLow") = "" then
			    Response.Write "&#8722 &nbsp;</td>"
			else
			    Response.Write NullCheck(RS_Spec("ValueLow")) & "&nbsp;</td>"
			end if
			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">"
			if IsNull(RS_Spec("ValueText")) or RS_Spec("ValueText") = "" then
			    Response.Write "&#8722 &nbsp;</td>"
			else
			    Response.Write NullCheck(RS_Spec("ValueText")) & "&nbsp;</td>"
			end if
			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">"
			if IsNull(RS_Spec("ValueDate")) or RS_Spec("ValueDate") = "" then
			    Response.Write "&#8722 &nbsp;</td>"
			else
			    Response.Write NullCheck(RS_Spec("ValueDate")) & "&nbsp;</td>"
			end if
	rw "</tr>"
	RS_Spec.MoveNext
	loop
	rw "<tr height='4px'><td colspan='6'></td></tr>"
End Sub

Sub OutputLastFiveWOBox

    rw "<fieldset style=""padding-top:5"">"
        rw "<legend class=""legendHeader"">Last 5 Issued Work Orders</legend>"
        rw "<table width='98%' align='center'>" 
    IF RS_L5WO.EOF THEN
        Response.Write "<tr><td class=""labels"">No Other Work Orders Found!</td></tr>"
    ELSE
        Call OutputLastFiveWOHeader()
        Call OutputLastFiveWO(RS_L5WO)
    END IF
        rw "</table>"
    rw "</fieldset>"

End Sub

Sub OutputLastFiveWOHeader()
	rw "<tr>"
		rw "<td class=""labels"" align=""left"" width=""55"">WO #</td>"
		rw "<td class=""labels"" align=""left"">Reason</td>"
		rw "<td class=""labels"" align=""center"" width=""55"">Status</td>"
		rw "<td class=""labels"" align=""center"" width=""55"">Target Date</td>"
		rw "<td class=""labels"" align=""left"">Labor Report</td>"
	rw "</tr>"		
End Sub

Sub OutputLastFiveWO(RS_L5WO)
    Do while not RS_L5WO.EOF
	    rw "<tr>"
		        rw "<td valign=""bottom"" class=""data_underline"" align=""left"">" & NullCheck(RS_L5WO("WOID")) & "&nbsp;</td>"
			    rw "<td valign=""bottom"" class=""data_underline"" align=""left"">" & NullCheck(RS_L5WO("Reason")) & "&nbsp;</td>"
			    rw "<td valign=""bottom"" class=""data_underline"" align=""center"">" & NullCheck(RS_L5WO("StatusDesc")) & "&nbsp;</td>"
			    rw "<td valign=""bottom"" class=""data_underline"" align=""center"">" & NullCheck(RS_L5WO("TargetDate")) & "&nbsp;</td>"
			    rw "<td valign=""bottom"" class=""data_underline"" align=""left"">" & NullCheck(RS_L5WO("LaborReport")) & "&nbsp;</td>"
	    rw "</tr>"
	    RS_L5WO.MoveNext
	loop
	rw "<tr height='4px'><td colspan='5'></td></tr>"
End Sub

Sub OutputReportBox()

	'If Not WO_REPORTSECTION and Not wostate = "WOC" Then
	'	Exit Sub
	'End If

	rw "<fieldset style=""padding-top:10; margin-bottom:7;"">"
		rw "<legend class=""legendHeader"">Labor Report</legend>"
		Call OutputWOReport()
		rw "<br>"
	rw "</fieldset>"

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
			rw "<td class=""data"" width=""50%"">"
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
						rw "<td class=""data_underline"" width=""100%"">" & NullCheckNBSP(RS_WO("Complete")) & "</td>"						
						End Select						
					rw "</tr>"
				rw "</table>"
			rw "</td>"
			rw "<td class=""data"" width=""50%"">"
				rw "<table border=""0"" cellspacing=""3"" cellpadding=""0"">"
					rw "<tr>"
						rw "<td nowrap class=""labels"" valign=""bottom"">Failure:&nbsp;</td>"
						Select Case wostate
						Case "WO"						
						rw "<td class=""data_underline"" width=""100%"">&nbsp;</td>"
						Case "CC"
						rw "<td class=""data_underline"" width=""100%"">"
							rw "<input class=""normal"" mcType=""C"" maxlength=""25"" mcRequired=""N"" type=""text"" name=""WO_Failure_" & RS_WO("WOPK") & """ value=""" & NullCheck(RS_WO("FailureID")) & """ size=""10"" onChange=""top.dovalid('FA',this,'WO');"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.checkKey(this,self);""><img src=""../../images/lookupiconxp3fk.gif"" border=""0"" align=""absbottom"" onclick=""top.dolookup('FA',WO_Failure_" & RS_WO("WOPK") & ",'WO')"" class=""lookupicon"" width=""16"" height=""20"">"
							rw "<span style=""display:none;"" id=""WO_Failure_" & RS_WO("WOPK") & "Desc"" class=""mc_lookupdesc"">" & NullCheck(RS_WO("FailureName")) & "</span>"
							rw "<input type=""hidden"" name=""WO_Failure_" & RS_WO("WOPK") & "PK"" class=""mc_pluggedvalue"">"				
						rw "</td>"												
						Case "WOC"
						rw "<td class=""data_underline"" width=""100%"">" & NullCheckNBSP(RS_WO("FailureID")) & " / " & NullCheckNBSP(RS_WO("FailureName")) & "</td>"						
						End Select						
					rw "</tr>"
				rw "</table>"
			rw "</td>"
		rw "</tr>"

		rw "<tr>"
			rw "<td colspan=""2"" class=""labels"" valign=""top"">"
				rw "<table style=""margin-top:10px;"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
					If wostate = "CC" Then
						rw "<tr>"
						rw "<td colspan=""3"" width=""100%"" align=""right"">"
						rw "<img border=""0"" src=""../../images/button_addarrow.gif"" onclick=""top.showpopup('actions','Actions',266,100,this,WO_LaborReport_" & RS_WO("WOPK") & ",self)"" WIDTH=""80"" HEIGHT=""15"">"
						rw "</td>"
						rw "</tr>"
					End If
					rw "<tr>"
					rw "<td valign=""top"" style=""padding-left:3;"" colspan=""2"" nowrap class=""labels"">"
						rw "Report:&nbsp;&nbsp;"
					rw "</td>"
					Select Case wostate
					Case "WO"						
					Case "CC"
					rw "<td width=""100%"">"						
					rw "<textarea mcType=""C"" name=""WO_LaborReport_" & RS_WO("WOPK") & """ wrap=""hard"" style=""width: 100%; height: 35;"" class=""normal"" onfocus=""top.fieldfocus(this);"" onkeypress=""return top.checkKey(this,self);"" onblur=""top.fieldblur(this);"" onChange=""top.fieldvalid(this);"" rows=""1"" cols=""20"">" & Replace(NullCheckNBSP(RS_WO("LaborReport")),"%0D%0A",chr(13)+chr(10)) & "</textarea>"
					rw "</td>"
					Case "WOC"
					rw "<td class=""data_underline"" width=""100%"">" & Replace(NullCheckNBSP(RS_WO("LaborReport")),"%0D%0A","<br>") & "</td>"						
					End Select											
					rw "</tr>"																		
					If wostate = "CC" Then
					rw "<tr>"
					rw "<td colspan=""4"">"

					rw "<table cellspacing=""0"" cellpadding=""0"" width=""100%"" class=""normaltext"" style=""margin-top:10px;"">"
					rw "<tr>"
					
					rw "<td>Set Tasks: Completed / Failed</td>"
					rw "<td>Set Actuals = Estimates: All / Labor / Materials / Other Costs</td>"
					rw "<td>Set Labor Hours...</td>"
					
					rw "</tr>"
					rw "</table>"
					
					rw "</td>"
					rw "</tr>"
					End If
				rw "</table>"						
			rw "</td>"					
		rw "</tr>"
		rw "<tr>"
			rw "<td colspan=""2"" class=""labels"" valign=""top"">&nbsp;</td>"					
		rw "</tr>"
		rw "<tr>"
			rw "<td colspan=""2"" class=""labels"" valign=""top"">&nbsp;</td>"					
		rw "</tr>"
		rw "<tr>"
			rw "<td colspan=""2"" class=""labels"" valign=""top"">&nbsp;</td>"					
		rw "</tr>"
		rw "<tr>"
			rw "<td colspan=""2"" class=""labels"" valign=""top"">&nbsp;</td>"					
		rw "</tr>"


		
	rw "</table>"
End Sub

Sub OutputWOHeaderRight(htype)
		'Response.Write WellSump
		'Response.End
		
	Select Case htype
	
	Case "WO"
		
		rw "<div style=""font-family:Arial;font-size:16px;color:#333333;font-weight:bold;margin-bottom:4;"">" & _
		   "Work Order " & NullCheck(RS_WO("WOID")) & _
			"</div>"
			
		rw "<div style=""font-family:Arial;font-size:11px;font-weight:normal"">" & _
			NullCheck(RS_WO("RepairCenterName")) & _
			"</div>"
		If reporthasfields Then
			Call OutputStatusControl()
		Else
			rw_fileonly "<div style=""font-family:Arial;font-size:11px;font-weight:normal"">" & _
			   "Sent " & CStr(Date()) & "&nbsp;-&nbsp;" & CStr(TimeNullCheckAT(Time())) & "&nbsp;"
				If RS_WO("PrintedBox") Then
				rw_fileonly "(Duplicate Copy)"
				 End If
				rw_fileonly "</div>"
			If Not FromAgent Then				
				Response.Write "<div style=""font-family:Arial;font-size:11px;font-weight:normal"">" & _
				   "Printed " & CStr(Date()) & "&nbsp;-&nbsp;" & CStr(TimeNullCheckAT(Time())) & "&nbsp;"
					If RS_WO("PrintedBox") Then
					Response.Write "(Duplicate Copy)"
					End If
					Response.Write "</div>"			
			End If
		End If
	
	Case "WOGROUP"

		rw "<div style=""font-family:Arial;font-size:16px;color:#333333;font-weight:bold;margin-bottom:4;"">" & _
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

		rw "<div style=""font-family:Arial;font-size:16px;color:#333333;font-weight:bold;margin-bottom:4;"">" & _
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

Sub OutputTaskBox(rs)

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
	
	  'If IsCalibration then
		'  rw "<table style=""margin-top:5px;"" border=""1"" cellspacing=""0"" cellpadding=""0"" width=""98%"" align=""center"" style=""border-collapse;"">"
  	'	  rw "<tr><td colspan=3 style=""FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: royalblue; FONT-FAMILY: Arial;"">Calibration Instrument Data</td></tr>"
    '    rw "<tr><td>"
    'Else	  
      rw "<fieldset style=""padding-top:14"">"
  		rw "<legend class=""legendHeader"">Tasks</legend>"
    'End If

		rw "<table style=""margin-top:5px;"" border=""0"" cellspacing=""3"" cellpadding=""0"" width=""98%"" align=""center"">"

			Call OutputTasksHeader() 
			' task data 	
			If Not rs.EOF Then
                If CheckIfFieldExists(RS_WOTask,"Comments") Then
				    Do While NullCheck(rs("WOPK")) = NullCheck(WOPK)
					    Call OutputTasksWithComments(rs)		
					    rs.MoveNext
				    Loop
				Else
				    Do While NullCheck(rs("WOPK")) = NullCheck(WOPK)
					    Call OutputTasks(rs)		
					    rs.MoveNext
				    Loop
				End If
			End If
			
      If IsCalibration then
        Call OutputBlankTaskRow(BlankRowNum,8)
      Else
  			Call OutputBlankTaskRow(BlankRowNum,7)
      End If

		rw "</table>"
		rw "<br>"	
  If IsCalibration then
    rw "</td></tr></table>"
  Else
  	rw "</fieldset>"
  End If

End Sub

Sub OutputTasksHeader()
	rw "<tr>"
		rw "<td class=""labels"" width=""40"" style=""padding-left:5"">#</td>"
		rw "<td class=""labels"">Description</td>"
		rw "<td class=""labels"" align=""center"" width=""55"">Rating</td>"
  'If IsCalibration then
	'	rw "<td class=""labels"" align=""center"" width=""55"">Initial Measure</td>"
	'	rw "<td class=""labels"" align=""center"" width=""55"">Final Measure</td>"
  'Else
		rw "<td class=""labels"" align=""center"" width=""55"">Meas.</td>"
  'End If
		rw "<td class=""labels"" align=""center"" width=""70"">Initials</td>"
		rw "<td class=""labels"" align=""right"" width=""35"">Failed</td>"
		rw "<td class=""labels"" align=""center"" width=""35"">&nbsp;Complete&nbsp;</td>"
	rw "</tr>"		
End Sub

Sub OutputTasks(RS_WOTask)
	rw "<tr>"
		If NullCheck(RS_WOTask("Header")) Then		
		rw "<td valign=""top"" class=""data"" colspan=""8"" style=""padding-top:8""><b>" & NullCheck(RS_WOTask("TaskAction")) & "</b></td>"
		Else
		rw "<td valign=""top"" class=""data_underline"">" & NullCheck(RS_WOTask("TaskNo")) & "&nbsp;</td>"
		Dim LineStyle
		Select Case NullCheck(RS_WOTask("LineStyle"))
			Case "I1"
				LineStyle=" style=""padding-left:10;"""
			Case "I2"
				LineStyle=" style=""padding-left:20;"""
			Case "I3"
				LineStyle=" style=""padding-left:30;"""
			Case "B"
				LineStyle=" style=""font-weight:bold;"""
			Case "BR"
				LineStyle=" style=""font-weight:bold;color:red;"""
			Case "I"
				LineStyle=" style=""font-style:italic;"""
			Case Else
				LineStyle=""
		End Select		
		rw "<td" & LineStyle & " valign=""top"" class=""data_underline"">" & Replace(NullCheck(RS_WOTask("TaskAction")),vbCrLf,"<br/>") & "&nbsp;</td>"
        
		Select Case wostate

		Case "WO","WOC"
			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">" & NullCheck(RS_WOTask("Rate")) & "&nbsp;</td>"
      'If IsCalibration then
			'  rw "<td valign=""bottom"" class=""data_underline"" align=""center"">&nbsp;</td>"
			'  rw "<td valign=""bottom"" class=""data_underline"" align=""center"">&nbsp;</td>"
      'Else
			  rw "<td valign=""bottom"" class=""data_underline"" align=""center"">" & NullCheck(RS_WOTask("Measurement")) & "&nbsp;</td>"
      'End If
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
			'If IsCalibration then
			'rw "<td valign=""bottom"" class=""data_underline"" align=""center"">"
			'rw "<input name=""TA_Measurement_" & RS_WOTask("PK") & """ value="""" class=""normalright"" mcType=""N"" maxlength=""12"" size=""1"" type=""text"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);"">"
			'rw "</td>"
			'rw "<td valign=""bottom"" class=""data_underline"" align=""center"">"
			'rw "<input name=""TA_Measurement_" & RS_WOTask("PK") & """ value="""" class=""normalright"" mcType=""N"" maxlength=""12"" size=""1"" type=""text"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);"">"
			'rw "</td>"
      'Else
			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">"
			rw "<input name=""TA_Measurement_" & RS_WOTask("PK") & """ value=""" & NullCheck(RS_WOTask("Measurement")) & """ class=""normalright"" mcType=""N"" maxlength=""12"" size=""1"" type=""text"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);"">"
			rw "</td>"
      'End If
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

Sub OutputTasksWithComments(RS_WOTask)
	rw "<tr>"
		If NullCheck(RS_WOTask("Header")) Then		
		rw "<td valign=""top"" class=""data"" colspan=""8"" style=""padding-top:8""><b>" & NullCheck(RS_WOTask("TaskAction")) & "</b></td>"
		Else
		rw "<td valign=""top"" class=""data_underline"">" & NullCheck(RS_WOTask("TaskNo")) & "&nbsp;</td>"
		Dim LineStyle
		Select Case NullCheck(RS_WOTask("LineStyle"))
			Case "I1"
				LineStyle=" style=""padding-left:10;"""
			Case "I2"
				LineStyle=" style=""padding-left:20;"""
			Case "I3"
				LineStyle=" style=""padding-left:30;"""
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
      'If IsCalibration then
			'  rw "<td valign=""bottom"" class=""data_underline"" align=""center"">&nbsp;</td>"
			'  rw "<td valign=""bottom"" class=""data_underline"" align=""center"">&nbsp;</td>"
      'Else
			  rw "<td valign=""bottom"" class=""data_underline"" align=""center"">" & NullCheck(RS_WOTask("Measurement")) & "&nbsp;</td>"
      'End If
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
			'If IsCalibration then
			'rw "<td valign=""bottom"" class=""data_underline"" align=""center"">"
			'rw "<input name=""TA_Measurement_" & RS_WOTask("PK") & """ value="""" class=""normalright"" mcType=""N"" maxlength=""12"" size=""1"" type=""text"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);"">"
			'rw "</td>"
			'rw "<td valign=""bottom"" class=""data_underline"" align=""center"">"
			'rw "<input name=""TA_Measurement_" & RS_WOTask("PK") & """ value="""" class=""normalright"" mcType=""N"" maxlength=""12"" size=""1"" type=""text"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);"">"
			'rw "</td>"
      'Else
			rw "<td valign=""bottom"" class=""data_underline"" align=""center"">"
			rw "<input name=""TA_Measurement_" & RS_WOTask("PK") & """ value=""" & NullCheck(RS_WOTask("Measurement")) & """ class=""normalright"" mcType=""N"" maxlength=""12"" size=""1"" type=""text"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);"">"
			rw "</td>"
      'End If
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

Sub OutputCalibrationInstrumentInfo
  'Do While NullCheck(RS_WOTask("WOPK")) = NullCheck(WOPK)
		rw "<table style=""margin-top:5px;"" border=""1"" cellspacing=""0"" cellpadding=""0"" width=""98%"" align=""center"" style=""border-collapse;"">"
  		rw "<tr><td colspan=3 style=""FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: royalblue; FONT-FAMILY: Arial;"">&nbsp;Calibration Instrument Information</td></tr>"
      rw "<tr><td valign=""top"">"
      
          rw "<table width=""100%"" border=0 cellspacing=5 cellpadding=0>"
            rw "<tr>"
              rw "<td class=""labels"">Manufacturer:</td>"
              rw "<td class=""data"">"&Nullcheck(RS_EQ("ManufacturerID"))&"</td>"
            rw "</tr>"
            rw "<tr>"
              rw "<td class=""labels"">Model No.:</td>"
              rw "<td class=""data"">"&Nullcheck(RS_EQ("Model"))&"</td>"
            rw "</tr>"
            rw "<tr>"
              rw "<td class=""labels"">Serial No.:</td>"
              rw "<td class=""data"">"&Nullcheck(RS_EQ("Serial"))&"</td>"
            rw "</tr>"
            rw "<tr>"
              rw "<td class=""labels"">Calibrated Range:</td>"
              rw "<td class=""data"">"&CalibratedRange&"</td>"
            rw "</tr>"
            rw "<tr>"
              rw "<td class=""labels"">Accuracy:</td>"
              rw "<td class=""data"">"&Accuracy&"</td>"
            rw "</tr>"
          rw" </table>"
          
        rw "</td>"
        rw "<td valign=""top"">"
        
          rw "<table width=""100%"" border=0 cellspacing=5 cellpadding=0>"
            rw "<tr>"
              rw "<td class=""labels"">Tag. No.:</td>"
              rw "<td class=""data"">"&Nullcheck(RS_EQ("AssetID"))&"</td>"
            rw "</tr>"
            rw "<tr>"
              rw "<td class=""labels"">Instrument Type/Description:</td>"
              rw "<td class=""data"">"&Nullcheck(RS_EQ("AssetName"))&"</td>"
            rw "</tr>"
            rw "<tr>"
              rw "<td class=""labels"">Original As Found Reading:</td>"
              rw "<td class=""data"" style=""border: 1px solid grey;"">"
              sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 90"
              Set rs = db.runSQLReturnRS(sql,"") 
              If Not rs.EOF Then
		            rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
              Else
                rw "&nbsp;"
		          End If   
              rw "</td>"
            rw "</tr>"
            rw "<tr>"
              rw "<td class=""labels"">Original As Left Reading:</td>"
              rw "<td class=""data"" style=""border: 1px solid grey;"">"
              sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 95"
              Set rs = db.runSQLReturnRS(sql,"") 
              If Not rs.EOF Then
		            rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
              Else
                rw "&nbsp;"
		          End If   
		          rw "</td>"
            rw "</tr>"
            rw "<tr>"
              rw "<td class=""labels"">Calibration Schedule:</td>"
              rw "<td class=""data"">"&CalibrationSchedule&"</td>"
            rw "</tr>"
          rw "</table>"
          
        rw "</td>"
        rw "<td valign=""top"">"
        
          rw "<table width=""100%"" border=0 cellspacing=5 cellpadding=0>"
            rw "<tr>"
              rw "<td class=""labels"" valign=""top"">Area:</td>"
            rw "</tr>"
            rw "<tr>"
              rw "<td class=""data"">"              
					      ' asset hierarchy								
					      NoIconFile_Location = "images/icons/facility_g.gif"
					      NoIconFile_Asset = "images/icons/gearsxp_g.gif"
					      skipfirstrow = True											
					      If Not RS_Asset.EOF Then
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
      												
						      AssetOutput="<img src='" & GetSession("webHTTP") & GetWebServer() & Application("Web_Path") & Application("mapp_path") & IconFile & "' border=0 style='margin-right:4'>" & AssetName & "<br>"
      																								
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
							      End If
							      rw "<td class=""" & stclass & """ valign=""top"">"
							      rw AssetOutput
							      rw "</td>"
							      rw "</tr>"									
						      End If											
					      Loop 
					      End If
					      rw "</table>"
              rw "</td>"
            rw "</tr>"
          rw "</table>"

        rw "</td>"
      rw "</tr>"
		rw "</table>"
    'RS_WOTask.MoveNext
  'Loop
End Sub

Sub OutputCalibrationPNR
	rw "<table style=""margin-top:5px;"" border=""1"" cellspacing=""0"" cellpadding=""0"" width=""98%"" align=""center"" style=""border-collapse;"">"
		rw "<tr><td colspan=3 style=""FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: royalblue; FONT-FAMILY: Arial;"">&nbsp;Performance And Review</td></tr>"
    rw "<tr><td>"
		rw "<table style=""margin-top:5px;"" border=""0"" cellspacing=""8"" cellpadding=""0"" width=""98%"" align=""center"">"
      rw "<tr>"
        rw "<td nowrap class=""labels"" valign=""bottom"">Performed By:&nbsp;</td>"
        rw "<td class=""data_underline"" width=""70%"">&nbsp;</td>"
        rw "<td nowrap class=""labels"" valign=""bottom"">Date:&nbsp;</td>"
        rw "<td class=""data_underline"" width=""30%"">"
        sql = "SELECT MeasurementInitial, Comments FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 340"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
        Else
          rw "&nbsp;"
        End If        
        rw "&nbsp;</td>"
      rw "</tr>"
      rw "<tr>"
        rw "<td colspan=4 class=""data"">"
        'sql = "SELECT Comments FROM WOTask WHERE WOPK = " & WOPK & " AND TaskNo = 340"
        'Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("Comments")) & "&nbsp;"
        Else
          rw "&nbsp;"
        End If
		    rw "</td>"	          
      rw "</tr>"
      rw "<tr>"
        rw "<td nowrap class=""labels"" valign=""bottom"">Reviewed By:&nbsp;</td>"
        rw "<td class=""data_underline"" width=""70%"">&nbsp;</td>"
        rw "<td nowrap class=""labels"" valign=""bottom"">Date:&nbsp;</td>"
        rw "<td class=""data_underline"" width=""30%"">"
        sql = "SELECT MeasurementInitial, Comments FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 350"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
        Else
          rw "&nbsp;"
        End If                
        rw "</td>"
      rw "</tr>"
      rw "<tr>"
        rw "<td colspan=4 class=""data"">"
        'sql = "SELECT Comments FROM WOTask WHERE WOPK = " & WOPK & " AND TaskNo = 350"
        'Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("Comments")) & "&nbsp;"
        Else
          rw "&nbsp;"
        End If
		    rw "</td>"	          
      rw "</tr>"      
		rw "</table>"
		rw "<br>"	
	rw "</td></tr></table>"
End Sub

Sub OutputCalibrationTaskBox
  rw "<table style=""margin-top:5px;"" border=""1"" cellspacing=""0"" cellpadding=""3"" width=""98%"" align=""center"" style=""border-collapse;"">"
    rw "<tr><td colspan=3 style=""FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: royalblue; FONT-FAMILY: Arial;"">Calibration Instrument Data</td></tr>"
    rw "<tr><td>"

		rw "<table style=""margin-top:5px;"" border=""0"" cellspacing=""3"" cellpadding="""" width=""100%"" align=""center"">"
		
		  rw "<tr>"
		    rw "<td class=""labels"" valign=""top"" width=""11%"">Standard Input</td><td class=""data"" valign=""top"" width=""11%"" style=""border: 1px solid grey;"">"
        sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 100"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
        Else
          rw "&nbsp;"
        End If
		    rw "</td>"
		    rw "<td class=""labels"" valign=""top"" width=""11%"">Instrument ID</td><td class=""data"" width=""11%"" valign=""top"" style=""border: 1px solid grey;"">"&Nullcheck(RS_EQ("Vicinity"))&"</td>"
		    rw "<td colspan=2 width=""22%"">&nbsp;</td>"
		    rw "<td class=""labels"" rowspan=14 valign=""top"" width=""34%"" style=""border: 1px solid grey;"">"
		      rw "<table width=""100%"">"
		        rw "<tr>"
		          rw "<td class=""labels"">Comments</td>"
		        rw "</tr>"
		        rw "<tr>"
		          rw "<td class=""data"">"
              'sql = "SELECT LaborReport FROM WO WHERE WOPK = " & WOPK 
              'Set rs = db.runSQLReturnRS(sql,"") 
              'If Not rs.EOF Then
                rw NullCheck(RS_WO("LaborReport")) & "&nbsp;"
              'Else
              '  rw "&nbsp;"
              'End If
		          rw "</td>"
		        rw "</tr>"
		        rw "<tr>"
		          rw "<td>&nbsp</td>"
		        rw "</tr>"
		      rw "</table>"
		    rw "</td>"
      rw "</tr>"
      
		  rw "<tr>"
		    rw "<td class=""labels"" valign=""top"">Actual Input Units IN:</td><td class=""data"" valign=""top"" style=""border: 1px solid grey;"">"
        sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 101"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
        Else
          rw "&nbsp;"
        End If
		    rw "</td>"
		    rw "<td class=""labels"" valign=""top"">Units:</td><td class=""data"" valign=""top"" style=""border: 1px solid grey;"">"
        sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 102"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
        Else
          rw "&nbsp;"
        End If
		    rw "</td>"
		    rw "<td colspan=2>&nbsp;</td>"
      rw "</tr>"
      
      rw "<tr>"
        rw "<td class=""labels"" align=""center"">Test Point</td>"
        rw "<td class=""labels"" align=""center"">Actual Input</td>"
        rw "<td class=""labels"" align=""center"">Desired</td>"
        rw "<td class=""labels"" align=""center"">As found</td>"
        rw "<td class=""labels"" align=""center"">As Left</td>"
        rw "<td class=""labels"" align=""center"">Error</td>"
      rw "</tr>"

      ' Tasks
      ' Test Point 1
      rw "<tr>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">1</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 105"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If	        
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 110"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If	              
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT Measurement, MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 115"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If	         
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        'sql = "SELECT Measurement FROM WOTask WHERE WOPK = " & WOPK & " AND TaskNo = 115"
        'Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("Measurement")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If	         
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT Measurement FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 120"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("Measurement")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If	         
        rw "</td>"
      rw "</tr>"
      ' Test Point 2
      rw "<tr>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">2</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 125"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If	                
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 130"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT Measurement, MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 135"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If           
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        'sql = "SELECT Measurement FROM WOTask WHERE WOPK = " & WOPK & " AND TaskNo = 135"
        'Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("Measurement")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If	        
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT Measurement FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 140"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("Measurement")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If        
        rw "</td>"
      rw "</tr>"
      ' Test Point 3
      rw "<tr>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">3</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 145"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 150"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If        
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT Measurement, MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 155"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If        
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        'sql = "SELECT Measurement FROM WOTask WHERE WOPK = " & WOPK & " AND TaskNo = 155"
        'Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("Measurement")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If        
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT Measurement FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 160"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("Measurement")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If        
        rw "</td>"
      rw "</tr>"
      ' Test Point 4
      rw "<tr>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">4</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 165"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If           
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 170"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If        
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT Measurement, MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 175"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If        
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        'sql = "SELECT Measurement FROM WOTask WHERE WOPK = " & WOPK & " AND TaskNo = 175"
        'Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("Measurement")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If        
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT Measurement FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 180"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("Measurement")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If        
        rw "</td>"
      rw "</tr>"
      ' Test Point 5
      rw "<tr>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">5</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 185"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If        
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 190"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If    
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT Measurement, MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 195"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If       
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        'sql = "SELECT Measurement FROM WOTask WHERE WOPK = " & WOPK & " AND TaskNo = 195"
        'Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("Measurement")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If
        rw "</td>"
        rw "<td class=data align=""center"" style=""border: 1px solid grey;"">"
        sql = "SELECT Measurement FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 200"
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
          rw Nullcheck(rs("Measurement")) & "&nbsp;"
	      Else
	        rw "&nbsp;"
	      End If
        rw "</td>"
      rw "</tr>"
      
      ' Standards / Equipment Used
      rw "<tr>"
        rw "<td class=""labels"" align=""center"" colspan=4>Standards / Equipment Used</td>"
        rw "<td class=""labels"" align=""center"" colspan=2>Cal. Due Date</td>"
      rw "</tr>"
      
      rw "<tr>"
        rw "<td class=""data"" valign=""top"" colspan=4 height=""100"" style=""border: 1px solid grey;"">"
        sql = "SELECT WOTask.Comments, ToolName FROM WOTask  WITH (NOLOCK) LEFT OUTER JOIN Tool  WITH (NOLOCK) ON Tool.ToolPK = WOTask.ToolPK WHERE WOPK = " & WOPK & " AND TaskNo = 300"
        'Response.Write "<textarea cols=80 rows=6>" & sql & "</textarea>"
        'Response.End
        Set rs = db.runSQLReturnRS(sql,"") 
        If Not rs.EOF Then
	        rw "<table width=""100%"">"
	        rw "<tr height=""50""><td class=""data"" style=""border-bottom: 1px solid grey;"" valign=""top"">" & NullCheck(rs("Comments")) 
          
            
          
          
          rw "&nbsp;</td></tr>"
         
         sql = "SELECT WOTask.Comments, ToolName FROM WOTask  WITH (NOLOCK) LEFT OUTER JOIN Tool  WITH (NOLOCK) ON Tool.ToolPK = WOTask.ToolPK WHERE WOPK = " & WOPK & " AND TaskNo = 312"
        Set rs = db.runSQLReturnRS(sql,"") 
         if Not rs.EOF Then         
	        rw "<tr height=""50""><td class=""data"" style=""border-bottom: 1px solid grey;"" valign=""top"">" & NullCheck(rs("Comments"))  & "&nbsp;</td></tr>"
          End If
	        rw "</table>"
	      Else
	        rw "&nbsp;"
	      End If
        rw "</td>"
        rw "<td class=""data"" valign=""top"" colspan=2 style=""border: 1px solid grey;"">"
	        rw "<table width=""100%"">"
	        rw "<tr height=""25""><td class=""data"" style=""border-bottom: 1px solid grey;"" valign=""middle"" align=""center"">"
          sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 305"
          Set rs = db.runSQLReturnRS(sql,"") 
          If Not rs.EOF Then
            rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	        Else
	          rw "&nbsp;"
	        End If
	        rw "</td></tr>"
	        rw "<tr height=""25""><td class=""data"" style=""border-bottom: 1px solid grey;"" valign=""middle"" align=""center"">"
          sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 310"
          Set rs = db.runSQLReturnRS(sql,"") 
          If Not rs.EOF Then
            rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	        Else
	          rw "&nbsp;"
	        End If	        
	        rw "</td></tr>"
	        rw "<tr height=""25""><td class=""data"" style=""border-bottom: 1px solid grey;"" valign=""middle"" align=""center"">"
          sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 315"
          Set rs = db.runSQLReturnRS(sql,"") 
          If Not rs.EOF Then
            rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	        Else
	          rw "&nbsp;"
	        End If
	        rw "</td></tr>"
	        rw "<tr height=""25""><td class=""data"" valign=""middle"" align=""center"">"
          sql = "SELECT MeasurementInitial FROM WOTask  WITH (NOLOCK) WHERE WOPK = " & WOPK & " AND TaskNo = 320"
          Set rs = db.runSQLReturnRS(sql,"") 
          If Not rs.EOF Then
            rw Nullcheck(rs("MeasurementInitial")) & "&nbsp;"
	        Else
	          rw "&nbsp;"
	        End If
	        rw "</td></tr>"	      
	        rw "</table>"
        rw "</td>"
		rw "</table>"
  rw "</td></tr></table>"
End Sub

'''''''''''
'############################### October 2016 Added by Remi

Sub OutputLaborBoxCustom(rs,nowocheck)

	Dim BlankRowNum,LaborFormat,LaborCols
	BlankRowNum = 0	

	If Not WO_LABORSECTION and Not wostate = "WOC" Then
		Exit Sub
	End If

	If rs.Eof Then 
	  'BlankRowNum = WO_LABORSECTION_BL
	  BlankRowNum = 2
		If wostate = "WOC" Then
			Exit Sub		
		End If
		If Not WO_LABORSECTION_B Then
			Exit Sub
		End If
	End If		
	
	rw "<fieldset style=""padding-top:14px"">"
		If nowocheck Then
			'rw "<legend class=""legendHeader"">Labor Summary (for all Work Orders in Group " & NullCheck(WOGroupPK) & ")</legend>"
			rw "<legend class=""legendHeader"">Labor (Summary)</legend>"
		Else
			rw "<legend class=""legendHeader"">Labor</legend>"
		End If
		
		rw "<table style=""margin-top:5px;"" border=""0"" cellspacing=""3"" cellpadding=""0"" width=""98%"" align=""center"">"

			Select Case wostate
			
			Case "WO","CC"
				If Not rs.eof and (NullCheck(rs("WOPK")) = NullCheck(WOPK) or nowocheck) Then
					'rs.Filter = "IsAssigned = 1"
					If rs("IsAssigned") Then
						LaborFormat = "ASSIGNED"
						'If WO_REPORT_LABOR_SHOWSTARTEND = "Yes" Then
						'  If WO_REPORT_LABOR_SHOWACCOUNT = "No" Then
						'    LaborCols = 9
						'  Else
						'    LaborCols = 10
						'  End If
						'Else
						  'If WO_REPORT_LABOR_SHOWACCOUNT = "No" Then
						    LaborCols = 7
						  'Else
						  '  LaborCols = 8
						  'End If
						'End If
					Else
						'rs.Filter = ""
						'rs.MoveFirst()
						LaborFormat = "ESTIMATED"
						'If WO_REPORT_LABOR_SHOWSTARTEND = "Yes" Then
						'  If WO_REPORT_LABOR_SHOWACCOUNT = "No" Then
						'    LaborCols = 9
						'  Else
						'    LaborCols = 10
						'  End If
						'Else
						'  If WO_REPORT_LABOR_SHOWACCOUNT = "No" Then
						    LaborCols = 7
						'  Else
						'    Laborcols = 8
						'  End If
						'End If
					End If
				Else
					LaborFormat = "NONE"
					'If WO_REPORT_LABOR_SHOWSTARTEND = "Yes" Then
					'  If WO_REPORT_LABOR_SHOWACCOUNT = "No" Then
					'    LaborCols = 7
					'  Else
					'    LaborCols = 8
					'  End If
					'Else
					'  If WO_REPORT_LABOR_SHOWACCOUNT = "No" Then
					    LaborCols = 5
					'  Else
					'    Laborcols = 6
					'  End If
					'End If
					'BlankRowNum = WO_LABORSECTION_BL
					BlankRowNum = 2
				End If	
			
			Case "WOC"
				LaborFormat = "NONE"
				'If WO_REPORT_LABOR_SHOWSTARTEND = "Yes" Then
				'  If WO_REPORT_LABOR_SHOWACCOUNT = "No" Then
				'    LaborCols = 7
				'  Else
				'    LaborCols = 8
				'  End If
				'Else
				'  If WO_REPORT_LABOR_SHOWACCOUNT = "No" Then
				    LaborCols = 5
				'  Else
				'    LaborCols = 6
				'  End If
				'End If
				'BlankRowNum = WO_LABORSECTION_BL
			  BlankRowNum = 2
			End Select						
		
			Call OutputLaborHeader(LaborFormat)
						
			' labor data 	
			If Not rs.EOF Then
				Do While Not rs.eof and (NullCheck(rs("WOPK")) = NullCheck(WOPK) or nowocheck)
					If (wostate = "WO" or wostate = "CC") and LaborFormat = "ASSIGNED" Then
						If Not rs("IsAssigned") Then
							' Do not print crafts if assignments are made
						Else
							Call OutputLabor(LaborFormat,rs)
						End If
					Else
						Call OutputLaborCustom(LaborFormat,rs)
					End If
					rs.MoveNext
				Loop
			End If
			
			Call OutputBlankRow(BlankRowNum,LaborCols)

		rw "</table><br>"
	rw "</fieldset>"
End Sub



Sub OutputLaborCustom(LaborFormat,RS_WOAssign)
	rw "<tr class=""blank_row"">"

	Select Case wostate

	Case "WO"
	
		Select Case LaborFormat
			Case "ASSIGNED"
				rw "<td class=""data_underline"">"
				rw NullCheck(RS_WOAssign("LaborName"))
				rw "</td>"
				'If WO_REPORT_LABOR_SHOWACCOUNT = "Yes" Then
				'  rw "<td width=""120"" class=""data_underline"">&nbsp;</td>"
				'End If
				rw "<td width=""90"" class=""data_underline"">" & DateNullCheck(RS_WOAssign("AssignedDate")) & "&nbsp;/&nbsp;"
				rw NullCheck(RS_WOAssign("AssignedHours")) & "&nbsp;</td>"
				rw "<td width=""90"" class=""data_underline"">&nbsp;</td>"
				'If WO_REPORT_LABOR_SHOWSTARTEND = "Yes" Then
				'  rw "<td width=""90"" class=""data_underline"">&nbsp;</td>"
				'  rw "<td width=""90"" class=""data_underline"">&nbsp;</td>"
				'End If
				rw "<td class=""data_underline"" align=""center"">&nbsp;</td>"
				rw "<td class=""data_underline"" align=""center"">&nbsp;</td>"
				rw "<td class=""data_underline"" align=""center"">&nbsp;</td>"
				rw "<td class=""data_underline"" align=""center"">&nbsp;</td>"
			Case "ESTIMATED"
				rw "<td class=""data_underline"">"
				rw NullCheck(RS_WOAssign("LaborName"))
				rw "</td>"
				rw "<td class=""data_underline"" align=""center"">" & NullCheck(RS_WOAssign("AssignedHours")) & "&nbsp;</td>"
				rw "<td width=""500"" class=""data_underline"">&nbsp;</td>"
				'If WO_REPORT_LABOR_SHOWACCOUNT = "Yes" Then
				'  rw "<td width=""120"" class=""data_underline"">&nbsp;</td>"
				'End If
				rw "<td width=""90"" class=""data_underline"">&nbsp;</td>"
				'If WO_REPORT_LABOR_SHOWSTARTEND = "Yes" Then
				'  rw "<td width=""90"" class=""data_underline"">&nbsp;</td>"
				'  rw "<td width=""90"" class=""data_underline"">&nbsp;</td>"
				'End If
				rw "<td class=""data_underline"" align=""center"">&nbsp;</td>" 
				rw "<td class=""data_underline"" align=""center"">&nbsp;</td>"
				rw "<td class=""data_underline"" align=""center"">&nbsp;</td>"
				rw "<td class=""data_underline"" align=""center"">&nbsp;</td>"
		End Select
		
	Case "CC"
	
		Select Case LaborFormat
			Case "ASSIGNED"
				rw "<td class=""data_underline"">"
				rw NullCheck(RS_WOAssign("LaborName"))
				rw "</td>"
				'If WO_REPORT_LABOR_SHOWACCOUNT = "Yes" Then
				'  rw "<td nowrap width=""80"" class=""data_underline"">"
			 	'  rw "<input class=""normal"" mcType=""C"" maxlength=""25"" mcRequired=""N"" type=""text"" name=""LA_Account_" & RS_WOAssign("PK") & """ size=""6"" onChange=""top.dovalid('AC',this,'WO');"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.checkKey(this,self);""><img src=""../../images/lookupiconxp3fk.gif"" border=""0"" align=""absbottom"" onclick=""top.dolookup('AC',LA_Account_" & RS_WOAssign("PK") & ",'WO')"" class=""lookupicon"" width=""16"" height=""20"">"
				'	  rw "<span style=""display:none;"" id=""LA_Account_" & RS_WOAssign("PK") & "Desc"" class=""mc_lookupdesc""></span>"
				'	  rw "<input type=""hidden"" name=""LA_Account_" & RS_WOAssign("PK") & "PK"" class=""mc_pluggedvalue"">"				
				'  rw "</td>"
				'End If
				rw "<td nowrap width=""90"" class=""data_underline"">" & DateNullCheck(RS_WOAssign("Comments")) & "&nbsp;/&nbsp;"
				rw "<td nowrap width=""90"" class=""data_underline"">" & DateNullCheck(RS_WOAssign("AssignedDate")) & "&nbsp;/&nbsp;"
				rw "<td nowrap width=""90"" class=""data_underline"">" &NullCheck(RS_WOAssign("AssignedHours")) & "&nbsp;Hrs</td>"
				
				rw "<td width=""90"" class=""data_underline"">"
					rw "<input class=""normal"" mcType=""D"" maxlength=""10"" mcRequired=""N"" type=""text"" name=""LA_WorkDate_" & RS_WOAssign("PK") & """ size=""8"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.checkKey(this,self);""><img src=""../../images/lookupiconxp3.gif"" border=""0"" onclick=""top.showpopup('calendar','Calendar',172,160,this,LA_WorkDate_" & RS_WOAssign("PK") & ",self)"" align=""absbottom"" class=""lookupicon"" WIDTH=""16"" HEIGHT=""20"">"
					rw "<span style=""display:none;"" id=""LA_WorkDate_" & RS_WOAssign("PK") & "Err"" class=""mc_lookupdesc""></span>"
				rw "</td>"								
				'If WO_REPORT_LABOR_SHOWSTARTEND = "Yes" Then
				'  rw "<td width=""70"" class=""data_underline"">"
				'	  rw "<input class=""normal"" mcType=""T"" maxlength=""10"" mcRequired=""N"" type=""text"" name=""LA_TimeIn_" & RS_WOAssign("PK") & """ size=""4"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.checkKey(this,self);""><img src=""../../images/lookupiconxp3.gif"" border=""0"" onclick=""top.showpopup('timepopup','Select Time',267,205,this,LA_TimeIn_" & RS_WOAssign("PK") & ",self)"" align=""absbottom"" class=""lookupicon"" WIDTH=""16"" HEIGHT=""20"">"
				'	  rw "<span style=""display:none;"" id=""LA_TimeIn_" & RS_WOAssign("PK") & "Err"" class=""mc_lookupdesc""></span>"
				'  rw "</td>"								
  				
				'  rw "<td width=""70"" class=""data_underline"">"
				'	  rw "<input class=""normal"" mcType=""T"" maxlength=""10"" mcRequired=""N"" type=""text"" name=""LA_TimeOut_" & RS_WOAssign("PK") & """ size=""4"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.checkKey(this,self);""><img src=""../../images/lookupiconxp3.gif"" border=""0"" onclick=""top.showpopup('timepopup','Select Time',267,205,this,LA_TimeOut_" & RS_WOAssign("PK") & ",self)"" align=""absbottom"" class=""lookupicon"" WIDTH=""16"" HEIGHT=""20"">"
				'	  rw "<span style=""display:none;"" id=""LA_TimeOut_" & RS_WOAssign("PK") & "Err"" class=""mc_lookupdesc""></span>"
				'  rw "</td>"								
        'End If
				rw "<td width=""60"" class=""data_underline"" align=""right"">"
				rw "<input name=""LA_RegularHours_" & RS_WOAssign("PK") & """ class=""normalright"" mcType=""N"" maxlength=""12"" size=""1"" type=""text"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);""><img src=""../../images/lookupiconxp3.gif"" border=""0"" onclick=""top.showpopup('calculator','Calculator',125,100,this,LA_RegularHours_" & RS_WOAssign("PK") & ",self)"" align=""absbottom"" class=""lookupicon"" WIDTH=""16"" HEIGHT=""20"">"
				rw "</td>"
			
				rw "<td width=""60"" class=""data_underline"" align=""right"">"
				rw "<input name=""LA_OvertimeHours_" & RS_WOAssign("PK") & """ class=""normalright"" mcType=""N"" maxlength=""12"" size=""1"" type=""text"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);""><img src=""../../images/lookupiconxp3.gif"" border=""0"" onclick=""top.showpopup('calculator','Calculator',125,100,this,LA_OvertimeHours_" & RS_WOAssign("PK") & ",self)"" align=""absbottom"" class=""lookupicon"" WIDTH=""16"" HEIGHT=""20"">"
				rw "</td>"

				rw "<td width=""60"" class=""data_underline"" align=""right"">"
				rw "<input name=""LA_OtherHours_" & RS_WOAssign("PK") & """ class=""normalright"" mcType=""N"" maxlength=""12"" size=""1"" type=""text"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);""><img src=""../../images/lookupiconxp3.gif"" border=""0"" onclick=""top.showpopup('calculator','Calculator',125,100,this,LA_OtherHours_" & RS_WOAssign("PK") & ",self)"" align=""absbottom"" class=""lookupicon"" WIDTH=""16"" HEIGHT=""20"""
				rw "</td>"
			Case "ESTIMATED"
				rw "<td class=""data_underline"">"
				rw NullCheck(RS_WOAssign("LaborName"))
				rw "</td>"
				rw "<td class=""data_underline"" align=""center"">" & NullCheck(RS_WOAssign("AssignedHours")) & "&nbsp;</td>"

				rw "<td nowrap width=""70"" class=""data_underline"">"
					rw "<input class=""normal"" mcType=""C"" maxlength=""25"" mcRequired=""N"" type=""text"" name=""LA_Labor_" & RS_WOAssign("PK") & """ size=""4"" onChange=""top.dovalid('LA',this,'WO');"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.checkKey(this,self);""><img src=""../../images/lookupiconxp3fk.gif"" border=""0"" align=""absbottom"" onclick=""top.dolookup('LA',LA_Labor_" & RS_WOAssign("PK") & ",'WO')"" class=""lookupicon"" width=""16"" height=""20"">"
					rw "<span style=""display:none;"" id=""LA_Labor_" & RS_WOAssign("PK") & "Desc"" class=""mc_lookupdesc""></span>"
					rw "<input type=""hidden"" name=""LA_Labor_" & RS_WOAssign("PK") & "PK"" class=""mc_pluggedvalue"">"				
				rw "</td>"
				'If WO_REPORT_LABOR_SHOWACCOUNT = "Yes" Then
				'  rw "<td nowrap width=""70"" class=""data_underline"">"
				'	  rw "<input class=""normal"" mcType=""C"" maxlength=""25"" mcRequired=""N"" type=""text"" name=""LA_Account_" & RS_WOAssign("PK") & """ size=""4"" onChange=""top.dovalid('AC',this,'WO');"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.checkKey(this,self);""><img src=""../../images/lookupiconxp3fk.gif"" border=""0"" align=""absbottom"" onclick=""top.dolookup('AC',LA_Account_" & RS_WOAssign("PK") & ",'WO')"" class=""lookupicon"" width=""16"" height=""20"">"
				'	  rw "<span style=""display:none;"" id=""LA_Account_" & RS_WOAssign("PK") & "Desc"" class=""mc_lookupdesc""></span>"
				'	  rw "<input type=""hidden"" name=""LA_Account_" & RS_WOAssign("PK") & "PK"" class=""mc_pluggedvalue"">"				
				'  rw "</td>"
				'End If
				rw "<td width=""90"" class=""data_underline"">"
					rw "<input class=""normal"" mcType=""D"" maxlength=""10"" mcRequired=""N"" type=""text"" name=""LA_WorkDate_" & RS_WOAssign("PK") & """ size=""8"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.checkKey(this,self);""><img src=""../../images/lookupiconxp3.gif"" border=""0"" onclick=""top.showpopup('calendar','Calendar',172,160,this,LA_WorkDate_" & RS_WOAssign("PK") & ",self)"" align=""absbottom"" class=""lookupicon"" WIDTH=""16"" HEIGHT=""20"">"
					rw "<span style=""display:none;"" id=""LA_WorkDate_" & RS_WOAssign("PK") & "Err"" class=""mc_lookupdesc""></span>"
				rw "</td>"								
				'If WO_REPORT_LABOR_SHOWSTARTEND = "Yes" Then
				'  rw "<td width=""70"" class=""data_underline"">"
				'	  rw "<input class=""normal"" mcType=""T"" maxlength=""10"" mcRequired=""N"" type=""text"" name=""LA_TimeIn_" & RS_WOAssign("PK") & """ size=""4"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.checkKey(this,self);""><img src=""../../images/lookupiconxp3.gif"" border=""0"" onclick=""top.showpopup('timepopup','Select Time',267,205,this,LA_TimeIn_" & RS_WOAssign("PK") & ",self)"" align=""absbottom"" class=""lookupicon"" WIDTH=""16"" HEIGHT=""20"">"
				'	  rw "<span style=""display:none;"" id=""LA_TimeIn_" & RS_WOAssign("PK") & "Err"" class=""mc_lookupdesc""></span>"
				'  rw "</td>"								
  				
				'  rw "<td width=""70"" class=""data_underline"">"
				'	  rw "<input class=""normal"" mcType=""T"" maxlength=""10"" mcRequired=""N"" type=""text"" name=""LA_TimeOut_" & RS_WOAssign("PK") & """ size=""4"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.checkKey(this,self);""><img src=""../../images/lookupiconxp3.gif"" border=""0"" onclick=""top.showpopup('timepopup','Select Time',267,205,this,LA_TimeOut_" & RS_WOAssign("PK") & ",self)"" align=""absbottom"" class=""lookupicon"" WIDTH=""16"" HEIGHT=""20"">"
				'	  rw "<span style=""display:none;"" id=""LA_TimeOut_" & RS_WOAssign("PK") & "Err"" class=""mc_lookupdesc""></span>"
				'  rw "</td>"								
        'End If
				rw "<td width=""60"" class=""data_underline"" align=""right"">"
				rw "<input name=""LA_RegularHours_" & RS_WOAssign("PK") & """ class=""normalright"" mcType=""N"" maxlength=""12"" size=""1"" type=""text"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);""><img src=""../../images/lookupiconxp3.gif"" border=""0"" onclick=""top.showpopup('calculator','Calculator',125,100,this,LA_RegularHours_" & RS_WOAssign("PK") & ",self)"" align=""absbottom"" class=""lookupicon"" WIDTH=""16"" HEIGHT=""20"">"
				rw "</td>"
			
				rw "<td width=""60"" class=""data_underline"" align=""right"">"
				rw "<input name=""LA_OvertimeHours_" & RS_WOAssign("PK") & """ class=""normalright"" mcType=""N"" maxlength=""12"" size=""1"" type=""text"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);""><img src=""../../images/lookupiconxp3.gif"" border=""0"" onclick=""top.showpopup('calculator','Calculator',125,100,this,LA_OvertimeHours_" & RS_WOAssign("PK") & ",self)"" align=""absbottom"" class=""lookupicon"" WIDTH=""16"" HEIGHT=""20"">"
				rw "</td>"

				rw "<td width=""60"" class=""data_underline"" align=""right"">"
				rw "<input name=""LA_OtherHours_" & RS_WOAssign("PK") & """ class=""normalright"" mcType=""N"" maxlength=""12"" size=""1"" type=""text"" onChange=""top.fieldvalid(this);"" onfocus=""top.fieldfocus(this);"" onblur=""top.fieldblur(this);"" onkeypress=""return top.fnTrapAlpha(this,self);""><img src=""../../images/lookupiconxp3.gif"" border=""0"" onclick=""top.showpopup('calculator','Calculator',125,100,this,LA_OtherHours_" & RS_WOAssign("PK") & ",self)"" align=""absbottom"" class=""lookupicon"" WIDTH=""16"" HEIGHT=""20"""
				rw "</td>"
				
		End Select
	
	Case "WOC"
				rw "<td class=""data_underline"">"
				rw NullCheck(RS_WOAssign("LaborName"))
				rw "</td>"
				rw "<td class=""data_underline"" align=""left"">" & NullCheckNBSP(RS_WOAssign("Comments")) & "</td>"   'WO87403: Add comments 10/2016
				'If WO_REPORT_LABOR_SHOWACCOUNT = "Yes" Then
				'  rw "<td width=""120"" class=""data_underline"">" & NullCheckNBSP(RS_WOAssign("AccountID")) & "</td>"
				'End If
				rw "<td width=""90"" class=""data_underline"">" & NullCheckNBSP(DateNullCheck(RS_WOAssign("WorkDate"))) & "</td>"
				'If WO_REPORT_LABOR_SHOWSTARTEND = "Yes" Then
				'  rw "<td width=""90"" class=""data_underline"">" & NullCheckNBSP(RS_WOAssign("TimeIn")) & "</td>"
				'  rw "<td width=""90"" class=""data_underline"">" & NullCheckNBSP(RS_WOAssign("TimeOut")) & "</td>"
				'End If
				
				rw "<td class=""data_underline"" align=""center"">" & NullCheckNBSP(RS_WOAssign("RegularHours")) & "</td>"
				rw "<td class=""data_underline"" align=""center"">" & NullCheckNBSP(RS_WOAssign("OvertimeHours")) & "</td>"
				rw "<td class=""data_underline"" align=""center"">" & NullCheckNBSP(RS_WOAssign("OtherHours")) & "</td>"
	
	End Select
	
	rw "</tr>"
End Sub


Sub OutputLaborHeader(LaborFormat)
	rw "<tr>"	

	Select Case LaborFormat
		Case "NONE"
			rw "<td class=""labels"">Labor</td>"
			'If WO_REPORT_LABOR_SHOWACCOUNT = "Yes" Then
			'  rw "<td class=""labels"" width=""120"">Account</td>"
			'End If
			If wostate = "WOC" Then
			rw "<td class=""labels"" align=""center"" width=""50%"" nowrap>&nbsp;Comments&nbsp;</td>"   ' 'WO87403: Add comments 10/2016
			end if
			rw "<td class=""labels"" width=""50"">Work&nbsp;Date</td>"
			'If WO_REPORT_LABOR_SHOWSTARTEND = "Yes" Then
			'  rw "<td nowrap class=""labels"" width=""50"">Start</td>"
			'  rw "<td class=""labels"" width=""50"">End</td>"
			'End If
			
			rw "<td class=""labels"" align=""center"" width=""40"" nowrap>&nbsp;Reg&nbsp;Hrs&nbsp;</td>"
			rw "<td class=""labels"" align=""center"" width=""40"" nowrap>&nbsp;OT&nbsp;Hrs&nbsp;</td>"
			rw "<td class=""labels"" align=""center"" width=""40"" nowrap>&nbsp;Other&nbsp;Hrs&nbsp;</td>"
		Case "ASSIGNED"
			rw "<td class=""labels"">Labor</td>"
			'If WO_REPORT_LABOR_SHOWACCOUNT = "Yes" Then
			'  rw "<td class=""labels"" width=""120"">Account</td>"			
			'End If
			rw "<td class=""labels"" width=""50"">Assigned&nbsp;</td>"
			rw "<td class=""labels"" width=""50"">Work&nbsp;Date</td>"
			'If WO_REPORT_LABOR_SHOWSTARTEND = "Yes" Then
			'  rw "<td nowrap class=""labels"" width=""50"">Start</td>"
			'  rw "<td class=""labels"" width=""50"">End</td>"
			'End If
			rw "<td class=""labels"" align=""center"" width=""40"" nowrap>&nbsp;Reg&nbsp;Hrs&nbsp;</td>"
			rw "<td class=""labels"" align=""center"" width=""40"" nowrap>&nbsp;OT&nbsp;Hrs&nbsp;</td>"
			rw "<td class=""labels"" align=""center"" width=""40"" nowrap>&nbsp;Other&nbsp;Hrs&nbsp;</td>"
		Case "ESTIMATED"
			rw "<td class=""labels"">Craft</td>"
			rw "<td nowrap class=""labels"" align=""center"" width=""40"">Est&nbsp;Hrs&nbsp;</td>"
			rw "<td class=""labels"" width=""200"">Labor</td>"
			'If WO_REPORT_LABOR_SHOWACCOUNT = "Yes" Then
			'  rw "<td class=""labels"" width=""120"">Account</td>"
			'End If
			rw "<td class=""labels"" width=""50"">Work&nbsp;Date</td>"
			'If WO_REPORT_LABOR_SHOWSTARTEND = "Yes" Then
			'  rw "<td nowrap class=""labels"" width=""50"">Start</td>"
			'  rw "<td class=""labels"" width=""50"">End</td>"
			'End If
			rw "<td class=""labels"" align=""center"" width=""40"" nowrap>&nbsp;Reg&nbsp;Hrs&nbsp;</td>"
			rw "<td class=""labels"" align=""center"" width=""40"" nowrap>&nbsp;OT&nbsp;Hrs&nbsp;</td>"
			rw "<td class=""labels"" align=""center"" width=""40"" nowrap>&nbsp;Other&nbsp;Hrs&nbsp;</td>"
	End Select
			
	rw "</tr>"
End Sub

%>