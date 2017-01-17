<%@ EnableSessionState=False Language=VBScript %>
<% Option Explicit %>
<!--#INCLUDE FILE="../common/mc_all.asp" -->
<!--#INCLUDE FILE="includes/mcReport_common.asp" -->

<%
'New Preferences for Version 5
Dim PO_REPORT_POIDUDF, PO_REPORT_SHOWPHONE, PO_REPORT_SHOWENSTATUS, PO_REPORT_SHOWSHIPRECEIVING, PO_REPORT_SHOWENDETAIL
Dim PO_REPORT_SHOWLIC2IU, PO_REPORT_SHOWLIAC, PO_REPORT_SHOWLISAC, PO_REPORT_SHOWLIDISC, PO_REPORT_SHOWLILOC
Dim PO_REPORT_SHOWOPTIONALAPPROVAL, PO_REPORT_SHOWBUYERINFO, ApprovalCount, SY_BASECURRENCY, POCurrency, CurrencySymbol

'New prefs/vars for v7
Dim PO_REPORT_UDF_DISPLAY_1, PO_REPORT_UDF_DISPLAY_2, PO_REPORT_UDF_DISPLAY_3, PO_REPORT_UDF_DISPLAY_4, PO_REPORT_UDF_DISPLAY_5
Dim RS_UDFLabel1, RS_UDFLabel2, RS_UDFLabel3, RS_UDFLabel4, RS_UDFLabel5, RS_PONote
Dim PO_REPORT_SHOWBILLTO, PO_REPORT_SHOWTAXABLE, PO_REPORT_SHOWTAX, PO_REPORT_ITEMDESC, PO_REPORT_ITEMCOMMENTS, PO_REPORT_SHOWSHIPDATE, PO_REPORT_NOTES
Dim numCols

' Standard variables
Dim POSQL, RS_PO, POPK, POID, RS_PODetails, RS_PODocument, RS_POApproval
'Response.Write("URL QueryString: <br>" & Request.QueryString)

'Buffer code begin
Dim BufferCount,BufferCountB
'Buffer end

'Start Report
Call SetPrintedFlag()
Call SetupPOBarcode()
'Get Prefs
Call GetPOReportPrefs()
  
'Get Data
Call SetupPOData()
    
'Output the report
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

	Call OutputHeader()
	Call OutputStandardBodyTag()
	Call OutputEmailMessage()

	If len(errormessage) or len(uf_errormessage) Then
	%>
	<font face="Arial" size="2" color="red"><% =uf_errormessage %></font><br>
	<!--	<font face="Arial" size="2" color="red"><% =errormessage %></font>	-->
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

		If Not RS_PO.EOF Then

            'Begin Buffer Code: 
            Response.Flush
            BufferCount = 0
            If SubReport or FromAgent or (Trim(UCase(Request("EmailReport"))) = "Y") or (Not Request.QueryString("ExportReportOutputType") = "") Then
    	    ' ----- No Script -----
            Else
            Response.Write "<script type='text/javascript'>try{ShowHideLoading('1');} catch(e){} </script> "
            End If
            'End buffer

		'loop through all POs
		Do While Not RS_PO.EOF

        'Begin Buffer Code: Buffer flush before and during report output loop
                BufferCount = BufferCount + 1
                If BufferCount > 20 Then
                    BufferCount = 0
                    Response.Flush
                End If
                'End buffer
		
			'set work order PK
			POPK = RS_PO("POPK")
			ApprovalCount = NumericNullCheck(RS_PO("AuthLevelsRequired"))
			POCurrency = NullCheck(RS_PO("CurrencySymbol"))

			'Check Approval Count - if 0 then 3 else +1
			If ApprovalCount = 0 Then
			  ApprovalCount = 3
			Else
			  ApprovalCount = ApprovalCount + 1
			End If

			'Set Currency
			If NullCheck(POCurrency) <> "" Then
			  CurrencySymbol = POCurrency
			End If
			'CurrencySymbol = Server.HTMLEncode(CurrencySymbol)

			If PO_REPORT_POIDUDF <> "NONE" then
			  If PO_REPORT_POIDUDF = "UDF1" Then
			    POID = NullCheck(RS_PO("UDFChar1"))
			  ElseIf PO_REPORT_POIDUDF = "UDF2" Then
			    POID = NullCheck(RS_PO("UDFChar2"))
			  ElseIf PO_REPORT_POIDUDF = "UDF3" Then
			    POID = NullCheck(RS_PO("UDFChar3"))
			  ElseIf PO_REPORT_POIDUDF = "UDF4" Then
			    POID = NullCheck(RS_PO("UDFChar4"))
			  ElseIf PO_REPORT_POIDUDF = "UDF5" Then
			    POID = NullCheck(RS_PO("UDFChar5"))
			  End If
			Else
			  POID = NullCheck(RS_PO("POID"))
			End If

			rw "<style> .headingPO {font-family:arial; font-size:9pt; color:royalblue; font-weight:bold;}</style>"

			rw "<table border=""0"" width=""100%"">"
				rw "<tr>"
					rw "<td valign=""top"">"
						Call OutputLogoOrName()
					rw "</td>"
					rw "<td align=""right"" valign=""bottom"">"
						Call OutputBarCode(PO_Barcode_POID,PO_BarcodeFormat_POID,POID,"White")
					rw "</td>"
					rw "<td valign=""top"" align=""right"">"
						Call OutputPOHeaderRight()
						If PO_REPORT_POIDUDF <> "NONE" then
						  rw "<div style=""font-family:Arial;font-size:11px;font-weight:normal"">MC PO#: "&POPK&"</div>"
						End If
						If RS_PO("AuthLevelsRequired") > 0 and Not BitNullCheck(RS_PO("IsApproved")) Then
              rw "<div style=""font-family:Arial;font-size:16px;color:red;font-weight:bold;margin-bottom:4px;"">DRAFT / NOT APPROVED</div>"
            End If
					rw "</td>"
				rw "</tr>"
			rw "</table>"
			rw "<br>"
			rw "<table border=""0"" width=""100%"">"
				rw "<tr>"
					rw "<td width=""48%"" valign=""top"">"
						rw "<table border=""0"" width=""100%"" bgColor=""#F5F5F5"" style=""border:1px solid"" height=""150"">"
							rw "<tr>"
								rw "<td valign=""top"" style=""padding-left:10px;padding-top:10px"" class=""headingPO"">To:</td>"
							rw "</tr>"
							rw "<tr>"
								rw "<td valign=""top"" style=""padding-left:30px;padding-bottom:10px"" class=""data"">"
									rw "<span style=""font-size:11pt;""><b>" & NullCheck(RS_PO("CompanyName")) & "</b></span><br>"
									If PO_REPORT_SHOWPHONE = "Yes" Then
									  If NullCheck(RS_PO("CompanyAttn")) <> "" Then
									    rw "Attn: " & NullCheck(RS_PO("CompanyAttn"))& "<br>"
									  End If
									End If
									rw NullCheck(RS_PO("CompanyAddress1")) & "<br>"
									If Not RS_PO("CompanyAddress2") = "" Then
										rw NullCheck(RS_PO("CompanyAddress2")) & "<br><br>"
									End If
									rw NullCheck(RS_PO("CompanyCity")) & "," & NullCheck(RS_PO("CompanyState")) & " " & NullCheck(RS_PO("CompanyZip")) & "<br><br>"

									If PO_REPORT_SHOWPHONE = "Yes" Then
									  If NullCheck(RS_PO("Phone")) <> "" Then
									  rw "Phone: " & NullCheck(RS_PO("Phone"))
									  End If
									  If NullCheck(RS_PO("Fax")) <> "" then
									    rw "<br>Fax: " & NullCheck(RS_PO("Fax"))
									  End If
									End If
									If (NullCheck(RS_PO("CompanyAttn")) <> "" AND PO_REPORT_SHOWPHONE = "No") Then
									  rw "Attn: " & NullCheck(RS_PO("CompanyAttn"))
									End If
								rw "</td>"
							rw "</tr>"
						rw "</table>"
					rw "</td>"
					rw "<td width=""4%"">&nbsp;</td>"
					rw "<td width=""48%"" valign=""top"" align=""right"">"
						rw "<table border=""0"" width=""100%"" bgColor=""#F5F5F5"" style=""border:1px solid"" height=""150"">"
							rw "<tr>"
								rw "<td valign=""top"" style=""padding-left:10px;padding-top:10px"" class=""headingPO"">Ship To:</td>"
							rw "</tr>"
							rw "<tr>"
								rw "<td valign=""top"" style=""padding-left:30px;padding-bottom:10px"" class=""data"">"
									rw "<span style=""font-size:11pt;""><b>" & NullCheck(RS_PO("ShipToName")) & "</b></span><br>"
									If PO_REPORT_SHOWPHONE = "Yes" Then
									  If NullCheck(RS_PO("ShipToAttention")) <> "" Then
										  rw "Attn: " & NullCheck(RS_PO("ShipToAttention")) & "<br>"
									  End If
									End If
									rw NullCheck(RS_PO("ShipToAddress1")) & "<br>"
									If Not RS_PO("ShipToAddress2") = "" Then
										rw NullCheck(RS_PO("ShipToAddress2")) & "<br>"
									End If
									If Not RS_PO("ShipToAddress3") = "" Then
										rw NullCheck(RS_PO("ShipToAddress3")) & "<br><br>"
									End If
									If PO_REPORT_SHOWPHONE = "Yes" Then
									  If NullCheck(RS_PO("ShipToPhone")) <> "" Then
									    rw "Phone: " & NullCheck(RS_PO("ShipToPhone"))
									  End If
									  If NullCheck(RS_PO("ShipToFax")) <> "" Then
									    rw "<Br>Fax: " & NullCheck(RS_PO("ShipToFax"))
									  End If
									End If
									If (NullCheck(RS_PO("ShipToAttention")) <> "" AND PO_REPORT_SHOWPHONE = "No") Then
										rw "Attn: " & NullCheck(RS_PO("ShipToAttention"))
									End If
								rw "</td>"
							rw "</tr>"
						rw "</table>"
					rw "</td>"
				rw "</tr>"
			rw "</table>"
			rw "<br>"
			rw "<table border=""0"" width=""100%"" bgColor=""#F5F5F5"" style=""border:1px solid"">"
				rw "<tr>"
				If PO_REPORT_SHOWENDETAIL = "Yes" Then
				  rw "<td width=""33%"">"
				Else
				  rw "<td width=""78%"">"
				End If
				rw "<table cellspacing=""0"" cellpadding=""0"" border=""0"" width=""50%"">"
                dim techName
                techName = NullCheck(RS_PO("BuyerName"))
				rw "<tr>"
                ''''''''''''''''''''''''' WO88677
                'rw "<td class=""headingPO"" valign=""top"" style=""padding-left:10px;padding-right:10px;padding-top:10px"">Purchasing Agent:</td>"
				rw "<td class=""headingPO"" valign=""top"" style=""padding-left:10px;padding-right:10px;padding-top:10px"">Technician:</td>"
				rw "<td class=""data"" valign=""top"" style=""padding-top:10px"">" & techName & "</td>"
                rw "</tr>"
                ''' Adding the Department here
                rw "<tr>"
                                
                dim zSQL                
                dim RS_Labor
                dim techPK
                techPK = NullCheck(RS_PO("BuyerPK"))
                zSQL = "select DepartmentID from Labor where LaborPK= '" & techPK & "'"
                Set RS_Labor = db.runSQLReturnRS(zSQL,"")
                dim zDepartment 
                zDepartment= RS_Labor("DepartmentID")
                rw "<td class=""headingPO"" valign=""top"" style=""padding-left:10px;padding-right:10px;"">Department:</td>"
		        rw "<td class=""data"" valign=""top"" style=""padding-left:0px;padding-right:0px"">" & zDepartment & "</td>"
                rw "</tr>"
                RS_Labor.Close : set RS_Labor= Nothing   
                ''''''''''''''''''''''''' WO88677
            	rw "<tr>"
				'	rw "<td class=""headingPO"" style=""padding-left:10px;padding-right:10px;"">Vendor ID:</td>"
				'	rw "<td class=""data"">" & NullCheck(RS_PO("VendorID")) & "</td>"
				'rw "</tr>"
				'If PO_REPORT_SHOWENSTATUS = "Yes" Then
				'  If NullCheck(RS_PO("Requested")) <> "" Then
				'    rw "<tr>"
				'	    rw "<td class=""headingPO"" valign=""top"" style=""padding-left:10px;padding-right:10px;"">Requested:</td>"
				'	    rw "<td class=""data"" valign=""top"" style=""padding-left:0px;"">"&NullCheck(RS_PO("Requested"))&" by " & NullCheck(RS_PO("RequestedBy")) & "</td>"
				'    rw "</tr>"
				'  End If
				  'If NullCheck(RS_PO("AuthLevelsRequired")) > 0 Then
				    'rw "<tr>"
				    '  If NullCheck(RS_PO("IsApproved")) Then
				    '    If NullCheck(RS_PO("Approved")) <> "" Then
				    '      rw "<td class=""headingPO"" valign=""top"" style=""padding-left:10px;padding-right:10px;"">Approved (final):</td>"
				    '      rw "<td class=""data"" valign=""top"" style=""padding-left:0px;"">"&NullCheck(RS_PO("Approved"))
				    '      If NullCheck(RS_PO("ApprovedBy")) <> "" Then
				    '        rw " by " & NullCheck(RS_PO("ApprovedBy"))
				    '      End If
				    '      rw "</td>"
				    '    End If
				    '  Else
					'      If NullCheck(RS_PO("Approved")) <> "" Then
					'        rw "<td class=""headingPO"" valign=""top"" style=""padding-left:10px;padding-right:10px;"">Approved (last):</td>"
					'        rw "<td class=""data"" valign=""top"" style=""padding-left:0px;"">"&NullCheck(RS_PO("Approved"))
					'        If NullCheck(RS_PO("ApprovedBy")) <> "" Then
					'          rw " by " & NullCheck(RS_PO("ApprovedBy"))
					'        End If
					'        rw "</td>"
					'      End If
					'    End If
				    'rw "</tr>"
				    'End If
				'  If NullCheck(RS_PO("Issued")) <> "" Then
				'    rw "<tr>"
					    rw "<td class=""headingPO"" valign=""top"" style=""padding-left:10px;padding-right:10px;"">Date of Issue:</td>"
		                rw "<td class=""data"" valign=""top"" style=""padding-left:0px;padding-right:0px"">"&NullCheck(RS_PO("Issued"))  & "</td>"
        'rw "<td class=""data"" valign=""top"" style=""padding-left:0px;padding-right:0px"">"&NullCheck(RS_PO("Issued"))&" by " & NullCheck(RS_PO("IssuedBy")) & "</td>"
				'    rw "</tr>"
				'  End If
				'  rw "<tr style=""height:10px;""><td></td></tr>"
				'Else
				'	rw "<tr>"
				'		rw "<td class=""headingPO"" valign=""top"" style=""padding-left:10px;padding-right:10px;"">Date of Issue:</td>"
				'		rw "<td class=""data"" valign=""top"">" & DateTimeNullCheckAT(RS_PO("Issued")) & "</td>"
				'	rw "</tr>"
				'	rw "<tr>"
				'		rw "<td class=""headingPO"" valign=""top"" style=""padding-left:10px;padding-right:10px;padding-bottom:10px"">Date Requested:</td>"
				'		rw "<td class=""data"" valign=""top"" style=""padding-bottom:10px"">" & DateTimeNullCheckAT(RS_PO("Requested")) & "</td>"
					rw "</tr>"
				'End If
 
				rw "</table>"
                rw "</td>"
                rw "<td>"            
               ' rw "<table style=""width:100%;vertical-align:top;"">"
                'rw "<tr>"
				'If PO_REPORT_SHOWENDETAIL = "Yes" Then
				'	rw "<td valign=""top"" width=""33%"">"
				'		rw "<table cellspacing=""0"" cellpadding=""0"" border=""0"">"
				'			If PO_REPORT_SHOWSHIPDATE = "Yes" Then
				'			'If NullCheck(RS_PO("ShipDate")) <> "" Then
				'				rw "<tr>"
				'					rw "<td class=""headingPO"" valign=""top"" style=""padding-left:10px;padding-right:10px;padding-top:10px"">Shipping Date:</td>"
				'					rw "<td class=""data"" valign=""top"" style=""padding-top:10px"">" & DateNullCheck(RS_PO("ShipDate")) & "</td>"
				'				rw "</tr>"
				'			'End If
				'			End If
				'			rw "<tr>"
				'				rw "<td class=""headingPO"" valign=""top"" style=""padding-left:10px;padding-right:10px;"">Repair Center:</td>"
				'				rw "<td class=""data"" valign=""top"" style="""">" & NullCheck(RS_PO("RepairCenterName")) & "</td>"
				'			rw "</tr>"
				'			rw "<tr>"
				'				rw "<td class=""headingPO"" valign=""top"" style=""padding-left:10px;padding-right:10px;padding-bottom:0px"">Department:</td>"
				'				rw "<td class=""data"" valign=""top"" style=""padding-bottom:0px;"">" & NullCheck(RS_PO("DepartmentName")) & "</td>"
				'			rw "</tr>"
				'			If NullCheck(RS_PO("TenantName")) <> "" Then
				'				rw "<tr>"
				'					rw "<td class=""headingPO"" valign=""top"" style=""padding-left:10px;padding-right:10px;"">Customer:</td>"
				'					rw "<td class=""data"" valign=""top"" style="""">" & NullCheck(RS_PO("TenantName")) & "</td>"
				'				rw "</tr>"
				'			End If
				'			If NullCheck(RS_PO("AccountName")) <> "" Then
				'				rw "<tr>"
				'					rw "<td class=""headingPO"" valign=""top"" style=""padding-left:10px;padding-right:10px;padding-botom:10px;"">Account:</td>"
				'					rw "<td class=""data"" valign=""top"" style=""padding-botom:10px;"">" & NullCheck(RS_PO("AccountName")) & "</td>"
				'				rw "</tr>"
				'			End If
				'		rw "</table>"
				'	rw "</td>"
				'	rw "<td valign=""top"" width=""34%"">"
				'Else
				'  rw "<td valign=""top"" width=""50%"">"
				'End If
				'If PO_REPORT_SHOWSHIPRECEIVING = "Yes" Then
				'  rw "<table cellspacing=""0"" cellpadding=""0"" border=""0"""
				'  rw "<tr>"
				'	  rw "<td class=""headingPO"" valign=""top"" style=""padding-top:10px;padding-right:10px;"">Payment Terms:</td>"
				'	  rw "<td class=""data"" valign=""top"" style=""padding-top:10px"">" & NullCheck(RS_PO("TermsDesc")) & "</td>"
				'  rw "</tr>"
				'  rw "<tr>"
				'	  rw "<td class=""headingPO"" valign=""top"" style=""padding-right:10px;"">Ship Via:</td>"
				'	  rw "<td class=""data"" valign=""top"">" & NullCheck(RS_PO("ShippingMethodDesc")) & "</td>"
				'  rw "</tr>"
				'  rw "<tr>"
				'	  rw "<td class=""headingPO"" valign=""top"" style=""padding-right:10px;"">Freight Terms:</td>"
				'	  rw "<td class=""data"" valign=""top"">" & NullCheck(RS_PO("FreightTermsDesc")) & "</td>"
				'  rw "</tr>"
				'  rw "<tr>"
				'	  rw "<td class=""headingPO"" valign=""top"" style=""padding-bottom:10px; padding-right:10px;"">F.O.B:</td>"
				'	  rw "<td class=""data"" valign=""top"" style=""padding-bottom:10px"">" & NullCheck(RS_PO("FOBPoint")) & "</td>"
				'  rw "</tr>"
				'  rw "</table>"
				'End If
				'rw "</td>"
                rw "<td class=""headingPO"" valign=""middle"" style=""padding-left:0px; padding-right:0px;padding-top:0px;text-align: left;"">HED ORDER #:</td>"
                rw "<td style=""width:50""></td>"
				rw "</tr>"
			'rw "</table>"
			'rw "<br>"
			
			'v7 UDF Fields
			'If (PO_REPORT_UDF_DISPLAY_1 <> "" OR PO_REPORT_UDF_DISPLAY_2 <> "" OR PO_REPORT_UDF_DISPLAY_3 <> "" OR PO_REPORT_UDF_DISPLAY_4 <> "" OR PO_REPORT_UDF_DISPLAY_5 <> "") Then
			'	rw "<table border=""0"" width=""100%"" bgColor=""#F5F5F5"" style=""border:1px solid"">"
			'		rw "<tr>"
			'			'Display up to 5 UDFs
			'			Dim UDFcount
			'			UDFcount = 0
            '
			'			rw "<tr valign=""top"">"
			'			'UDF POSITION 1
			'			If PO_REPORT_UDF_DISPLAY_1 <> "" Then
			'				UDFcount = UDFcount + 1
			'				sql = _
			'				"SELECT " &_
			'				"   data_type " &_
			'				"  ,column_name " &_
			'				"  ,Label =  CASE WHEN column_name <> ISNULL(field_label,'') THEN field_label ELSE column_name END " &_
			'				"FROM " &_
			'				"  DataDict WITH (NOLOCK) " &_
			'				"WHERE " &_
			'				"  table_name = 'PurchaseOrder' " &_
			'				"  AND column_name = '" & PO_REPORT_UDF_DISPLAY_1 & "' "
            '
			'				Set RS_UDFLabel1 = db.runSQLReturnRS(sql,"")
			'				Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
            '
			'				rw "<td nowrap width=""50%"" style=""padding-left:5px;"">"
			'					rw "<table border=""0"" cellspacing=""0"" cellpadding=""0"">"
			'						rw "<tr>"
			'							If NullCheck(RS_UDFLabel1("data_type")) = "bit" Then
			'								rw "<td class=""headingPO"" valign=""top"">" & NullCheck(RS_UDFLabel1("Label")) & "</td>"
			'								If BitNullCheck(RS_PO("" & PO_REPORT_UDF_DISPLAY_1 & "")) Then
			'									rw "<td valign=""top"" style=""padding-left:0.5em;""><img src=""" & GetSession("webHTTP") & GetWebServer() & Application("web_path") & Application("mapp_path") & "images/checkbox_checked.jpg"" border=""0""></td>"
			'								Else
			'									rw "<td valign=""top"" style=""padding-left:0.5em;""><img src=""" & GetSession("webHTTP") & GetWebServer() & Application("web_path") & Application("mapp_path") & "images/checkbox_notchecked.jpg"" border=""0""></td>"
			'								End If
			'							Else
			'								rw "<td class=""headingPO"" valign=""top"">" & NullCheck(RS_UDFLabel1("Label")) & ":</td>"
			'								rw "<td class=""data"" valign=""top"" style=""padding-left:0.5em;"">" & NullCheck(RS_PO("" & PO_REPORT_UDF_DISPLAY_1 & "")) & "</td>"
			'							End If
			'						rw "</tr>"
			'					rw "</table>"
			'				rw "</td>"
			'			End If
            '
			'			'UDF POSITION 2
			'			If PO_REPORT_UDF_DISPLAY_2 <> "" Then
			'				UDFcount = UDFcount + 1
			'				sql = _
			'				"SELECT " &_
			'				"   data_type " &_
			'				"  ,column_name " &_
			'				"  ,Label =  CASE WHEN column_name <> ISNULL(field_label,'') THEN field_label ELSE column_name END " &_
			'				"FROM " &_
			'				"  DataDict WITH (NOLOCK) " &_
			'				"WHERE " &_
			'				"  table_name = 'PurchaseOrder' " &_
			'				"  AND column_name = '" & PO_REPORT_UDF_DISPLAY_2 & "' "
            '
			'				Set RS_UDFLabel2 = db.runSQLReturnRS(sql,"")
			'				Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
            '
			'				rw "<td nowrap width=""50%"" style=""padding-left:5px;"">"
			'					rw "<table border=""0"" cellspacing=""0"" cellpadding=""0"">"
			'						rw "<tr>"
			'							If NullCheck(RS_UDFLabel2("data_type")) = "bit" Then
			'								rw "<td class=""data"" valign=""bottom"">" & NullCheck(RS_UDFLabel2("Label")) & "</td>"
			'								If BitNullCheck(RS_PO("" & PO_REPORT_UDF_DISPLAY_2 & "")) Then
			'									rw "<td valign=""top"" style=""padding-left:0.5em;""><img src=""" & GetSession("webHTTP") & GetWebServer() & Application("web_path") & Application("mapp_path") & "images/checkbox_checked.jpg"" border=""0""></td>"
			'								Else
			'									rw "<td valign=""top"" style=""padding-left:0.5em;""><img src=""" & GetSession("webHTTP") & GetWebServer() & Application("web_path") & Application("mapp_path") & "images/checkbox_notchecked.jpg"" border=""0""></td>"
			'								End If
			'							Else
			'								rw "<td class=""headingPO"" valign=""top"">" & NullCheck(RS_UDFLabel2("Label")) & ":</td>"
			'								rw "<td class=""data"" valign=""top"" style=""padding-left:0.5em;"">" & NullCheck(RS_PO("" & PO_REPORT_UDF_DISPLAY_2 & "")) & "</td>"
			'							End If
			'						rw "</tr>"
			'					rw "</table>"
			'				rw "</td>"
			'			End If
            '
			'			If UDFCount = 2 Then
			'				UDFCount = 0
			'				rw "</tr>"
			'				rw "<tr valign=""top"">"
			'			End If
            '
			'			'UDF POSITION 3
			'			If PO_REPORT_UDF_DISPLAY_3 <> "" Then
			'				UDFcount = UDFcount + 1
			'				sql = _
			'				"SELECT " &_
			'				"   data_type " &_
			'				"  ,column_name " &_
			'				"  ,Label =  CASE WHEN column_name <> ISNULL(field_label,'') THEN field_label ELSE column_name END " &_
			'				"FROM " &_
			'				"  DataDict WITH (NOLOCK) " &_
			'				"WHERE " &_
			'				"  table_name = 'PurchaseOrder' " &_
			'				"  AND column_name = '" & PO_REPORT_UDF_DISPLAY_3 & "' "
            '
			'				Set RS_UDFLabel3 = db.runSQLReturnRS(sql,"")
			'				Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
            '
			'				rw "<td nowrap width=""50%"" style=""padding-left:5px;"">"
			'					rw "<table border=""0"" cellspacing=""0"" cellpadding=""0"">"
			'						rw "<tr>"
			'							If NullCheck(RS_UDFLabel3("data_type")) = "bit" Then
			'								rw "<td class=""headingPO"" valign=""bottom"">" & NullCheck(RS_UDFLabel3("Label")) & ":</td>"
			'								If BitNullCheck(RS_PO("" & PO_REPORT_UDF_DISPLAY_3 & "")) Then
			'									rw "<td valign=""top"" style=""padding-left:0.5em;""><img src=""" & GetSession("webHTTP") & GetWebServer() & Application("web_path") & Application("mapp_path") & "images/checkbox_checked.jpg"" border=""0""></td>"
			'								Else
			'									rw "<td valign=""top"" style=""padding-left:0.5em;""><img src=""" & GetSession("webHTTP") & GetWebServer() & Application("web_path") & Application("mapp_path") & "images/checkbox_notchecked.jpg"" border=""0""></td>"
			'								End If
			'							Else
			'								rw "<td class=""headingPO"" valign=""top"">" & NullCheck(RS_UDFLabel3("Label")) & ":</td>"
			'								rw "<td class=""data"" valign=""top"" style=""padding-left:0.5em;"">" & NullCheck(RS_PO("" & PO_REPORT_UDF_DISPLAY_3 & "")) & "</td>"
			'							End If
			'						rw "</tr>"
			'					rw "</table>"
			'				rw "</td>"
			'			End If
            '
			'			If UDFCount = 2 Then
			'				UDFCount = 0
			'				rw "</tr>"
			'				rw "<tr valign=""top"">"
			'			End If
            '
			'			'UDF POSITION 4
			'			If PO_REPORT_UDF_DISPLAY_4 <> "" Then
			'				UDFcount = UDFcount + 1
			'				sql = _
			'				"SELECT " &_
			'				"   data_type " &_
			'				"  ,column_name " &_
			'				"  ,Label =  CASE WHEN column_name <> ISNULL(field_label,'') THEN field_label ELSE column_name END " &_
			'				"FROM " &_
			'				"  DataDict WITH (NOLOCK) " &_
			'				"WHERE " &_
			'				"  table_name = 'PurchaseOrder' " &_
			'				"  AND column_name = '" & PO_REPORT_UDF_DISPLAY_4 & "' "
            '
			'				Set RS_UDFLabel4 = db.runSQLReturnRS(sql,"")
			'				Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
            '
			'				rw "<td nowrap width=""50%"" style=""padding-left:5px;"">"
			'					rw "<table border=""0"" cellspacing=""0"" cellpadding=""0"">"
			'						rw "<tr>"
			'							If NullCheck(RS_UDFLabel4("data_type")) = "bit" Then
			'								rw "<td class=""headingPO"" valign=""top"">" & NullCheck(RS_UDFLabel4("Label")) & "</td>"
			'								If BitNullCheck(RS_PO("" & PO_REPORT_UDF_DISPLAY_4 & "")) Then
			'									rw "<td valign=""top"" style=""padding-left:0.5em;""><img src=""" & GetSession("webHTTP") & GetWebServer() & Application("web_path") & Application("mapp_path") & "images/checkbox_checked.jpg"" border=""0""></td>"
			'								Else
			'									rw "<td valign=""top"" style=""padding-left:0.5em;""><img src=""" & GetSession("webHTTP") & GetWebServer() & Application("web_path") & Application("mapp_path") & "images/checkbox_notchecked.jpg"" border=""0""></td>"
			'								End If
			'							Else
			'								rw "<td class=""headingPO"" valign=""top"">" & NullCheck(RS_UDFLabel4("Label")) & ":</td>"
			'								rw "<td class=""data"" valign=""top"" style=""padding-left:0.5em;"">" & NullCheck(RS_PO("" & PO_REPORT_UDF_DISPLAY_4 & "")) & "</td>"
			'							End If
			'						rw "</tr>"
			'					rw "</table>"
			'				rw "</td>"
			'			End If
            '
			'			If UDFCount = 2 Then
			'				UDFCount = 0
			'				rw "</tr>"
			'				rw "<tr valign=""top"">"
			'			End If
            '
			'			'UDF POSITION 5
			'			If PO_REPORT_UDF_DISPLAY_5 <> "" Then
			'				UDFcount = UDFcount + 1
			'				sql = _
			'				"SELECT " &_
			'				"   data_type " &_
			'				"  ,column_name " &_
			'				"  ,Label =  CASE WHEN column_name <> ISNULL(field_label,'') THEN field_label ELSE column_name END " &_
			'				"FROM " &_
			'				"  DataDict WITH (NOLOCK) " &_
			'				"WHERE " &_
			'				"  table_name = 'PurchaseOrder' " &_
			'				"  AND column_name = '" & PO_REPORT_UDF_DISPLAY_5 & "' "
            '
			'				Set RS_UDFLabel5 = db.runSQLReturnRS(sql,"")
			'				Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
            '
			'				rw "<td nowrap width=""50%"" style=""padding-left:5px;"">"
			'					rw "<table border=""0"" cellspacing=""0"" cellpadding=""0"">"
			'						rw "<tr>"
			'							If NullCheck(RS_UDFLabel5("data_type")) = "bit" Then
			'								rw "<td class=""headingPO"" valign=""top"">" & NullCheck(RS_UDFLabel5("Label")) & "</td>"
			'								If BitNullCheck(RS_PO("" & PO_REPORT_UDF_DISPLAY_5 & "")) Then
			'									rw "<td valign=""top"" style=""padding-left:0.5em;""><img src=""" & GetSession("webHTTP") & GetWebServer() & Application("web_path") & Application("mapp_path") & "images/checkbox_checked.jpg"" border=""0""></td>"
			'								Else
			'									rw "<td valign=""top"" style=""padding-left:0.5em;""><img src=""" & GetSession("webHTTP") & GetWebServer() & Application("web_path") & Application("mapp_path") & "images/checkbox_notchecked.jpg"" border=""0""></td>"
			'								End If
			'							Else
			'								rw "<td class=""headingPO"" valign=""top"">" & NullCheck(RS_UDFLabel5("Label")) & ":</td>"
			'								rw "<td class=""data"" valign=""top"" style=""padding-left:0.5em;"">" & NullCheck(RS_PO("" & PO_REPORT_UDF_DISPLAY_5 & "")) & "</td>"
			'							End If
			'						rw "</tr>"
			'					rw "</table>"
			'				rw "</td>"
			'			End If
			'		rw "</tr>"
			'	rw "</table>"
			'	rw "<br>"
			'End If

			'PO_REPORT_SHOWLIC2IU, PO_REPORT_SHOWLIAC, PO_REPORT_SHOWLISAC, PO_REPORT_SHOWLIDISC, PO_REPORT_SHOWLILOC
            rw "</td></tr></table></td></tr></table><br>"
			rw "<table width=""100%"" style=""border:1px solid #CCCCCC; border-collapse:collapse;"" cellpadding=""2"" cellspacing=""0"">"
				rw "<tr>"
				rw "<td class=""heading""  style=""border-bottom:1px solid gray;vertical-align:top; text-align:left;width:6%;"">Line #</td>"
                rw "<td class=""heading""  style=""border-bottom:1px solid gray;vertical-align:top; text-align:center;width:14%;"">CKE Part #</td>"
                rw "<td class=""heading""  style=""border-bottom:1px solid gray;vertical-align:top; text-align:left;width:15%;"">HED Part #</td>"
                rw "<td class=""heading""  style=""border-bottom:1px solid gray;vertical-align:top; text-align:left;"">Description</td>"
                rw "<td class=""heading""  style=""border-bottom:1px solid gray;vertical-align:top; text-align:left;width:10%"">Order Qty</td>"
                rw "<td class=""heading""  style=""border-bottom:1px solid gray;vertical-align:top; text-align:center;width:12%"">Back Order Qty</td>"
					'If PO_Barcode_PartID Then
					'	numCols = numCols + 1
					'	rw "<td valign=""top"" class=""heading"" width=""1%"" style=""border-bottom:1px solid gray;"">Item Barcode</td>"
					'End If
					'If PO_REPORT_SHOWLIC2IU = "Yes" Then
					'	rw "<td valign=""top"" class=""heading"" align=""left"" width=""10%"" style=""border-bottom:1px solid gray;"">Order Qty</td>"
					'	rw "<td valign=""top"" class=""heading"" align=""center"" width=""10%"" style=""border-bottom:1px solid gray;"">Issue Qty Total</td>"
					'ELSE
					'	rw "<td valign=""top"" class=""heading"" align=""right"" width=""10%"" style=""border-bottom:1px solid gray;""> Order Qty</td>"
					'	rw "<td valign=""top"" class=""heading"" width=""10%"" style=""border-bottom:1px solid gray;"">Units</td>"
					'End If
					'If PO_REPORT_ITEMDESC = "Yes" Then
					'	rw "<td valign=""top"" class=""heading"" width=""30%"" style=""border-bottom:1px solid gray;"">Name</td>"
					'Else
					'	rw "<td valign=""top"" class=""heading"" width=""30%"" style=""border-bottom:1px solid gray;"">Description</td>"
					'End If
					'rw "<td nowrap valign=""top"" class=""heading"" width=""15%"" style=""border-bottom:1px solid gray;"">Vendor Item #</td>"
					'If PO_REPORT_SHOWLIAC = "Yes" Then
					'	numCols = numCols + 1
					'	rw "<td nowrap valign=""top"" class=""heading"" width=""10%"" style=""border-bottom:1px solid gray;"">Account</td>"
					'End If
					'If PO_REPORT_SHOWLISAC = "Yes" Then
					'	numCols = numCols + 1
					'	rw "<td nowrap valign=""top"" class=""heading"" width=""10%"" style=""border-bottom:1px solid gray;"">Sub-Account</td>"
					'End If
					'If PO_REPORT_SHOWLILOC = "Yes" Then
					'	numCols = numCols + 1
					'	rw "<td nowrap valign=""top"" class=""heading"" width=""10%"" style=""border-bottom:1px solid gray;"">Stockroom / Bin</td>"
					'End If
					'rw "<td nowrap valign=""top"" class=""heading"" align=""right"" width=""10%"" style=""border-bottom:1px solid gray;"">Unit Price</td>"
					'If PO_REPORT_SHOWLIDISC = "Yes" Then
					'	numCols = numCols + 1
					'	rw "<td nowrap valign=""top"" class=""heading"" align=""right"" width=""10%"" style=""border-bottom:1px solid gray;"">Discount %</td>"
					'End If
					'If PO_REPORT_SHOWTAXABLE = "Yes" Then
					'	numCols = numCols + 1
					'	rw "<td nowrap valign=""top"" class=""heading"" align=""right"" width=""10%"" style=""border-bottom:1px solid gray;"">Taxable?</td>"
					'End If
					'If PO_REPORT_SHOWTAX = "Yes" Then
					'	numCols = numCols + 2
					'	rw "<td valign=""top"" class=""heading"" align=""right"" width=""10%"" style=""border-bottom:1px solid gray;"">Tax Rate</td>"
					'	rw "<td valign=""top"" class=""heading"" align=""right"" width=""10%"" style=""border-bottom:1px solid gray;"">Tax Cost</td>"
					'End If
					'rw "<td nowrap valign=""top"" class=""heading"" align=""right"" width=""10%"" style=""border-bottom:1px solid gray;"">Line Total</td>"
				rw "</tr>"
				numCols = numCols + 7

				' set 2 alternating colors
				bgColor1 = "#EFEFEF"
				bgColor2 = "#FFFFFF"

				Dim GotPOLineItem
				GotPOLineItem = False
				Do While NullCheck(RS_PODetails("POPK")) = NullCheck(POPK)
					GotPOLineItem = True
					' alternate bg colors
					If altBgColor = bgColor1 Then
						altBgColor = bgColor2
					Else
						altBgColor = bgColor1
					End If

					rw "<TR bgColor=""" & altBgColor & """>"
					rw "<td style=""vertical-align:top;text-align:center;"" class=""data"">" & NullCheck(RS_PODetails("LineItemNo")) & "</td>"
                    rw "<td style=""vertical-align:top;text-align:center;"" class=""data"">" & NullCheck(RS_PODetails("PartID")) & "</td>"
                    rw "<td style=""vertical-align:top;text-align:left;"" class=""data"">" & NullCheck(RS_PODetails("VendorPartNumber")) & "</td>"
                    rw "<td style=""vertical-align:top;text-align:left;"" class=""data"">" & NullCheck(RS_PODetails("PartName")) & "</td>"
					'If PO_Barcode_PartID Then
					'	rw "<td style=""padding-right:10px;"" valign=""top"" class=""data"">"
					'	    If Not UCase(NullCheck(RS_PODetails("PartID"))) = "MANUAL ENTRY" Then
					'			If altBgColor = "#FFFFFF" Then
					'				Call OutputBarCode(PO_Barcode_PartID,IN_BarcodeFormat_PartID,NullCheck(RS_PODetails("PartID")),"White")
					'			Else
					'				Call OutputBarCode(PO_Barcode_PartID,IN_BarcodeFormat_PartID,NullCheck(RS_PODetails("PartID")),"Gray")
					'			End If
					'		Else
					'			rw "&nbsp;"
					'		End If
					'	rw "</td>"
					'End If
					'If PO_REPORT_SHOWLIC2IU = "Yes" Then
					  'dim totalIssueUnits, orderQty, orderQtyDI, totalIssueUnitsDI
					 ' orderQty = CStr(NullCheck(RS_PODetails("OrderUnitQty"))) + " " + NullCheck(RS_PODetails("OrderUnitsDesc")) + " of " + CStr(NumericNullCheck(RS_PODetails("ConversionToIssueUnits"))) + " " +  NullCheck(RS_PODetails("IssueUnitsDesc"))
					 ' orderQtyDI = CStr(NullCheck(RS_PODetails("OrderUnitQty"))) + " " + NullCheck(RS_PODetails("OrderUnitsDesc"))
					 ' If CInt(NumericNullCheck(RS_PODetails("ConversionToIssueUnits"))) > 1 Then
					  '  orderQtyDI = orderQtyDI + " of " + CStr(NumericNullCheck(RS_PODetails("ConversionToIssueUnits")))
					 ' End If
					 ' totalIssueUnits = (CInt(NumericNullCheck(RS_PODetails("ConversionToIssueUnits")))*CInt(NumericNullCheck(RS_PODetails("OrderUnitQty"))))

					'  If Not BitNullCheck(RS_PODetails("DirectIssue")) Then
					'    rw "<td valign=""top"" class=""data"" align=""left"">" & orderQty & "</td>"
					    'rw "<td valign=""top"" class=""data"" align=""center"">" & totalIssueUnits & "</td>"
					'  Else
					'    rw "<td valign=""top"" class=""data"" align=""left"">" & orderQtyDI & "</td>"
					    'rw "<td valign=""top"" class=""data"" align=""center"">" & totalIssueUnits & "</td>"
					'  End If
					'Else
					rw "<td valign=""top"" class=""data"" align=""center"">" & NullCheck(RS_PODetails("OrderUnitQty")) & "</td>"
                    rw "<td></td>"
					  'rw "<td valign=""top"" class=""data"">" & NullCheck(RS_PODetails("OrderUnitsDesc")) & "</td>"
					'End If
					'rw "<td class=""data"">" & NullCheck(RS_PODetails("PartID")) & "</td>"
					'If Not UCase(NullCheck(RS_PODetails("PartID"))) = "MANUAL ENTRY" Then
					'	rw "<td valign=""top"" class=""data"">" & NullCheck(RS_PODetails("PartName")) & " (" & NullCheck(RS_PODetails("PartID")) & ")</td>"
					'Else
					'	rw "<td valign=""top"" class=""data"">" & NullCheck(RS_PODetails("PartName")) & "</td>"
					'End If
					'rw "<td valign=""top"" class=""data"">" & NullCheck(RS_PODetails("VendorPartNumber")) & "</td>"
					'If PO_REPORT_SHOWLIAC = "Yes" Then
					'  rw "<td valign=""top"" class=""data"">" & NullCheck(RS_PODetails("AccountName")) & "</td>"
					'End If
					'If PO_REPORT_SHOWLISAC = "Yes" Then
					'  rw "<td valign=""top"" class=""data"">" & NullCheck(RS_PODetails("SubAccountName")) & "</td>"
					'End If
				'	If PO_REPORT_SHOWLILOC = "Yes" Then
				'	  If NullCheck(RS_PODetails("WOPK")) <> "" Then
				'	    rw "<td valign=""top"" class=""data"" nowrap><i>(Directly Issued for " & NullCheck(RS_PODetails("WOID")) & ")</i></td>"
				'	  Else
				'	    If NullCheck(RS_PODetails("LocationID")) = "" Then
				'	      rw "<td valign=""top"" class=""data"" nowrap><i>(Directly Issued)</i>"
				'	    Else
				'	      rw "<td valign=""top"" class=""data"" nowrap>" & NullCheck(RS_PODetails("LocationID"))
				'	      If NullCheck(RS_PODetails("Bin")) <> "" Then
				'	        rw " / " & NullCheck(RS_PODetails("Bin"))
				'	      End If
				'	    End If
				'	    rw "</td>"
				'	  End If
				'	End If
				'	rw "<td nowrap valign=""top"" class=""data"" align=""right"">" & CurrencySymbol & " " & FormatNumber(RS_PODetails("OrderUnitPrice")) & "</td>"
				'	If PO_REPORT_SHOWLIDISC = "Yes" Then
				'	  rw "<td valign=""top"" class=""data"" align=""right"">" & NullCheck(RS_PODetails("Discount")) & "%</td>"
				'	End If
				'    IF PO_REPORT_SHOWTAXABLE = "Yes" Then
				'		rw "<td valign=""top"" class=""data"" align=""center"">" 
				'			Call OutputBit(RS_PODetails("IsTax").Value)
				'		rw "</td>"
				'	End If
				'    If PO_REPORT_SHOWTAX = "Yes" Then
				'		rw "<td nowrap valign=""top"" class=""data"" align=""right"">" & NullCheck(RS_PODetails("TaxRate")) & "%</td>"
				'		rw "<td nowrap valign=""top"" class=""data"" align=""right"">" & CurrencySymbol & " " & FormatNumber(NullCheck(RS_PODetails("TaxAmount"))) & "</td>"
				'	End If
				'	rw "<td nowrap valign=""top"" class=""data"" align=""right"">" & CurrencySymbol & " " & FormatNumber(NullCheck(RS_PODetails("LineItemTotal"))) & "</td>"
							
				rw "</tr>"
                '
					'Add horizontal rule if displaying Description or Comments
					'If (PO_REPORT_ITEMDESC = "Yes" OR PO_REPORT_ITEMCOMMENTS = "Yes") AND 1=0 Then
					'	rw "<tr bgColor=""" & altBgColor & """ >"
					'		rw "<td colspan=""" & numCols & """ style=""height:2px; padding:0px 10px 0px 10px; "">"
					'			rw "<div style=""border-bottom:1px solid gray;""></div>"
					'		rw "</td>" 
					'	rw "</tr>"
					'End If

					'If PO_REPORT_ITEMDESC = "Yes" Then
					'	If Not NullCheck(RS_PODetails("PartDescription")) = "" Then
					'		If PO_REPORT_ITEMCOMMENTS = "Yes" AND Not NullCheck(RS_PODetails("Comments")) = "" Then
					'			rw "<tr bgColor=""" & altBgColor & """><td colspan=""" & numCols & """ class=""data"" style=""padding-left:0.5em;""><b>Description:</b> " & Replace(NullCheck(RS_PODetails("PartDescription")),"%0D%0A"," ") & "</td></tr>"
					'		Else
					'			'add padding bottom for description since comments is not being displayed
					'			rw "<tr bgColor=""" & altBgColor & """><td colspan=""" & numCols & """ class=""data"" style=""padding-left:0.5em; padding-bottom:10px;""><b>Description:</b> " & Replace(NullCheck(RS_PODetails("PartDescription")),"%0D%0A"," ") & "</td></tr>"
					'		End If
					'	End If
					'End If

					'If PO_REPORT_ITEMCOMMENTS = "Yes" Then
					'	If Not NullCheck(RS_PODetails("Comments")) = "" Then
					'		rw "<tr bgColor=""" & altBgColor & """><td colspan=""" & numCols & """ class=""data"" style=""padding-left:0.5em; padding-bottom:10px;""><b>Comments:</b> " & Replace(NullCheck(RS_PODetails("Comments")),"%0D%0A"," ") & "</td></tr>"
					'	End If
					'End If
						
					RS_PODetails.MoveNext
				Loop

			rw "</table>"

			If Not GotPOLineItem Then
				rw "<table cellspacing=""0"" cellpadding=""0"" width=""100%"">"
				'OutputBlankRow(10)
				rw "</table>"
			Else
				rw "<br>"
			End If
            ' WO88677 Remove Bottom table (Bill to, tax, sub-total and total)

			'rw "<table id=""BillTo"" border=""0"" width=""100%"" cellpadding=""0"" cellspacing=""0"">"
			'	rw "<tr>"
			'		rw "<td width=""33%"" valign=""top"">"
			'			If PO_REPORT_SHOWBILLTO <> "Yes" Then
			'				If Not RS_PO("BillToPK") = RS_PO("ShipToPK") Then  
			'					rw "<table border=""0"" height=""100"">"
			'						rw "<tr>"
			'							rw "<td valign=""top"" style=""padding-left:10px"" class=""headingPO"">Bill To:</td>"
			'						rw "</tr>"
			'						rw "<tr>"
			'							rw "<td valign=""top"" style=""padding-left:30px;padding-bottom:10px"" class=""data"">"
			'								rw "<b>" & NullCheck(RS_PO("BillToName")) & "</b><br>"
			'								If PO_REPORT_SHOWPHONE = "Yes" Then
			'								  If NullCheck(RS_PO("BillToAttention")) <> "" Then
			'									  rw "Attn: " & NullCheck(RS_PO("BillToAttention")) & "<br>"
			'								  End If
			'								End If
			'								rw NullCheck(RS_PO("BillToAddress1")) & "<br>"
			'								If Not RS_PO("BillToAddress2") = "" Then
			'								 rw NullCheck(RS_PO("BillToAddress2")) & "<br>"
			'								End If
			'								If Not RS_PO("BillToAddress3") = "" Then
			'								 rw NullCheck(RS_PO("BillToAddress3")) & "<br><br>"
			'								End If
			'								If PO_REPORT_SHOWPHONE = "Yes" Then
			'								  If NullCheck(RS_PO("BillToPhone")) <> "" Then
			'								  rw "Phone: " & NullCheck(RS_PO("BillToPhone"))
			'								  End If
			'								  If NullCheck(RS_PO("BillToFax")) <> "" Then
			'									rw "<Br>Fax: " & NullCheck(RS_PO("BillToFax"))
			'								  End If
			'								End If
			'								If (NullCheck(RS_PO("BillToAttention")) <> "" And PO_REPORT_SHOWPHONE = "No") Then
			'									rw "Attn: " & NullCheck(RS_PO("BillToAttention"))
			'								End If
			'							rw "</td>"
			'						rw "</tr>"
			'					rw "</table>"
			'				End If
			'			Else
			'				rw "<table border=""0"" height=""100"">"
			'					rw "<tr>"
			'						rw "<td valign=""top"" style=""padding-left:10px"" class=""headingPO"">Bill To:</td>"
			'					rw "</tr>"
			'					rw "<tr>"
			'						rw "<td valign=""top"" style=""padding-left:30px;padding-bottom:10px"" class=""data"">"
			'							rw "<b>" & NullCheck(RS_PO("BillToName")) & "</b><br>"
			'							If PO_REPORT_SHOWPHONE = "Yes" Then
			'								If NullCheck(RS_PO("BillToAttention")) <> "" Then
			'									rw "Attn: " & NullCheck(RS_PO("BillToAttention")) & "<br>"
			'								End If
			'							End If
			'							rw NullCheck(RS_PO("BillToAddress1")) & "<br>"
			'							If Not RS_PO("BillToAddress2") = "" Then
			'								rw NullCheck(RS_PO("BillToAddress2")) & "<br>"
			'							End If
			'							If Not RS_PO("BillToAddress3") = "" Then
			'								rw NullCheck(RS_PO("BillToAddress3")) & "<br><br>"
			'							End If
			'							If PO_REPORT_SHOWPHONE = "Yes" Then
			'								If NullCheck(RS_PO("BillToPhone")) <> "" Then
			'								rw "Phone: " & NullCheck(RS_PO("BillToPhone"))
			'								End If
			'								If NullCheck(RS_PO("BillToFax")) <> "" Then
			'								rw "<Br>Fax: " & NullCheck(RS_PO("BillToFax"))
			'								End If
			'							End If
			'							If (NullCheck(RS_PO("BillToAttention")) <> "" And PO_REPORT_SHOWPHONE = "No") Then
			'								rw "Attn: " & NullCheck(RS_PO("BillToAttention"))
			'							End If
			'						rw "</td>"
			'					rw "</tr>"
			'				rw "</table>"
			'			End If
			'			
			'		rw "</td>"
			'		rw "<td width=""33%"" valign=""top"">"
  			'		If PO_REPORT_SHOWBUYERINFO = "Yes" Then
			'			  'rw "<table border=""0"" width=""100%"" bgColor=""#F5F5F5"" style=""border:1px solid"" height=""150"">"
			'			rw "<table border=""0"" height=""100"">"
			'				rw "<tr>"
			'					rw "<td valign=""top"" style=""padding-left:10px;"" class=""headingPO"">Buyer:</td>"
			'				rw "</tr>"
			'				rw "<tr>"
			'					rw "<td valign=""top"" style=""padding-left:30px;padding-bottom:10px"" class=""data"">"
			'						rw "<b>" & NullCheck(RS_PO("BuyerCoName")) & "</b><br>"
			'						If PO_REPORT_SHOWPHONE = "Yes" Then
			'							If NullCheck(RS_PO("BuyerAttn")) <> "" Then
			'							rw "Attn: " & NullCheck(RS_PO("BuyerAttn")) & "<br>"
			'							End If
			'						End If
			'						rw NullCheck(RS_PO("BuyerAddress1")) & "<br>"
			'						If Not RS_PO("BuyerAddress2") = "" Then
			'							rw NullCheck(RS_PO("BuyerAddress2")) & "<br>"
			'						End If
			'						rw NullCheck(RS_PO("BuyerCity")) & "," & NullCheck(RS_PO("BuyerState")) & " " & NullCheck(RS_PO("BuyerZip")) & "<br><br>"
            '
			'						If PO_REPORT_SHOWPHONE = "Yes" Then
			'							If NullCheck(RS_PO("BuyerPhone")) <> "" Then
			'							rw "Phone: " & NullCheck(RS_PO("BuyerPhone"))
			'							End If
			'							If Nullcheck(RS_PO("BuyerFax")) <> "" Then
			'							rw "<br>Fax: " & NullCheck(RS_PO("BuyerFax"))
			'							End If
			'						End If
			'						If (NullCheck(RS_PO("BuyerAttn")) <> "" And PO_REPORT_SHOWPHONE = "No") Then
			'							rw "Attn: " & NullCheck(RS_PO("BuyerAttn"))
			'						End If
			'					rw "</td>"
			'				rw "</tr>"
			'			rw "</table>"
			'		End If
			'		rw "</td>"
			'		rw "<td width=""33%"" valign=""top"" align=""right"">"
			'			rw "<table cellpadding=""2"" cellspacing=""0"" border=""0"" width=""200"" style=""border:1px solid #CCCCCC"" cellpadding=""2"">"
			'				rw "<tr>"
			'					rw "<td class=""data""><b>Subtotal:</b></td>"
			'					rw "<td nowrap class=""data"" align=""right"">" & CurrencySymbol & " " & FormatNumber(NullCheck(RS_PO("Subtotal"))) & "</td>"
			'				rw "</tr>"
			'				If NumericNullCheck(RS_PO("FreightCharge")) > 0 Then
			'				  rw "<tr>"
			'					  rw "<td class=""data""><b>Freight:</b></td>"
			'					  rw "<td nowrap class=""data"" align=""right"">" & CurrencySymbol & " " & FormatNumber(NullCheck(RS_PO("FreightCharge"))) & "</td>"
			'				  rw "</tr>"
			'				End If
			'				rw "<tr>"
			'					rw "<td class=""data""><b>Tax:</b></td>"
			'					rw "<td nowrap class=""data"" align=""right"">" & CurrencySymbol & " " & FormatNumber(NullCheck(RS_PO("TaxAmount"))) & "</td>"
			'				rw "</tr>"
			'			rw "</table>"
			'			rw "<br>"
			'			rw "<table cellpadding=""2"" cellspacing=""0"" border=""0"" width=""200"" style=""border:2px solid RoyalBlue"" cellpadding=""2"">"
			'				rw "<tr>"
			'					rw "<td class=""data""><b>Total:</b></td>"
			'					rw "<td nowrap class=""data"" align=""right"">" & CurrencySymbol & " " & FormatNumber(NullCheck(RS_PO("Total"))) & "</td>"
			'				rw "</tr>"
			'			rw "</table>"
			'		rw "</td>"
			'	rw "</tr>"
			'rw "</table>"
            'WO88677 end

			Call OutputApprovalsBox()
			If PO_REPORT_NOTES = "Yes" Then
				Call OutputPONotes()
			End If
			Call OutputDocumentText(RS_POdocument,False,"PO")

			RS_PO.MoveNext

			' do not output a page break on the last record
			If Not RS_PO.EOF Then
				rw "<P style='page-break-before: always'></p>"
			End If

		Loop
		Else
			rw "<div style=""padding-top:5px; padding-left:5px; font-family:arial; font-size:10pt; color:gray; font-weight:bold;"">"
			rw "(No Records)"
			rw "</div>"
		End If
	End If

	Call OutputSQL()
	Call OutputFooter()
	Call EndFile()

	CloseObj RS_PO

End Sub

Sub SetupPOData()

	If errormessage = "" Then
    ' Set Currency Symbol
    sql = "SELECT CodeIcon FROM LookupTableValues WHERE LookupTable = 'CURRENCY' and CodeName = '"&SY_BASECURRENCY&"'"
    'Response.Write "<textarea rows=6 cols=100>" & sql & "</textarea>"
    'Response.End
    Set rs = db.runSQLReturnRS(sql,"")
    If db.dok Then
      If Not rs.EOF Then
        CurrencySymbol = NullCheck(rs(0))
      Else
        CurrencySymbol = "$"
      End If
    End If

		' We must inner join project because our criteria from the purchase order
		' module or the criteria dialog will prefix all fields as PurchaseOrder.
		POsql = "SELECT DISTINCT PurchaseOrder.POPK FROM PurchaseOrder WITH (NOLOCK) INNER JOIN PurchaseOrderDetail WITH (NOLOCK) ON PurchaseOrderDetail.POPK = PurchaseOrder.POPK " & sql_where

		' We are not using the above technique which only allows us to accept
		' criteria fields from the least common denominator of joined tables
		' which in this case is Purchase Order and Purchase Order Details.

		' GET PURCHASE ORDERS
		sql = "SELECT DISTINCT PurchaseOrder.POPK, PurchaseOrder.POID, PurchaseOrder.POName, PurchaseOrder.AuthLevelsRequired, PurchaseOrder.RepairCenterName, PurchaseOrder.Currency, PurchaseOrder.CurrencySymbol, " &_
		"PurchaseOrder.PriorityDesc, PurchaseOrder.StatusDesc, PurchaseOrder.PODate, PurchaseOrder.Subtotal, PurchaseOrder.FreightCharge, PurchaseOrder.TaxAmount, "&_
		"PurchaseOrder.Total, PurchaseOrder.BillToPK, PurchaseOrder.BillToName, PurchaseOrder.BillToAddress1, PurchaseOrder.BillToAddress2, PurchaseOrder.BillToAttention, "&_
		"PurchaseOrder.BuyerName, PurchaseOrder.BuyerPK, PurchaseOrder.BillToAddress3, PurchaseOrder.ShipToName, PurchaseOrder.ShipToAttention, PurchaseOrder.ShipToPK, PurchaseOrder.ShipToAddress1, "&_
		"PurchaseOrder.ShipToAddress2, PurchaseOrder.ShipToAddress3, PurchaseOrder.Comments, PurchaseOrder.VendorID, PurchaseOrder.Issued, PurchaseOrder.Requested, "&_
		"PurchaseOrder.ShippingMethodDesc, PurchaseOrder.FreightTermsDesc, PurchaseOrder.TermsDesc, Company.CompanyName, Company.Address AS CompanyAddress1, "&_
		"Company.Address2 AS CompanyAddress2, Company.City AS CompanyCity, Company.State AS CompanyState, Company.Zip AS CompanyZip, Company.Attention AS CompanyAttn, "&_
		"Company.FOBPoint, PurchaseOrder.AuthLevelsRequired, PurchaseOrder.IsApproved, PurchaseOrder.PrintedBox, PurchaseOrder.UDFChar1, PurchaseOrder.UDFChar2, "&_
		"PurchaseOrder.UDFChar3, PurchaseOrder.UDFChar4, PurchaseOrder.UDFChar5, PurchaseOrder.UDFDate1, PurchaseOrder.UDFDate2, PurchaseOrder.UDFBit1, PurchaseOrder.UDFBit2, "&_
		"Company.Phone, Company.Fax, ShipToPhone = s.Phone, ShipToFax = s.Fax, BillToPhone = x.Phone, ShipDate, "&_
		"BillToFax = x.Fax, BuyerCoName = b.CompanyName, BuyerAddress1 = b.Address, BuyerAddress2 = b.Address2, BuyerCity = b.City, BuyerState = b.State, BuyerZip = b.Zip, "&_
		"PurchaseOrder.AccountID, PurchaseOrder.AccountName, PurchaseOrder.DepartmentID, PurchaseOrder.DepartmentName, PurchaseOrder.TenantID, PurchaseOrder.TenantName, "&_
		"BuyerPhone = b.Phone, BuyerFax = b.Fax, BuyerAttn = b.Attention, RequestedBy = rq.LaborName, IssuedBy = (SELECT TOP 1 ISNULL(LaborName,'') FROM PurchaseOrderStatusHistory "&_
		"INNER JOIN Labor ON Labor.LaborPK = PurchaseOrderStatusHistory.RowVersionUserPK WHERE POPK = PurchaseOrder.POPK AND Status = 'ISSUED' ORDER BY StatusDate DESC), " &_
		"ApprovedBy = (SELECT TOP 1 ISNULL(LaborName,'') FROM PurchaseOrderStatusHistory INNER JOIN Labor ON Labor.LaborPK = PurchaseOrderStatusHistory.RowVersionUserPK WHERE "&_
		"POPK = PurchaseOrder.POPK AND Status = 'APPROVED' ORDER BY StatusDate DESC), IsApproved, Approved = (SELECT TOP 1 ISNULL(StatusDate,'') FROM PurchaseOrderStatusHistory "&_
		"WHERE POPK = PurchaseOrder.POPK AND Status = 'APPROVED' ORDER BY StatusDate DESC) " &_
			  "FROM PurchaseOrder WITH (NOLOCK) LEFT OUTER JOIN " +_
			  "PurchaseOrderDetail ON PurchaseOrderDetail.POPK = PurchaseOrder.POPK " & _
			  "LEFT OUTER JOIN Company ON Company.CompanyPK = PurchaseOrder.VendorPK "&_
			  "LEFT JOIN Company b ON b.CompanyPK = PurchaseOrder.BuyerCompanyPK " &_
			  "LEFT JOIN Company s ON s.CompanyPK = PurchaseOrder.ShipToPK " &_
			  "LEFT JOIN Company x ON x.CompanyPK = PurchaseOrder.BillToPK " &_
			  "LEFT JOIN Labor rq ON rq.LaborPK = PurchaseOrder.RequesterPK " & sql_where & _
			  "ORDER BY PurchaseOrder.POPK Desc "
    'Response.Write "<textarea rows=6 cols=100>" & sql & "</textarea>"
    'Response.End
		Set RS_PO = db.runSQLReturnRS(sql,"")

		' GET PURCHASE ORDER DETAILS
		sql = "SELECT PurchaseOrderDetail.POPK, PurchaseOrderDetail.PartID, PurchaseOrderDetail.PartName, VendorPartNumber, OrderUnitPrice, OrderUnitQty, PurchaseOrderDetail.OrderUnitsDesc, LineItemNo, PurchaseOrderDetail.DirectIssue, "&_
		"PurchaseOrderDetail.Discount, LineItemTotal, PurchaseOrderDetail.AccountName, PurchaseOrderDetail.SubAccountName, PurchaseOrderDetail.WOPK, PurchaseOrderDetail.WOID, "&_
		"PurchaseOrderDetail.LocationID, PurchaseOrderDetail.LocationName, PartLocation.Bin, PurchaseOrderDetail.ConversionToIssueUnits, PartLocation.IssueUnitsDesc, " &_
		"PurchaseOrderDetail.IsTax, PurchaseOrderDetail.TaxRate, PurchaseOrderDetail.TaxAmount, Part.PartDescription, PurchaseOrderDetail.Comments " &_
		"FROM PurchaseOrderDetail WITH (NOLOCK) " &_
		"INNER JOIN PurchaseOrder WITH (NOLOCK) ON PurchaseOrder.POPK = PurchaseOrderDetail.POPK " &_
		"LEFT JOIN PartLocation ON PartLocation.LocationPK = PurchaseOrderDetail.LocationPK AND PartLocation.PartPK = PurchaseOrderDetail.PartPK AND PartLocation.Bin = PurchaseOrderDetail.Bin " &_
		"INNER JOIN Part ON Part.PartPK = PurchaseOrderDetail.PartPK " &_
		sql_where & _
		"ORDER BY PurchaseOrder.POPK Desc, PurchaseOrderDetail.LineItemNo "
    'Response.Write "<textarea rows=6 cols=100>" & sql & "</textarea>"
    'Response.End

		Set RS_PODetails = db.runSQLReturnRS(sql,"")

		' Get Purchase Order Approval
		'sql = "SELECT POPK, StatusDate, StatusDesc, PurchaseOrderStatusHistory.RowVersionInitials, Labor.LaborName " +_
		'	  "FROM PurchaseOrderStatusHistory WITH (NOLOCK) INNER JOIN Labor ON Labor.LaborPK = PurchaseOrderStatusHistory.RowVersionUserPK "
		'	  If Not sql_where = "" Then
		'		sql = sql & "WHERE PurchaseOrderStatusHistory.IsAuthStatus = 1 and (POPK in (" & POsql & ")) "
		'	  Else
		'		sql = sql & "WHERE PurchaseOrderStatusHistory.IsAuthStatus = 1 "
		'	  End If
		'	  If PO_REPORT_SHOWOPTIONALAPPROVAL = "ENHANCED" Then
		'	    sql = sql & " AND Status = 'APPROVED' "
		'	  End If
		'	  sql = sql & _
		'	  "ORDER BY PurchaseOrderStatusHistory.POPK Desc, PurchaseOrderStatusHistory.StatusDate"
    'Response.Write "<textarea rows=6 cols=100>" & sql & "</textarea>"
    'Response.End

		'Set RS_POApproval = db.runSQLReturnRS(sql,"")

		' GET PURCHASE ORDER DOCUMENTS
		sql = _
		"SELECT       md.PK, md.POPK, d.LocationType, d.LocationTypeDesc, d.DocumentPK, d.DocumentID, d.DocumentName, md.ModuleID, d.DocumentTypeDesc,  " & _
		"                      d.Location, md.PrintWithWO, md.SendWithEmail, md.RowVersionDate, d.Photo, " & _
		"                      MCModule.TitleforDocumentList, d.DocumentText " & _
		"FROM         PurchaseOrderdocument md WITH (NOLOCK) LEFT OUTER JOIN " & _
		"                      Document d WITH (NOLOCK) ON md.DocumentPK = d.DocumentPK INNER JOIN " & _
		"                      MCModule WITH (NOLOCK) ON md.ModuleID = MCModule.ModuleID " & _
		"WHERE "
		If Not sql_where = "" Then
		  sql = sql & "(md.POPK in (" & POsql & ")) AND "
		End If
		sql = sql & _
		"(md.PrintWithWO = 1 OR md.SendWithEmail = 1) " & _
		" ORDER BY md.POPK Desc, md.ModuleID, d.DocumentID "

		Set RS_PODocument = db.runSQLReturnRS(sql,"")

		If Trim(Request.QueryString("sqlwhere")) = "" Then
			Call dok_check_afterflush(db,"Report Message","There was a problem building the report. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")
		Else
			'rw sql
			'Response.Write db.derror
			'Response.End
			Call dok_check_afterflush_noinfo(db,"Report Message","You have chosen to run a report that is not compatible with the current module criteria. You can either choose a different report, or you can run the selected report from the Maintenance Reporter applicaiton by clicking the Reports button on the toolbar.")
		End If

	End If

End Sub

Sub OutputPOHeaderRight()
	rw "<div style=""font-family:Arial;font-size:16px;color:#333333;font-weight:bold;margin-bottom:4px;"">" & _
	   "Purchase Order " & POID & _
		"</div>"
	rw "<div style=""font-family:Arial;font-size:11px;font-weight:normal"">" & _
		NullCheck(RS_PO("RepairCenterName")) & _
		"</div>"
    rw_fileonly "<div style=""font-family:Arial;font-size:11px;font-weight:normal"">" & _
       "Sent " & CStr(Date()) & "&nbsp;-&nbsp;" & CStr(TimeNullCheckAT(Time())) & "&nbsp;"
        If BitNullCheck(RS_PO("PrintedBox")) Then
        rw_fileonly "(Duplicate Copy)"
         End If
        rw_fileonly "</div>"
    If Not FromAgent Then
        Response.Write "<div style=""font-family:Arial;font-size:11px;font-weight:normal"">" & _
           "Printed " & CStr(Date()) & "&nbsp;-&nbsp;" & CStr(TimeNullCheckAT(Time())) & "&nbsp;"
            If BitNullCheck(RS_PO("PrintedBox")) Then
            Response.Write "(Duplicate Copy)"
            End If
            Response.Write "</div>"
    End If
End Sub

Sub OutputCurrency(mny)
	'If Not IsNull(mny) Then
		rw "$" & FormatNumber(mny)
	'Else
	'	rw "$" & NullCheck(mny)
	'End If
End Sub

Sub OutputApprovalsBox()
  If PO_REPORT_SHOWOPTIONALAPPROVAL <> "NONE" Then
	  ' Get Purchase Order Approval
    If PO_REPORT_SHOWOPTIONALAPPROVAL = "DEFAULT" Then
		  sql = "SELECT POPK, StatusDate, StatusDesc, PurchaseOrderStatusHistory.RowVersionInitials, Labor.LaborName " +_
		  	  "FROM PurchaseOrderStatusHistory WITH (NOLOCK) INNER JOIN Labor ON Labor.LaborPK = PurchaseOrderStatusHistory.RowVersionUserPK "
		  	  If Not sql_where = "" Then
		  		sql = sql & "WHERE PurchaseOrderStatusHistory.IsAuthStatus = 1 and POPK = " & POPK
		  	  Else
		  		sql = sql & "WHERE PurchaseOrderStatusHistory.IsAuthStatus = 1 "
		  	  End If
		  	  sql = sql & _
		  	  "ORDER BY PurchaseOrderStatusHistory.POPK Desc, PurchaseOrderStatusHistory.StatusDate"
    Else
	    sql = "SELECT a.*, Labor.Initials, Labor.LaborName FROM dbo.MCUDF_GetPOApprovals("&POPK&") a INNER JOIN Labor ON Labor.LaborPK = a.LAPK"
    End If
    'Response.Write "<textarea rows=6 cols=100>" & sql & "</textarea>"
    'Response.End
	  Set RS_POApproval = db.runSQLReturnRS(sql,"")

	  Dim BlankRowNum

	  rw "<fieldset style=""padding-top:14px"">"
		  rw "<legend style=""font-size:9pt;"" class=""legendHeader"">Approvals</legend>"
		  rw "<table style=""margin-top:5px;"" border=""0"" cellspacing=""3"" cellpadding=""0"" width=""98%"" align=""center"">"

	    If PO_REPORT_SHOWOPTIONALAPPROVAL = "STATIC" Then
        Call OutputStaticApporvalHeader()
        Dim cnt
        cnt=0
        Do While cnt < ApprovalCount
			    Call OutputBlankRow(1,2)
			    cnt = cnt+1
			  Loop
	    ElseIf PO_REPORT_SHOWOPTIONALAPPROVAL = "ENHANCED" Then
        Call OutputEnhancedApprovalHeader()
			  ' Approval data
			  If Not RS_POApproval.EOF Then
				  Do While NullCheck(RS_POApproval("POPK")) = NullCheck(POPK)
					  Call OutputEnhancedApproval(RS_POApproval)
					  RS_POApproval.MoveNext
				  Loop
			  Else
	        rw "<tr class=""blank_row"">"
		        rw "<td valign=""bottom"" class=""data_underline"" align=""left"">&nbsp;</td>"
		        rw "<td valign=""top"" class=""data_underline"">&nbsp;</td>"
		        rw "<td valign=""top"" class=""data_underline"">&nbsp;</td>"
	        rw "</tr>"
			  End If
	    Else  'DEFAULT
			  Call OutputApprovalHeader()
			  ' Approval data
			  If Not RS_POApproval.EOF Then
				  Do While NullCheck(RS_POApproval("POPK")) = NullCheck(POPK)
					  Call OutputApproval(RS_POApproval)
					  RS_POApproval.MoveNext
				  Loop
			  End If
	    End If

		  rw "</table>"
		  rw "<br>"
	  rw "</fieldset>"
  End If
End Sub

Sub OutputApprovalHeader()
	rw "<tr>"
		rw "<td class=""labels"" width=""100"">Date</td>"
		rw "<td class=""labels"">Description</td>"
		rw "<td class=""labels"">Initials</td>"
	rw "</tr>"
End Sub

Sub OutputEnhancedApprovalHeader
	rw "<tr>"
		rw "<td class=""labels"" width=""330"">Name</td>"
		rw "<td class=""labels"" width=""190"">Date</td>"
		rw "<td class=""labels"">Description</td>"
	rw "</tr>"
End Sub

Sub OutputStaticApporvalHeader()
	rw "<tr>"
		rw "<td class=""labels"" style="""">Signature / Name</td>"
		rw "<td class=""labels"" width=""200"">Approval Date</td>"
	rw "</tr>"
End Sub

Sub OutputApproval(rs)
	rw "<tr>"
		rw "<td valign=""top"" class=""data_underline"">" & DateTimeNullCheckAT(rs("StatusDate")) & "&nbsp;</td>"
		If Not UCase(NullCheck(rs("StatusDesc"))) = "(NOT REQUIRED)" and Not UCase(NullCheck(rs("StatusDesc"))) = "NOT REQUIRED" Then
		rw "<td valign=""bottom"" class=""data_underline"" align=""left"">" & NullCheck(rs("StatusDesc")) & "&nbsp;</td>"
		Else
		rw "<td valign=""bottom"" class=""data_underline"" align=""left"">&nbsp;</td>"
		End If
		rw "<td valign=""bottom"" class=""data_underline"" align=""left"">" & NullCheck(rs("RowVersionInitials")) & "&nbsp;</td>"
	rw "</tr>"
End Sub

Sub OutputEnhancedApproval(rs)
  rw "<tr>"
	  rw "<td valign=""bottom"" class=""data_underline"" align=""left"">" & NullCheck(rs("LaborName")) & "&nbsp;</td>"
	  rw "<td valign=""top"" class=""data_underline"">" & NullCheck(rs("StatusDate")) & "&nbsp;</td>"
	  rw "<td valign=""top"" class=""data_underline"">" & NullCheck(rs("POStatus")) & "&nbsp;</td>"
  rw "</tr>"
End Sub

Sub OutputPONotes()
'date, initials, note
    sql = "SELECT POPK, NoteDate, Initials, Note FROM PurchaseOrderNote WITH (NOLOCK) WHERE PurchaseOrderNote.POPK = " & NullCheck(POPK) & " ORDER BY NoteDate "

    'Response.Write "<textarea rows=6 cols=100>" & sql & "</textarea>"
    'Response.End
	Set RS_PONote = db.runSQLReturnRS(sql,"")

	rw "<fieldset style=""padding-top:14px"">"
		rw "<legend style=""font-size:9pt;"" class=""legendHeader"">Notes</legend>"
		rw "<table style=""margin-top:5px;"" border=""0"" cellspacing=""3"" cellpadding=""0"" width=""98%"" align=""center"">"
			rw "<tr>"
				rw "<td class=""labels"" width=""15%"">Date</td>"
				rw "<td class=""labels"" width=""10%"">Initials</td>"
				rw "<td class=""labels"">Note</td>"
			rw "</tr>"
			If Not RS_PONote.EOF Then
				Do While NullCheck(RS_PONote("POPK")) = NullCheck(POPK)
					rw "<tr>"
						rw "<td valign=""bottom"" class=""data_underline"">" & DateNullCheckAT(RS_PONote("NoteDate")) & "&nbsp;</td>"
						rw "<td valign=""bottom"" class=""data_underline"" align=""left"">" & NullCheck(RS_PONote("Initials")) & "&nbsp;</td>"
						rw "<td valign=""bottom"" class=""data_underline"">" &  Replace(NullCheck(RS_PONote("Note")),"%0D%0A"," ") & "&nbsp;</td>"
					rw "</tr>"
					RS_PONote.MoveNext
				Loop
			Else
				Call OutputBlankRow(1,3)
			End If
		rw "</table>"
		rw "<br>"
	rw "</fieldset>"

End Sub

Sub OutputDocumentText(rs,nopocheck,modid)

	If Not rs.EOF Then
		If NullCheck(rs("POPK")) = NullCheck(POPK) or nopocheck Then

			Do While Not rs.eof and (NullCheck(rs("POPK")) = NullCheck(POPK) or nopocheck)

				' FOR NOW WE ARE ONLY PRINTING LIBRARY DOCS ON POs! MC7
				If UCase(Trim(rs("LocationType"))) = "LIBRARY" Then

					If rs("PrintWithWO") Then
						Response.Write "<P style='page-break-before: always'></p>"
						'Response.Write "<fieldset>"
						'Response.Write "<legend class=""legendHeader"">Purchase Order " & NullCheck(RS_PO("POID")) & " " & NullCheck(rs("DocumentTypeDesc")) & "</legend>"
						'Response.Write "<div style=""margin-top:1px; padding-left:10px; padding-right:10px; padding-bottom:10px;"">"
					End If

					If rs("SendWithEmail") Then
						rw_fileonly "<P style='page-break-before: always'></p>"
						'rw_fileonly "<fieldset>"
						'rw_fileonly "<legend class=""legendHeader"">Purchase Order " & NullCheck(RS_PO("POID")) & " " & NullCheck(rs("DocumentTypeDesc")) & "</legend>"
						'rw_fileonly "<div style=""margin-top:1px; padding-left:10px; padding-right:10px; padding-bottom:10px;"">"
					End If

					Dim remotehtml, localhtml

					Select Case UCase(Trim(rs("LocationType")))
					Case "LIBRARY"
						If rs("PrintWithWO") Then
							Response.Write rs("DocumentText")
						End If
						If rs("SendWithEmail") Then
							rw_fileonly rs("DocumentText")
						End If
					Case "HTTP", "HTTPLIBRARY"
						If Not NullCheck(rs("Location")) = "" Then
							If rs("PrintWithWO") or rs("SendWithEmail") Then
								remotehtml = GrabAndFixImages(FixLibraryLinkURL(rs("Location").value))
								If rs("PrintWithWO") Then
									Response.Write remotehtml
								End If
								If rs("SendWithEmail") Then
									rw_fileonly remotehtml
								End If
							End If
						End If
					Case "FILE"
						If Not NullCheck(rs("Location")) = "" Then
							If Instr(UCase(rs("Location")),".HTM") > 0 Then
								If rs("PrintWithWO") or rs("SendWithEmail") Then
									localhtml = GrabAndFixImages(FixLibraryLinkURL(rs("Location").value))
									If rs("PrintWithWO") Then
										Response.Write localhtml
									End If
									If rs("SendWithEmail") Then
										rw_fileonly localhtml
									End If
								End If
							End If
						End If
					End Select

					If rs("PrintWithWO") Then
						'Response.Write "</div>"
						'Response.Write "</fieldset>"
					End If

					If rs("SendWithEmail") Then
						'rw_fileonly "</div>"
						'rw_fileonly "</fieldset>"
					End If
				
				Else
					' MC7 Let's output URL Images
					If UCase(Trim(rs("LocationType"))) = "HTTP" or UCase(Trim(rs("LocationType"))) = "HTTPLIBRARY" Then
						Dim FileExt, docloc
						FileExt = UCase(Trim(GetFileExt(rs("Location"))))
						If FileExt = "JPG" or FileExt = "WMF" or FileExt = "GIF" or FileExt = "PNG" Then

                            docloc = NullCheck(RS("Location"))
                            If UCase(Trim(rs("LocationType"))) = "HTTPLIBRARY" Then                                
                                If Not InStr(1,docloc,"http://",1) > 0 and Not InStr(1,docloc,"https://",1) > 0 Then
                                    docloc = LCase(GetSession("webHTTP") & GetWebServer() & Application("imageserver") & GetSession("DB")) & "/fileStore/" & docloc
                                End If
                            End If

                            remotehtml = "<div style=""width:100%; overflow:auto;""><a target=""_blank"" href=""" & Trim(FixLibraryLinkURL(docloc)) & """><img style=""height:500px;"" src=""" & Trim(docloc) & """ border=""0""/></a></div>"
							If rs("PrintWithWO") Then
								Response.Write "<P style='page-break-before: always'></p>"
								Response.Write remotehtml
							End If
							If rs("SendWithEmail") Then
								rw_fileonly "<P style='page-break-before: always'></p>"
								rw_fileonly remotehtml
							End If
						End If
					End If
				End If

				rs.MoveNext

			Loop

		End If
	End If

End Sub

Sub SetupPOBarcode()

	' If Bar Coding is turned on for either Purchase Order # or Inventory Item #
	' then get the rest of the Barcode Preferences
	If PO_Barcode_POID or PO_Barcode_PartID Then
		Call SetupBarcodePreferences()
	End If

End Sub

Sub SetPrintedFlag()

	If Trim(UCase(Request("setprintedflag"))) = "Y" Then

		' UPDATE PURCHASE ORDER PRINTBOX FLAG
		sql = "UPDATE PurchaseOrder SET PrintedBox = 1, RowVersionAction = 'EDIT', RowVersionDate = getDate() FROM PurchaseOrder " & sql_where & " "
		Call db.runSQL(sql,"")

		If Trim(UCase(Request("EmailReport"))) = "Y" Then
			Call dok_check_afterflush(db,"Report Message","Setting Printed / Emailed failed. Please report this to your Maintenance Manager.")
			Exit Sub
		Else
			Call dok_check(db,"Report Message","Setting Printed / Emailed failed. Please report this to your Maintenance Manager.")
		End If

		donocacheheader
		Response.Write("<!DOCTYPE html>")
		Response.Write("<html dir=""ltr"" lang=""en-us"">")
		Response.Write("<head>")
		Response.Write("<meta content=""IE=edge"" http-equiv=""X-UA-Compatible"" />")
		Response.Write("<meta content=""text/html;charset=UTF-8"" http-equiv=""Content-type"" />")
		Response.Write("<meta content=""text/javascript"" http-equiv=""content-script-type"" />")
		Response.Write("<meta content=""text/css"" http-equiv=""content-style-type"" />")
		Response.Write("<meta content=""favorite"" name=""save"">")
		Response.Write("<meta content=""noindex"" name=""robots"">")
		Response.Write("<meta content=""no"" name=""allow-search"" />")
		Response.Write("<meta content=""yes"" name=""apple-mobile-web-app-capable"" />")
		Response.Write("<meta content=""black"" name=""apple-mobile-web-app-status-bar-style"" />")
		Response.Write("<meta HTTP-EQUIV=""Expires"" content=""-1"">")+nl
		Response.Write("<title></title>")+nl
		Response.Write("<script language=""JavaScript"">")+nl
		Response.Write("try {top.refreshcurrentrecord();} catch(e) {};")+nl
		'Response.Write("top.playsound('sounds/done.wav');")+nl
		Response.Write("</script>")+nl
		Response.Write("</head>")+nl
		Response.Write("<body>")+nl
		FlushItNoStore
		If Application("ASPDEBUG") then
			aspdebug
		End If
		Response.Write("</body>")+nl
		Response.Write("</html>")+nl
		Response.End

	End If

End Sub

Sub GetPOReportPrefs()
	
	Dim CSVPreferenceList, RS_POPref

	CSVPreferenceList = _
	"SY_BASECURRENCY" &_
	",PO_REPORT_LINEHEIGHT" &_
	",PO_REPORT_POIDUDF" &_
	",PO_REPORT_SHOWPHONE" &_
	",PO_REPORT_SHOWENSTATUS" &_
	",PO_REPORT_SHOWSHIPRECEIVING" &_
	",PO_REPORT_SHOWLIC2IU" &_
	",PO_REPORT_SHOWLIAC" &_
	",PO_REPORT_SHOWLISAC" &_
	",PO_REPORT_SHOWLIDISC" &_
	",PO_REPORT_SHOWLILOC" &_
	",PO_REPORT_SHOWOPTIONALAPPROVAL" &_
	",PO_REPORT_SHOWBUYERINFO" &_
	",PO_REPORT_SHOWENDETAIL" &_
	",PO_REPORT_UDF_DISPLAY_1" &_
	",PO_REPORT_UDF_DISPLAY_2" &_
	",PO_REPORT_UDF_DISPLAY_3" &_
	",PO_REPORT_UDF_DISPLAY_4" &_
	",PO_REPORT_UDF_DISPLAY_5" &_
	",PO_REPORT_SHOWBILLTO" &_
	",PO_REPORT_SHOWTAXABLE" &_
	",PO_REPORT_SHOWTAX" &_
	",PO_REPORT_ITEMDESC" &_
	",PO_REPORT_ITEMCOMMENTS" &_
	",PO_REPORT_SHOWSHIPDATE" &_
	",PO_REPORT_NOTES" &_
	",PO_Barcode_POID" &_
	",PO_Barcode_PartID" &_
	",PO_BarcodeFormat_POID" &_
	",IN_BarcodeFormat_PartID"

	Set RS_POPref = db.runSPReturnRS("MC_GetMultiplePreferences",Array(Array("@LaborPK", adInteger, adParamInput, 4, GetSession("USERPK")),Array("@RepairCenterPK", adInteger, adParamInput, 4, GetSession("RCPK")),Array("@CSVPreferenceList", adVarChar, adParamInput, 4000, CSVPreferenceList)),"")
	Call dok_check_afterflush(db,"Report Message","There was a problem retrieving the Work Order Preferences. The details of the problem are described below. You can try again but if this message continues to appear, you may want to exit the Maintenance Connection and try again later.")

	If db.dok Then
		Do While Not RS_POPref.Eof
			Select Case UCASE(RS_POPref("PreferenceName"))
				Case "SY_BASECURRENCY"
					SY_BASECURRENCY = RS_POPref("PreferenceValue")
				Case "PO_REPORT_LINEHEIGHT"
					PO_REPORT_LINEHEIGHT = RS_POPref("PreferenceValue")
				Case "PO_REPORT_POIDUDF"
					PO_REPORT_POIDUDF = RS_POPref("PreferenceValue")
				Case "PO_REPORT_SHOWPHONE"
					PO_REPORT_SHOWPHONE = RS_POPref("PreferenceValue")
				Case "PO_REPORT_SHOWENSTATUS"
					PO_REPORT_SHOWENSTATUS = RS_POPref("PreferenceValue")
				Case "PO_REPORT_SHOWSHIPRECEIVING"
					PO_REPORT_SHOWSHIPRECEIVING = RS_POPref("PreferenceValue")
				Case "PO_REPORT_SHOWLIC2IU"
					PO_REPORT_SHOWLIC2IU = RS_POPref("PreferenceValue")
				Case "PO_REPORT_SHOWLIAC"
					PO_REPORT_SHOWLIAC = RS_POPref("PreferenceValue")
				Case "PO_REPORT_SHOWLISAC"
					PO_REPORT_SHOWLISAC = RS_POPref("PreferenceValue")
				Case "PO_REPORT_SHOWLIDISC"
					PO_REPORT_SHOWLIDISC = RS_POPref("PreferenceValue")
				Case "PO_REPORT_SHOWLILOC"
					PO_REPORT_SHOWLILOC = RS_POPref("PreferenceValue")
				Case "PO_REPORT_SHOWOPTIONALAPPROVAL"
					PO_REPORT_SHOWOPTIONALAPPROVAL = RS_POPref("PreferenceValue")
				Case "PO_REPORT_SHOWBUYERINFO"
					PO_REPORT_SHOWBUYERINFO = RS_POPref("PreferenceValue")
				Case "PO_REPORT_SHOWENDETAIL"
					PO_REPORT_SHOWENDETAIL = RS_POPref("PreferenceValue")
				Case "PO_REPORT_UDF_DISPLAY_1"
					PO_REPORT_UDF_DISPLAY_1 = RS_POPref("PreferenceValue")
				Case "PO_REPORT_UDF_DISPLAY_2"
					PO_REPORT_UDF_DISPLAY_2 = RS_POPref("PreferenceValue")
				Case "PO_REPORT_UDF_DISPLAY_3"
					PO_REPORT_UDF_DISPLAY_3 = RS_POPref("PreferenceValue")
				Case "PO_REPORT_UDF_DISPLAY_4"
					PO_REPORT_UDF_DISPLAY_4 = RS_POPref("PreferenceValue")
				Case "PO_REPORT_UDF_DISPLAY_5"
					PO_REPORT_UDF_DISPLAY_5 = RS_POPref("PreferenceValue")
				Case "PO_REPORT_SHOWBILLTO"
					PO_REPORT_SHOWBILLTO = RS_POPref("PreferenceValue")
				Case "PO_REPORT_SHOWTAXABLE"
					PO_REPORT_SHOWTAXABLE = RS_POPref("PreferenceValue")
				Case "PO_REPORT_SHOWTAX"
					PO_REPORT_SHOWTAX = RS_POPref("PreferenceValue")
				Case "PO_REPORT_ITEMDESC"
					PO_REPORT_ITEMDESC = RS_POPref("PreferenceValue")
				Case "PO_REPORT_ITEMCOMMENTS"
					PO_REPORT_ITEMCOMMENTS = RS_POPref("PreferenceValue")
				Case "PO_REPORT_SHOWSHIPDATE"
					PO_REPORT_SHOWSHIPDATE = RS_POPref("PreferenceValue")
				Case "PO_REPORT_NOTES"
					PO_REPORT_NOTES = RS_POPref("PreferenceValue")
				Case "PO_BARCODE_POID"
					PO_Barcode_POID = RS_POPref("PreferenceValue")
				Case "PO_BARCODE_PARTID"
					PO_Barcode_PartID = RS_POPref("PreferenceValue")
				Case "PO_BARCODEFORMAT_POID"
					PO_BarcodeFormat_POID = RS_POPref("PreferenceValue")
				Case "IN_BARCODEFORMAT_PARTID"
					IN_BarcodeFormat_PartID = RS_POPref("PreferenceValue")
			End Select
		RS_POPref.MoveNext
		Loop
	End If

	CloseObj RS_POPref

	' Is Barcoding Turned On For Purchase Order #?
	If UCase(NullCheck(PO_Barcode_POID)) = "YES" Then
		PO_Barcode_POID = True
	Else
		PO_Barcode_POID = False
	End If

	' Is Barcoding Turned On For Inventory Item #?
	If UCase(NullCheck(PO_Barcode_PartID)) = "YES" Then
		PO_Barcode_PartID = True
	Else
		PO_Barcode_PartID = False
	End If

  ' Default Currency
  If NullCheck(SY_BASECURRENCY) = "" Then
    SY_BASECURRENCY = "USD"
  End If

  ' Line Height preference
  If NullCheck(PO_REPORT_LINEHEIGHT) = "" Then
    PO_REPORT_LINEHEIGHT = "SMALL"
  End If

  ' Use UDF field as POID - Value Example: UDFChar1
  If NullCheck(PO_REPORT_POIDUDF) = "" Then
    PO_REPORT_POIDUDF = "NONE"
  End If

  ' Show Phone and Fax Numbers
  If NullCheck(PO_REPORT_SHOWPHONE) = "" Then
    PO_REPORT_SHOWPHONE = "No"
  End If

  ' Show Enhanced Status instead of Basic Status
  If NullCheck(PO_REPORT_SHOWENSTATUS) = "" Then
    PO_REPORT_SHOWENSTATUS = "No"
  End If

  ' Show Shipping and Receiveing Information
  If NullCheck(PO_REPORT_SHOWSHIPRECEIVING) = "" Then
    PO_REPORT_SHOWSHIPRECEIVING = "Yes"
  End If

  ' Line Items - Show Issue Units column
  If NullCheck(PO_REPORT_SHOWLIC2IU) = "" Then
    PO_REPORT_SHOWLIC2IU = "No"
  End If

 ' Line Items - Show Account Column
  If NullCheck(PO_REPORT_SHOWLIAC) = "" Then
    PO_REPORT_SHOWLIAC = "Yes"
  End If

  ' Line Item - Show Sub Account Column
  If NullCheck(PO_REPORT_SHOWLISAC) = "" Then
    PO_REPORT_SHOWLISAC = "No"
  End If

  ' Line Items - Show Discount Column
  If NullCheck(PO_REPORT_SHOWLIDISC) = "" Then
    PO_REPORT_SHOWLIDISC = "Yes"
  End If

  ' Line Items - Show Location column
  If NullCheck(PO_REPORT_SHOWLILOC) = "" Then
    PO_REPORT_SHOWLILOC = "No"
  End If

  ' Show Static Approval in place of basic approval
  If NullCheck(PO_REPORT_SHOWOPTIONALAPPROVAL) = "" Then
    PO_REPORT_SHOWOPTIONALAPPROVAL = "DEFAULT"
  End If

  ' Show Buyer Information Section
  If NullCheck(PO_REPORT_SHOWBUYERINFO) = "" Then
    PO_REPORT_SHOWBUYERINFO = "No"
  End If

  ' Show Additional Detail Information
  If NullCheck(PO_REPORT_SHOWENDETAIL) = "" Then
    PO_REPORT_SHOWENDETAIL = "No"
  End If

	'PO_REPORT_SHOWBILLTO="Yes"
	If NullCheck(PO_REPORT_SHOWBILLTO) = "" Then
	  PO_REPORT_SHOWBILLTO = "No"
	End If

	If NullCheck(PO_REPORT_SHOWTAXABLE) = "" Then
	  PO_REPORT_SHOWTAXABLE = "No"
	End If

	'PO_REPORT_SHOWTAX="Yes"
	If NullCheck(PO_REPORT_SHOWTAX) = "" Then
	  PO_REPORT_SHOWTAX = "No"
	End If

	'PO_REPORT_ITEMDESC="Yes"
	If NullCheck(PO_REPORT_ITEMDESC) = "" Then
	  PO_REPORT_ITEMDESC = "No"
	End If
		
	'PO_REPORT_ITEMCOMMENTS="Yes"
	If NullCheck(PO_REPORT_ITEMCOMMENTS) = "" Then
	  PO_REPORT_ITEMCOMMENTS = "No"
	End If

	'PO_REPORT_SHOWSHIPDATE="Yes"
	If NullCheck(PO_REPORT_SHOWSHIPDATE) = "" Then
	  PO_REPORT_SHOWSHIPDATE = "No"
	End If

	'PO_REPORT_NOTES="Yes"
	If NullCheck(PO_REPORT_NOTES) = "" Then
	  PO_REPORT_NOTES = "No"
	End If

End Sub

Sub OutputBlankRow(r,c)

	Dim row,col

	For row = 1 to r
		rw "<tr class=""blank_row"">"
		For col = 1 to c
			rw "<td class='data_underline'>&nbsp;</td>"
		Next
		rw "</tr>"
	Next

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

	rw "<HTML>"
	rw "<HEAD>"
	rw "<TITLE>" & reportName & "</TITLE>"
	rw "<META http-equiv=Content-Type content=""text/html; charset=utf-8"">"
	If Not FromAgent Then
		If Not Request.QueryString("ExportReportOutputType") = "HTM" Then
			If Application("SCRIPTENCODE") Then %>
			<script LANGUAGE="JScript.Encode" SRC="../../javascript/encode/mc_rpt_common.jse"></script>
			<% Else %>
			<script type="text/javascript" SRC="../../javascript/normal/mc_rpt_common.js"></script>
            <script type="text/javascript" SRC="../../javascript/normal/mc_flashobject.js"></script>

			<% End If %>

			<script type="text/javascript">
				//window.onerror = top.doError;


				var reportid = "<% =reportid %>";
				var reportname = "<% =reportName %>";
				var custom = <% =LCase(custom) %>;
				var reportcopy = <% =LCase(reportcopy) %>;
				var showSQL = "<% =Request.QueryString("showSQL") %>";
				var reporthasfields = <% =LCase(reporthasfields) %>;

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
				    try {
				    <% If SCDefault = "S" and SCAvailable Then %>
		            parent.togglesmartcritpane(reportid,'ON',true);
				    <% Else %>
				    <% If SCDefault = "H" and SCAvailable Then %>
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
					self.focus();
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
				}
			</script>

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
	  rw "    height: "
	  If PO_REPORT_LINEHEIGHT = "MEDIUM" Then
	    rw "24"
	  ElseIf PO_REPORT_LINEHEIGHT = "LARGE" Then
	    rw "30"
	  End If
	  rw "px;"
	  rw "  }"
	  rw "</style>"
	End If

	rw "</HEAD>"
End Sub

%>