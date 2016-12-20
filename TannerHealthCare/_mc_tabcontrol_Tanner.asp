<%
Class TabControl
	Private TabList()
	Private TabHeight
	Private TabWidth
	Private BGColor
	Private bStatic
	Private sSelected
	Private TCClass
	Private HeightSubtract
	Private MainDivStyle
	Public Debug
	Public ControlName
    '''''Init "TabControl2", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
	Public Sub Init (sControlName, Width, Height, BackColor, sSelectedTab, bAmIStatic, sTCClass, nHeightSubtract, sMainDivStyle)
		BGColor=BackColor
		TabHeight=Height
		TabWidth=Width
		sSelected = sSelectedTab
		bStatic = bAmIStatic
		Redim TabList(4,0)
		Debug = 0
		ControlName = sControlName
		TCClass = sTCClass
		HeightSubtract = nHeightSubtract
		MainDivStyle = sMainDivStyle
	End Sub

	Public Sub Build (sTabGroup,sSelectedTab)
		Select Case sTabGroup
		Case "AS_HISTORY_STATIC"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, True, "TabContents", 15, ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Projects","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Predictive Problems","location='staticTab.asp'","../../images/icons/predictive_g.gif", ""
			AddTab "Location Changes","location='staticTab.asp'","../../images/icons/room1_g.gif", ""
			AddTab "Downtime","location='staticTab.asp'","../../images/icons/bellsystem_g.gif", ""
		Case "AS_HISTORY_DYNAMIC"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Projects","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Predictive Problems","location='staticTab.asp'","../../images/icons/predictive_g.gif", ""
			AddTab "Location Changes","location='staticTab.asp'","../../images/icons/room1_g.gif", ""
			AddTab "Downtime","location='staticTab.asp'","../../images/icons/bellsystem_g.gif", ""
		Case "AS_HISTORY_MCCRM"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Zendesk","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Location Changes","location='staticTab.asp'","../../images/icons/room1_g.gif", ""
			AddTab "Downtime","location='staticTab.asp'","../../images/icons/bellsystem_g.gif", ""
        Case "AS_DATABASE_MCCRM"
            Init "TabControl9", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Databases","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
		Case "WO_HISTORY_DYNAMIC"
			Init "TabControlSH", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Work Order History","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Asset History","location='staticTab.asp'","../../images/icons/building5_g.gif", ""
		Case "AS_DOCSIMAGES_DYNAMIC"
			Init "TabControl2", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Documents","location='staticTab.asp'","../../images/icons/documents_g.gif", ""
			AddTab "Notes","location='staticTab.asp'","../../images/icons/notepadxp_g.gif", ""
			AddTab "Asset / Location Images","location='staticTab.asp'","../../images/icons/scenicxp_g.gif", ""
			AddTab "Work Order Images","location='staticTab.asp'","../../images/icons/photoxp_g.gif", ""
			AddTab "Rules","location='staticTab.asp'","../../images/icons/rules_manager_g.gif", ""
			'If Application("Onsite") = 1 Then
			    'AddTab "Misc Attachments","location='staticTab.asp'","../../images/icons/cabinets_g.gif", ""
		    'End If
		Case "AS_LEASE_TABS"
		    'CB2->Removed history tab until reports are built.
			Init "TabControlAL", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Lease Information","location='staticTab.asp'","../../images/icons/house_g.gif", ""
			AddTab "Contacts","location='staticTab.asp'","../../images/icons/tenant_g.gif", ""
			AddTab "Responsibilities","location='staticTab.asp'","../../images/icons/hammerit_g.gif", ""
			AddTab "Payment Schedule","location='staticTab.asp'","../../images/icons/status_waiting_g.gif", ""
			'If Application("Onsite") = 1 Then
			    AddTab "Documents","location='staticTab.asp'","../../images/icons/documents_g.gif", ""
			'End If
		Case "AS_CONTRACT_TABS"
			Init "TabControlAL", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Contract Information","location='staticTab.asp'","../../images/icons/house_g.gif", ""
			AddTab "Assets","location='staticTab.asp'","../../images/icons/facility_g.gif", ""
			AddTab "Notes","location='staticTab.asp'","../../images/icons/notepadxp_g.gif", ""
            'If Application("Onsite") = 1 Then
			    AddTab "Documents","location='staticTab.asp'","../../images/icons/documents_g.gif", ""
			'End If
		Case "RPT_SETUP_DYNAMIC"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:30px;"
			AddTab "General","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
			'AddTab "Criteria","location='staticTab.asp'","../../images/icons/gearsxp_g.gif", ""
			AddTab "Style / Format","location='staticTab.asp'","../../images/icons/equipxp4_g.gif", ""
			AddTab "Charts / KPIs","location='staticTab.asp'","../../images/icons/chartit_g.gif", ""
			AddTab "Sub-Reports","location='staticTab.asp'","../../images/icons/projectplan2_g.gif", ""
			If GetAccessRight(db,"REESMART",0) or GetSession("IsAdmin") = "Y" Then
			    AddTab "Smart Elements","location='staticTab.asp'","../../images/icons/smartlink_g.gif", ""
            End If
			If GetAccessRight(db,"REESCHEDULE",0) or GetSession("IsAdmin") = "Y" Then
			    AddTab "Schedule","location='staticTab.asp'","../../images/icons/calendarxp_g.gif", ""
    	    End If
			AddTab "Groups","location='staticTab.asp'","../../images/icons/exploreit_g.gif", ""
			AddTab "Security","location='staticTab.asp'","../../images/icons/keysitxp_g.gif", ""
			If GetAccessRight(db,"REEADV",0) or GetSession("IsAdmin") = "Y" Then
    			AddTab "Advanced","location='staticTab.asp'","../../images/icons/gears_g.gif", ""
			End If
		Case "RPT_SETUPCUSTOM_DYNAMIC"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:30px;"
			'AddTab "Smart Elements","location='staticTab.asp'","../../images/icons/smartlink_g.gif", ""
			If GetAccessRight(db,"REESCHEDULE",0) or GetSession("IsAdmin") = "Y" Then
			    AddTab "Schedule","location='staticTab.asp'","../../images/icons/calendarxp_g.gif", ""
            End If
			AddTab "Groups","location='staticTab.asp'","../../images/icons/exploreit_g.gif", ""
			AddTab "Security","location='staticTab.asp'","../../images/icons/keysitxp_g.gif", ""
			If GetAccessRight(db,"REESMART",0) or GetSession("IsAdmin") = "Y" Then
			    AddTab "Smart Elements","location='staticTab.asp'","../../images/icons/smartlink_g.gif", ""
            End If
			If GetAccessRight(db,"REEADV",0) or GetSession("IsAdmin") = "Y" Then
    			AddTab "Advanced","location='staticTab.asp'","../../images/icons/gears_g.gif", ""
			End If
		Case "RPT_SUBREPORTS_DYNAMIC"
			Init "TabControl4", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:0px;"
			AddTab "Report 1","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
			AddTab "Report 2","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
			AddTab "Report 3","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
			AddTab "Report 4","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
			AddTab "Report 5","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
		Case "RPT_STYLEFORMAT_DYNAMIC"
			Init "TabControlSF", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:0px;"
			AddTab "Style","location='staticTab.asp'","../../images/icons/equipxp4_g.gif", ""
			AddTab "Format 1","location='staticTab.asp'","../../images/icons/gearsxp_g.gif", ""
			AddTab "Format 2","location='staticTab.asp'","../../images/icons/gearsxp_g.gif", ""
			AddTab "Format 3","location='staticTab.asp'","../../images/icons/gearsxp_g.gif", ""
		Case "RPT_GENERAL_DYNAMIC"
			Init "TabControlGeneral", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:0px;"
			AddTab "Sort / Group","location='staticTab.asp'","../../images/icons/arrowsdouble_g.gif", ""
			AddTab "Layout","location='staticTab.asp'","../../images/icons/desktopxp_g.gif", ""
			AddTab "Settings","location='staticTab.asp'","../../images/icons/gearpage_g.gif", ""
		Case "RPT_GROUPS_DYNAMIC"
			Init "TabControl1", "0", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Base Groups","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
			AddTab "Custom Groups","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
		Case "RPT_CHART_DYNAMIC"
			Init "TabControl2", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:0px;"
			AddTab "Chart 1","location='staticTab.asp'","../../images/icons/chartit_g.gif", ""
			AddTab "Chart 2","location='staticTab.asp'","../../images/icons/chartit_g.gif", ""
			AddTab "Chart 3","location='staticTab.asp'","../../images/icons/chartit_g.gif", ""
			AddTab "KPIs","location='staticTab.asp'","../../images/icons/equip3_g.gif", ""
		Case "RPT_SMART_DYNAMIC"
			Init "TabControl3", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:0px;"
			AddTab "Smart Settings","location='staticTab.asp'","../../images/icons/smartlink_g.gif", ""
			AddTab "Smart Pane","location='staticTab.asp'","../../images/icons/smartlink2_g.gif", ""
			'AddTab "Smart Tabs","location='staticTab.asp'","../../images/icons/smartlink2_g.gif", ""
			AddTab "Smart Search","location='staticTab.asp'","../../images/icons/smartlink2_g.gif", ""
			AddTab "Smart Actions","location='staticTab.asp'","../../images/icons/smartlink2_g.gif", ""
			AddTab "Smart Buttons","location='staticTab.asp'","../../images/icons/smartbutton_g.gif", ""
			AddTab "Smart Email","location='staticTab.asp'","../../images/icons/smartemail_g.gif", ""
		Case "RPT_SMARTCUSTOM_DYNAMIC"
			Init "TabControl3", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:0px;"
			AddTab "Smart Pane","location='staticTab.asp'","../../images/icons/smartlink2_g.gif", ""
		Case "RPT_SCHEDULE_DYNAMIC"
			Init "TabControl5", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:0px;"
			AddTab "Email","location='staticTab.asp'","../../images/icons/emailitemxp_g.gif", ""
			AddTab "Smart Email","location='staticTab.asp'","../../images/icons/smartemail_g.gif", ""
			AddTab "File","location='staticTab.asp'","../../images/icons/openxp_g.gif", ""
		Case "RPT_SCHEDULECUSTOM_DYNAMIC"
			Init "TabControl5", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:0px;"
			AddTab "Email","location='staticTab.asp'","../../images/icons/emailitemxp_g.gif", ""
			AddTab "File","location='staticTab.asp'","../../images/icons/openxp_g.gif", ""
		Case "RPT_SETUPGROUPS_DYNAMIC"
			Init "TabControl6", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:0px;"
			AddTab "Base Groups","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
			AddTab "Custom Groups","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
		Case "RPT_SECURITY_DYNAMIC"
			Init "TabControl8", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:0px;"
			AddTab "Lock Report Setup","location='staticTab.asp'","../../images/icons/keysitxp_g.gif", ""
			AddTab "Access Groups that can Unlock","location='staticTab.asp'","../../images/icons/keysitxp_g.gif", ""
		Case "RPT_ADVANCED_DYNAMIC"
			Init "TabControl7", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:0px;"
			AddTab "Report Profile","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
			AddTab "SQL Structure","location='staticTab.asp'","../../images/icons/gears_g.gif", ""
			AddTab "Tables used for Available Fields","location='staticTab.asp'","../../images/icons/cylindar2_g.gif", ""
		Case "RPT_ADVANCEDCUSTOM_DYNAMIC"
			Init "TabControl7", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:0px;"
			AddTab "Report Profile","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
			AddTab "SQL Structure","location='staticTab.asp'","../../images/icons/gears_g.gif", ""
			AddTab "Tables used for Available Fields","location='staticTab.asp'","../../images/icons/cylindar2_g.gif", ""
		Case "IN_ORDER_DYNAMIC"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:7px;margin-top:10px;margin-left:10px;width:98%"
			AddTab "Stocked Items","location='staticTab.asp'","../../images/icons/house_g.gif", ""
			AddTab "Direct Issue Items","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
		Case "PM_SCHEDULE_DYNAMIC"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Schedule","location='staticTab.asp'","../../images/icons/calmeter_g.gif", ""
			AddTab "Start / End","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
		Case "SY_ATTACHIMAGE_DYNAMIC"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Upload Image","location='staticTab.asp'","../../images/icons/worldphotoxp_g.gif", ""
			AddTab "Take Snapshot","location='staticTab.asp'","../../images/icons/cameraxp_g.gif", ""
			AddTab "Image Library","location='staticTab.asp'","../../images/icons/scenicxp_g.gif", ""
		Case "PO_SHIPBILL_DYNAMIC"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Shipping Info","location='staticTab.asp'","../../images/icons/fedex_g.gif", ""
			AddTab "Billing Info","location='staticTab.asp'","../../images/icons/actuals_g.gif", ""
		Case "PO_RORI_DYNAMIC"
			Init "TabControl2", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Receipts","location='staticTab.asp'","../../images/icons/cabinets_g.gif", ""
			AddTab "Invoices","location='staticTab.asp'","../../images/icons/calcxp_g.gif", ""
		Case "AS_DETAILS_DYNAMIC"
			Init "TabControl3", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:10px;"
			AddTab "Details","location='staticTab.asp'","../../images/icons/rules_manager_g.gif", ""
			AddTab "Model","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
			AddTab "Costs","location='staticTab.asp'","../../images/icons/calcitxp_g.gif", ""
			AddTab "Insurance","location='staticTab.asp'","../../images/icons/bookxp_g.gif", ""
			AddTab "Manage","location='staticTab.asp'","../../images/icons/hospitalbld_g.gif", ""
			AddTab "Other","location='staticTab.asp'","../../images/icons/task_header_g.gif", ""
		Case "AS_DETAILS_DYNAMIC_MCCRM"
			Init "TabControl3", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:0px;"
			AddTab "Summary","location='staticTab.asp'","../../images/icons/rules_manager_g.gif", ""
			AddTab "Address","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
			AddTab "Billing","location='staticTab.asp'","../../images/icons/calcitxp_g.gif", ""
		'	AddTab "Manage","location='staticTab.asp'","../../images/icons/hospitalbld_g.gif", ""
		'	AddTab "Other","location='staticTab.asp'","../../images/icons/task_header_g.gif", ""
		Case "AS_DETAILS_GIS_DYNAMIC"
			Init "TabControl3", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-top:10px;"
			AddTab "Details","location='staticTab.asp'","../../images/icons/rules_manager_g.gif", ""
			'AddTab "Model","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
			AddTab "Meter","location='staticTab.asp'","../../images/icons/equip3_g.gif", ""
			AddTab "Costs","location='staticTab.asp'","../../images/icons/calcitxp_g.gif", ""
			AddTab "Insurance","location='staticTab.asp'","../../images/icons/bookxp_g.gif", ""
			AddTab "Manage","location='staticTab.asp'","../../images/icons/hospitalbld_g.gif", ""
			AddTab "Other","location='staticTab.asp'","../../images/icons/task_header_g.gif", ""
		Case "AS_RELATED_DYNAMIC"
			Init "TabControl4", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Specifications","location='staticTab.asp'","../../images/icons/object_g.gif", ""
			AddTab "Parts","location='staticTab.asp'","../../images/icons/box3d_g.gif", ""
			AddTab "Labor / Contacts","location='staticTab.asp'","../../images/icons/laborsm_g.gif", ""
			AddTab "Contracts","location='staticTab.asp'","../../images/icons/cabnetfiles_g.gif", ""
			AddTab "Occupants","location='staticTab.asp'","../../images/icons/peoplexp_g.gif", ""
        Case "AS_RELATED_DYNAMIC_MCCRM"
			Init "TabControl4", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Specifications","location='staticTab.asp'","../../images/icons/object_g.gif", ""
			AddTab "Labor / Contacts","location='staticTab.asp'","../../images/icons/laborsm_g.gif", ""
			AddTab "Contracts","location='staticTab.asp'","../../images/icons/cabnetfiles_g.gif", ""
			AddTab "Occupants","location='staticTab.asp'","../../images/icons/peoplexp_g.gif", ""
		Case "AS_PM_DYNAMIC"
			Init "TabControl5", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Preventive Maintenance","location='staticTab.asp'","../../images/icons/status_gened_g.gif", ""
			AddTab "Tracked Tasks","location='staticTab.asp'","../../images/icons/tasks_g.gif", ""

        Case "LA_LOOKUP_FILTERS"
			Init "TabControl1", "0", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:4px;"
			AddTab "Craft","location='staticTab.asp'","../../images/icons/wrenchitxp_g.gif", ""
			AddTab "Category","location='staticTab.asp'","../../images/icons/cabinets_g.gif", ""
			AddTab "Shift","location='staticTab.asp'","../../images/icons/clockxp_g.gif", ""
		'@$CUSTOMISED'
        Case "INV_LOOKUP_FILTERS"
			Init "TabControl1", "0", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:4px;"
			AddTab "Category","location='staticTab.asp'","../../images/icons/cabinets_g.gif", ""
			AddTab "Bin","location='staticTab.asp'","../../images/icons/binxp_g.gif", ""

		Case "WO_COSTS_DYNAMIC"
			Init "TabControl2", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-bottom:10px;"
			AddTab "Estimates","location='staticTab.asp'","../../images/icons/status_wapprove_g.gif", ""
			AddTab "Actuals","location='staticTab.asp'","../../images/icons/actuals_g.gif", ""
			AddTab "All Costs","location='staticTab.asp'","../../images/icons/status_none_g.gif", ""
			AddTab "Purchase Orders","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
		Case "PR_TASKS_DYNAMIC"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-bottom:10px;"
			AddTab "Tasks","location='staticTab.asp'","../../images/icons/tasks_g.gif", ""
			AddTab "Documents","location='staticTab.asp'","../../images/icons/documents_g.gif", ""
			AddTab "Special Instructions","location='staticTab.asp'","../../images/icons/task_header_g.gif", ""
		Case "WO_TASKS_DYNAMIC"
			Init "TabControl3", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "margin-bottom:10px;"
			AddTab "Tasks","location='staticTab.asp'","../../images/icons/tasks_g.gif", ""
			AddTab "Documents","location='staticTab.asp'","../../images/icons/documents_g.gif", ""
			AddTab "Special Instructions","location='staticTab.asp'","../../images/icons/task_header_g.gif", ""
			'AddTab "Notes / Status History","location='staticTab.asp'","../../images/icons/notepadxp_g.gif", ""
			AddTab "Labor Report","location='staticTab.asp'","../../images/icons/laborsm_g.gif", ""
		Case "NEWDOC_DYNAMIC"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Upload","location='staticTab.asp'","../../images/icons/cabinets_g.gif", ""
			If GetAccessRight(db,"DOA",0) and GetAccessRight(db,"DON",0) Then
			    AddTab "Editor","location='staticTab.asp'","../../images/icons/exploreit_g.gif", ""
			    AddTab "Link","location='staticTab.asp'","../../images/icons/online_g.gif", ""
            End If
			'AddTab "Signature","location='staticTab.asp'","../../images/icons/documents_g.gif", ""
		Case "NEWDOC_DYNAMIC_DOCUMENTMODULE"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			'AddTab "Upload","location='staticTab.asp'","../../images/icons/cabinets_g.gif", ""
			AddTab "Editor","location='staticTab.asp'","../../images/icons/exploreit_g.gif", ""
			AddTab "Link","location='staticTab.asp'","../../images/icons/online_g.gif", ""
			'AddTab "Signature","location='staticTab.asp'","../../images/icons/documents_g.gif", ""
		Case "PD_DOCSIMAGES_DYNAMIC"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Images","location='staticTab.asp'","../../images/icons/scenicxp_g.gif", ""
			AddTab "Rules","location='staticTab.asp'","../../images/icons/rules_manager_g.gif", ""
		Case "WO_DOCSIMAGES_DYNAMIC"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Documents","location='staticTab.asp'","../../images/icons/documents_g.gif", ""
			AddTab "Images","location='staticTab.asp'","../../images/icons/scenicxp_g.gif", ""
			AddTab "Rules","location='staticTab.asp'","../../images/icons/rules_manager_g.gif", ""
			'CB2->Added for generic attachments
			'If Application("Onsite") = 1 Then
			    'AddTab "Misc Attachments","location='staticTab.asp'","../../images/icons/cabinets_g.gif", ""
		    'End If
		Case "WO_ASSIGN_DYNAMIC"
			Init "TabControl4", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Assignment List","location='staticTab.asp'","../../images/icons/status_requested_g.gif", ""
			AddTab "Assignment Calendar","location='staticTab.asp'","../../images/icons/calendarxp_g.gif", ""
		Case "WO_REPORT_DYNAMIC"
			Init "TabControl5", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Work Order","location='staticTab.asp'","../../images/icons/status_inprogress_g.gif", ""
			AddTab "Work Order (Statement)","location='staticTab.asp'","../../images/icons/status_approvalrequired_g.gif", ""
		Case "SY_DB_SCHED"
			Init "TabControl1", "0", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Scheduled Jobs","location='staticTab.asp'","../../images/icons/calendarxp_g.gif", ""
			AddTab "Rule-Based Jobs","location='staticTab.asp'","../../images/icons/gearsxp_g.gif", ""
		Case "SY_RULES_MANAGER"
			Init "TabControl1", "0", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Assignment Rules","location='staticTab.asp'","../../images/icons/status_requested_g.gif", ""
			AddTab "Notification Rules","location='staticTab.asp'","../../images/icons/envelopexp2_g.gif", ""
		Case "SY_RULES_MANAGERN"
			Init "TabControl1", "0", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Event Rules","location='staticTab.asp'","../../images/icons/file_light_g.gif", ""
			AddTab "Sensor Rules","location='staticTab.asp'","../../images/icons/equip3_g.gif", ""
			'AddTab "Notification Rules","location='staticTab.asp'","../../images/icons/envelopexp2_g.gif", ""
			AddTab "Report Schedule Rules","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
		Case "SY_RULE_EDIT_EVENT"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Event","location='staticTab.asp'","../../images/icons/file_light_g.gif", ""
			AddTab "Criteria","location='staticTab.asp'","../../images/icons/gearsxp_g.gif", ""
			AddTab "Action","location='staticTab.asp'","../../images/icons/comp_light_g.gif", ""
		Case "SY_RULE_EDIT_SENSOR"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Sensor","location='staticTab.asp'","../../images/icons/equip3_g.gif", ""
			AddTab "Criteria","location='staticTab.asp'","../../images/icons/gearsxp_g.gif", ""
			AddTab "Action","location='staticTab.asp'","../../images/icons/comp_light_g.gif", ""
		Case "SY_RULE_EDIT_AUTO_WOUPDATE"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Routing","location='staticTab.asp'","../../images/icons/arrowsdouble_g.gif", ""
			AddTab "Details","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Status / Authorization","location='staticTab.asp'","../../images/icons/status_closed_g.gif", ""
			AddTab "Indicators","location='staticTab.asp'","../../images/icons/controls_g.gif", ""
			AddTab "Special Instructions","location='staticTab.asp'","../../images/icons/cleanup_g.gif", ""
			AddTab "User-Defined","location='staticTab.asp'","../../images/icons/mnu_edit_g.gif", ""
		Case "SY_RULE_EDIT_AUTO_ASSIGN"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Criteria","location='staticTab.asp'","../../images/icons/gearsxp_g.gif", ""
			AddTab "Labor Assignments","location='staticTab.asp'","../../images/icons/status_requested_g.gif", ""
		Case "SY_RULE_EDIT_AUTO_ASSIGNN"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Assignments","location='staticTab.asp'","../../images/icons/status_requested_g.gif", ""
		Case "SY_RULE_EDIT_AUTO_EMAILN2"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Event / Sensor","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Criteria","location='staticTab.asp'","../../images/icons/gearsxp_g.gif", ""
			AddTab "Email Settings","location='staticTab.asp'","../../images/icons/envelopexp2_g.gif", ""
			AddTab "Event Recipients","location='staticTab.asp'","../../images/icons/supportxp_g.gif", ""
			AddTab "Recipients","location='staticTab.asp'","../../images/icons/usersxp2_g.gif", ""
			AddTab "Message (Text)","location='staticTab.asp'","../../images/icons/report_system_g.gif", ""
			AddTab "Message (HTML)","location='staticTab.asp'","../../images/icons/report_copy_g.gif", ""
			AddTab "Attachments","location='staticTab.asp'","../../images/icons/tasks_g.gif", ""

		Case "SY_RULE_EDIT_AUTO_EMAIL"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Event","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Criteria","location='staticTab.asp'","../../images/icons/gearsxp_g.gif", ""
			AddTab "Actions","location='staticTab.asp'","../../images/icons/comp_light_g.gif", ""

		Case "SY_RULE_EDIT_AUTO_EMAIL_RECORD"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Actions","location='staticTab.asp'","../../images/icons/comp_light_g.gif", ""
			'AddTab "Criteria","location='staticTab.asp'","../../images/icons/gearsxp_g.gif", ""

		Case "SY_RULE_EDIT_AUTO_EMAIL_CLONE"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Event","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			
        Case "SY_RULE_EDIT_EMAIL"
			Init "TabControlEmail", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Settings","location='staticTab.asp'","../../images/icons/envelopexp2_g.gif", ""
			AddTab "Event Recipients","location='staticTab.asp'","../../images/icons/supportxp_g.gif", ""
			AddTab "Recipients","location='staticTab.asp'","../../images/icons/usersxp2_g.gif", ""
			AddTab "Message (Text)","location='staticTab.asp'","../../images/icons/report_system_g.gif", ""
			AddTab "Message (HTML)","location='staticTab.asp'","../../images/icons/report_copy_g.gif", ""
			AddTab "Attachments","location='staticTab.asp'","../../images/icons/tasks_g.gif", ""

        Case "SY_RULE_EDIT_SMS"
			Init "TabControlSMS", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Settings","location='staticTab.asp'","../../images/icons/smsphone_g.gif", ""
			AddTab "Event Recipients","location='staticTab.asp'","../../images/icons/supportxp_g.gif", ""
			AddTab "Recipients","location='staticTab.asp'","../../images/icons/usersxp2_g.gif", ""
			AddTab "Message","location='staticTab.asp'","../../images/icons/report_system_g.gif", ""

        Case "SY_RULE_EDIT_CALL"
			Init "TabControlCALL", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Event Recipients","location='staticTab.asp'","../../images/icons/supportxp_g.gif", ""
			AddTab "Recipients","location='staticTab.asp'","../../images/icons/usersxp2_g.gif", ""
			AddTab "Message","location='staticTab.asp'","../../images/icons/report_system_g.gif", ""

        Case "SY_RULE_EDIT_ALERT"
			Init "TabControlAlert", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Message Center","location='staticTab.asp'","../../images/icons/messagecenter_g.gif", ""
			AddTab "Alert","location='staticTab.asp'","../../images/icons/messagealert_g.gif", ""
			AddTab "Client Command","location='staticTab.asp'","../../images/icons/comp_light_g.gif", ""
			AddTab "Event Recipients","location='staticTab.asp'","../../images/icons/supportxp_g.gif", ""
			AddTab "Recipients","location='staticTab.asp'","../../images/icons/usersxp2_g.gif", ""
			AddTab "Message","location='staticTab.asp'","../../images/icons/report_system_g.gif", ""
			'AddTab "Click Action","location='staticTab.asp'","../../images/icons/comp_light_g.gif", ""

        Case "SY_RULE_EDIT_ALERT1"
			Init "TabControlAlert1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Settings","location='staticTab.asp'","../../images/icons/messagealert_g.gif", ""
			AddTab "Click Action","location='staticTab.asp'","../../images/icons/comp_light_g.gif", ""

        Case "SY_RULE_EDIT_EMAIL_ATTACH"
			Init "TabControlEmailAttach", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Reports","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
			'AddTab "Misc Attachments","location='staticTab.asp'","../../images/icons/cabinets_g.gif", ""

		Case "SY_RULE_EDIT_AUTO_EMAILN"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Email Settings","location='staticTab.asp'","../../images/icons/envelopexp2_g.gif", ""
			AddTab "Event Recipients","location='staticTab.asp'","../../images/icons/supportxp_g.gif", ""
			AddTab "Recipients","location='staticTab.asp'","../../images/icons/usersxp2_g.gif", ""
			AddTab "Message (Text)","location='staticTab.asp'","../../images/icons/report_system_g.gif", ""
			AddTab "Message (HTML)","location='staticTab.asp'","../../images/icons/report_copy_g.gif", ""
			AddTab "Attachments","location='staticTab.asp'","../../images/icons/tasks_g.gif", ""
		Case "SY_APP_CONFIG"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Main Menu","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
			AddTab "Submit Service Request","location='staticTab.asp'","../../images/icons/status_phoned_g.gif", ""
			AddTab "Appearance","location='staticTab.asp'","../../images/icons/paintxp_g.gif", ""
			AddTab "Integration","location='staticTab.asp'","../../images/icons/computerxp_g.gif", ""
			'AddTab "Preferences","location='staticTab.asp'","../../images/icons/exploreit_g.gif", ""
		Case "SY_WO_RAPIDENTRY"
			Init "TabControl1", "0", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:1px;"
			AddTab "Create Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			'AddTab "Complete / Close Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "View Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
            If GetAccessRight(db,"SYT",0) AND GetAccessRight(db,"SYP",0) Then
			    AddTab "Preferences","location='staticTab.asp'","../../images/icons/exploreit_g.gif", ""
            End If
		Case "SY_PM_BALANCER"
			Init "TabControl1", "0", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Projection - Year","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			'AddTab "PMs by Dept","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			'AddTab "PMs by Class","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			'AddTab "PMs by Mfr","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			'AddTab "PMs by Location","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
            'If GetAccessRight(db,"SYT",0) AND GetAccessRight(db,"SYP",0) Then
			'    AddTab "Preferences","location='staticTab.asp'","../../images/icons/exploreit_g.gif", ""
            'End If
		Case "SY_OPTIONS"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			If GetAccessRight(db,"SYA",0) Then
			'AddTab "Actions","location='staticTab.asp'","../../images/icons/gearsxp_g.gif", ""
			End If
			If GetAccessRight(db,"SYT",0) Then
				'AddTab "Tools","location='staticTab.asp'","../../images/kpidefault_small.gif", ""
				If GetAccessRight(db,"SYP",0) Then
					AddTab "Preferences","location='staticTab.asp'","../../images/icons/exploreit_g.gif", ""
				End If
			End If
		Case "WORKOFFLINEDATASETUP"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Setup Offline Criteria","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Current Offline Data","location='staticTab.asp'","../../images/icons/exploreit_g.gif", ""
		Case "WORKOFFLINE"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "General","location='staticTab.asp'","../../images/icons/exploreit_g.gif", ""
			AddTab "Settings","location='staticTab.asp'","../../images/icons/gearpage_g.gif", ""
		Case "TWCMAIN_DYNAMIC"
			Init "TabControl1", "100%", "100%" , "#ffffff", sSelectedTab, False, "TabContents", 80, ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Assets","location='staticTab.asp'","../../images/icons/equipa_g.gif", ""
			AddTab "Parts","location='staticTab.asp'","../../images/icons/box3d_g.gif", ""
			AddTab "Tools","location='staticTab.asp'","../../images/icons/item_g.gif", ""
			AddTab "Contacts","location='staticTab.asp'","../../images/icons/peoplexp_g.gif", ""
			AddTab "Procedures","location='staticTab.asp'","../../images/icons/status_inprogress_g.gif", ""
			AddTab "Predictive","location='staticTab.asp'","../../images/icons/predictive_g.gif", ""
			AddTab "Documents","location='staticTab.asp'","../../images/icons/documents_g.gif", ""
			AddTab "Profile","location='staticTab.asp'","../../images/icons/usersxp2_g.gif", ""
			AddTab "Reports","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
		Case "SmartCriteria"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Filter","location='staticTab.asp'","../../images/icons/funnel_g.gif", ""
		Case "KPI_PORTAL"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Filter","location='staticTab.asp'","../../images/icons/funnel_g.gif", ""
			AddTab "Compare","location='staticTab.asp'","../../images/icons/file_light_g.gif", ""
		Case "KPI_CHOICE"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "","location='staticTab.asp'","../../images/icons/equip3_g.gif", ""
			AddTab "","location='staticTab.asp'","../../images/icons/calmeter_g.gif", ""
			AddTab "","location='staticTab.asp'","../../images/icons/exploreit_g.gif", ""
			AddTab "","location='staticTab.asp'","../../images/icons/piechart_g.gif", ""
		Case "KPI_CHOICE2"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "KPIs","location='staticTab.asp'","../../images/icons/equip3_g.gif", ""
			AddTab "Reports","location='staticTab.asp'","../../images/icons/exploreit_g.gif", ""
			AddTab "Charts","location='staticTab.asp'","../../images/icons/piechart_g.gif", ""
		Case "SY_FIELDS_MANAGER"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Priority Actions","location='staticTab.asp'","../../images/icons/file_light_g.gif", ""
		Case "SY_BACKRESTORE_MANAGER"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Set Backup Schedule","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
			AddTab "Backup Manually","location='staticTab.asp'","../../images/icons/status_phoned_g.gif", ""
			'AddTab "Restore Manaully","location='staticTab.asp'","../../images/icons/paintxp_g.gif", ""
		Case "AS_PASTE_SPECIAL"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "General","location='staticTab.asp'","../../images/icons/gearsxp_g.gif", ""
			AddTab "Search/Replace","location='staticTab.asp'","../../images/icons/file_light_g.gif", ""
			AddTab "Details","location='staticTab.asp'","../../images/icons/rules_manager_g.gif", ""
			AddTab "Repair Center","location='staticTab.asp'","../../images/icons/building4_g.gif", ""
		Case "SY_CHECKINCHECKOUT_MANAGER"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Checkout","location='staticTab.asp'","../../images/icons/status_phoned_g.gif", ""
			AddTab "Checkin","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
		Case "SPEC_TYPE"
			Init "TabControl2", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Asset","location='staticTab.asp'","../../images/icons/building5_g.gif", ""
			AddTab "Item","location='staticTab.asp'","../../images/icons/box3d_g.gif", ""
		Case "SPEC_TYPE_GIS"
			Init "TabControl2", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "GIS Attributes","location='staticTab.asp'","../../images/icons/scenicxp_g.gif", ""
		Case "SPEC_VALUE_TYPE"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Text","location='staticTab.asp'","../../images/icons/dbfield_g.gif", ""
			AddTab "Date","location='staticTab.asp'","../../images/icons/dbfield_g.gif", ""
			AddTab "Numeric","location='staticTab.asp'","../../images/icons/dbfield_g.gif", ""
		Case "SY_CHECKINCHECKOUT_MANAGER"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Checkout","location='staticTab.asp'","../../images/icons/status_phoned_g.gif", ""
			AddTab "Checkin","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
		Case "SY_CHECKOUTTO_CHOICE"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Person","location='staticTab.asp'","../../images/icons/status_phoned_g.gif", ""
			AddTab "Work Order","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
            AddTab "Account","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
		Case "SY_CHECKOUTWHAT_CHOICE"
			Init "TabControl2", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Items","location='staticTab.asp'","../../images/icons/status_phoned_g.gif", ""
			AddTab "Tools","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
            AddTab "Assets","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
		Case "SY_TRANSFERWHAT_CHOICE"
			Init "TabControl2", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Items","location='staticTab.asp'","../../images/icons/status_phoned_g.gif", ""
		Case "SY_CHECKINWHAT_CHOICE"
			Init "TabControl2", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "All","location='staticTab.asp'","../../images/icons/status_phoned_g.gif", ""
			AddTab "Items","location='staticTab.asp'","../../images/icons/status_phoned_g.gif", ""
			AddTab "Tools","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
            AddTab "Assets","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
		Case "SY_KEYTRACKING_MANAGER"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:3px;"
			AddTab "Keys","location='staticTab.asp'","../../images/icons/keysitxp_g.gif", ""
			AddTab "Key Holders","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
            AddTab "Key Transactions","location='staticTab.asp'","../../images/icons/dbfield_g.gif", ""
            AddTab "Report","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
        Case "PMsTab"
			Init "TabControlPM", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Assets","location='staticTab.asp'","../../images/icons/building5_g.gif", ""
			AddTab "PM Schedules","location='staticTab.asp'","../../images/icons/status_gened_g.gif", ""
			AddTab "PM Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "PM Backlog","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
        Case "DPReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Assets","location='staticTab.asp'","../../images/icons/building5_g.gif", ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Backlog","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Cost Breakdown","location='staticTab.asp'","../../images/icons/money_g.gif", ""
        Case "RCReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Assets","location='staticTab.asp'","../../images/icons/building5_g.gif", ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Backlog","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Cost Breakdown","location='staticTab.asp'","../../images/icons/money_g.gif", ""
        'CB2->Approvals Sub Tabs
        Case "RCApprovalsTab"
			Init "TabControlRCA", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Purchase Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			'AddTab "Work Orders","location='staticTab.asp'","../../images/icons/building5_g.gif", ""
        Case "SHReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Assets","location='staticTab.asp'","../../images/icons/building5_g.gif", ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Backlog","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Cost Breakdown","location='staticTab.asp'","../../images/icons/money_g.gif", ""
        Case "TNReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Assets","location='staticTab.asp'","../../images/icons/building5_g.gif", ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Backlog","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Cost Breakdown","location='staticTab.asp'","../../images/icons/money_g.gif", ""
        Case "ACReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Assets","location='staticTab.asp'","../../images/icons/building5_g.gif", ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Backlog","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Cost Breakdown","location='staticTab.asp'","../../images/icons/money_g.gif", ""
        Case "ZNReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Assets","location='staticTab.asp'","../../images/icons/building5_g.gif", ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Backlog","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Cost Breakdown","location='staticTab.asp'","../../images/icons/money_g.gif", ""
        Case "SFReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Backlog","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Cost Breakdown","location='staticTab.asp'","../../images/icons/money_g.gif", ""
        Case "INReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "All","location='staticTab.asp'","../../images/icons/box3d_g.gif", ""
			AddTab "Issues","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Receipts","location='staticTab.asp'","../../images/icons/money_g.gif", ""
			AddTab "Transfers","location='staticTab.asp'","../../images/icons/fedex_g.gif", ""
			AddTab "Adjustments","location='staticTab.asp'","../../images/icons/calcitxp_g.gif", ""
			AddTab "Rotations","location='staticTab.asp'","../../images/icons/refresh_g.gif", ""
			AddTab "Reserved","location='staticTab.asp'","../../images/icons/status_generated_g.gif", ""
			AddTab "Direct Issues","location='staticTab.asp'","../../images/icons/status_inprogress_g.gif", ""
			AddTab "On-Order","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
        Case "INWhereIssuedInstalledTab"
			Init "TabControlWU", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Stock Rooms","location='staticTab.asp'","../../images/icons/house_g.gif", ""
			AddTab "Rotating Assets","location='staticTab.asp'","../../images/icons/building5_g.gif", ""
        Case "FAReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Backlog","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Cost Breakdown","location='staticTab.asp'","../../images/icons/money_g.gif", ""
        Case "CAReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Backlog","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Cost Breakdown","location='staticTab.asp'","../../images/icons/money_g.gif", ""
        Case "LAReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Backlog","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Cost Breakdown","location='staticTab.asp'","../../images/icons/money_g.gif", ""
        Case "RQReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Backlog","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Cost Breakdown","location='staticTab.asp'","../../images/icons/money_g.gif", ""
        Case "CLReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Assets","location='staticTab.asp'","../../images/icons/building5_g.gif", ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Backlog","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Cost Breakdown","location='staticTab.asp'","../../images/icons/money_g.gif", ""
        Case "PRReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Backlog","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Cost Breakdown","location='staticTab.asp'","../../images/icons/money_g.gif", ""
        Case "PMReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Work Orders","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Backlog","location='staticTab.asp'","../../images/icons/equipxp_g.gif", ""
			AddTab "Cost Breakdown","location='staticTab.asp'","../../images/icons/money_g.gif", ""
        Case "TRReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Labor","location='staticTab.asp'","../../images/icons/laborsm_g.gif", ""
        Case "SPReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Assets","location='staticTab.asp'","../../images/icons/building5_g.gif", ""
			AddTab "Items","location='staticTab.asp'","../../images/icons/box3d_g.gif", ""
        Case "AGReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Members","location='staticTab.asp'","../../images/icons/building5_g.gif", ""
        Case "CRReportsTab"
			Init "TabControlRPT", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Labor","location='staticTab.asp'","../../images/icons/laborsm_g.gif", ""
		Case "SY_RESOURCESCHEDULER_MANAGER"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Schedule","location='staticTab.asp'","../../images/icons/status_phoned_g.gif", ""
			AddTab "Find","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
		Case "SY_CALENDAR_DWM"
			Init "TabControl2", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "Day","location='staticTab.asp'","../../images/icons/status_phoned_g.gif", ""
			AddTab "Week","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
			AddTab "Month","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
		Case "SY_CALENDAR_GAE"
			Init "TabControl3", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:2px;"
			AddTab "General","location='staticTab.asp'","../../images/icons/status_phoned_g.gif", ""
			AddTab "Attendees","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
			AddTab "Equipment","location='staticTab.asp'","../../images/icons/user_requester_g.gif", ""
        Case "ASGIS"
			Init "TabControlGIS", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Map Viewer","location='staticTab.asp'","../../images/icons/worldxp3_g.gif", ""
        Case "SY_FORMSMGR"
			Init "TabControl1", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:4px;"
			AddTab "Forms","location='staticTab.asp'","../../images/icons/paperxp_g.gif", ""
			AddTab "Fields","location='staticTab.asp'","../../images/icons/mnu_edit_g.gif", ""
		Case "SYMEMBERS"
			Init "TabMembers", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, "position:relative; top:4px;"
			AddTab "Members","location='staticTab.asp'","../../images/icons/member_attach_g.gif", ""
			If GetAccessRight(db,"SYPasswordSettings",0) Then
                AddTab "Security Settings","location='staticTab.asp'","../../images/icons/gearpage_g.gif", ""
			End If
                   'CB2//Commented out while we come up with a good way to do member activity.
			'AddTab "Recent Activity","location='staticTab.asp'","../../images/icons/glasses2_g.gif", ""
        		
        ' New Tabs custom Risk page WO85594 12/2016      
		Case "AS_RISK_TAB"
			Init "TabControlAEM", "100%", "0" , "transparent", sSelectedTab, False, "TabContentsNoBorder", 15, ""
			AddTab "Details","location='staticTab.asp'","../../images/icons/gearsxp_g.gif", ""
			AddTab "AEM Program","location='staticTab.asp'","../../images/icons/calendarxp_g.gif", ""
        		' End of changes....
		End Select
		'_SBS_
		Draw
	End Sub

	Public Sub AddTab (sName, sAction, sIcon, sDisabled)
		dim item
		item = Ubound(TabList,2) + 1
		Redim Preserve TabList(4,item)
		TabList(0,item) = SName
		TabList(1,item) = sAction
		TabList(2,item) = sIcon
		TabList(3,item) = sDisabled

		'If no tab has been initialized as selected make the first tab active
		If sSelected = "" Then sSelected = SName

	End Sub

	Public Sub Draw()
		If bStatic then
			DrawStatic
		Else
			DrawDynamic
		End If
	End Sub

	Private Sub DrawDynamic()

		dim i, sImage,bLastOn,iTabOn
		dim drawclass
		dim drawclassStart
		dim drawWidth

		%>
		<%if debug = 0 then %>
		<DIV STYLE="height:22px; border:none; <% =MainDivStyle %>">
		<TABLE bgcolor="<%=BGColor%>" width="<%=TabWidth%>" CELLPADDING="0" CELLSPACING="0" BORDER="0">
			<TR onselectStart="return false;">
			<%
				For i = 1 to ubound(TabList,2)

					If TabList(0,i) = sSelected and TabList(3,i) = "" Then
						drawclass="SelectedTab"
						bLastOn = True
						iTabOn = i
						If i=1 Then
							drawWidth = "8"
							sImage = "start.on"
						Else
							sImage = "off.on"
						End If
					Else
						drawclass="Tab" & TabList(3,i)
						If i=1 Then
							drawWidth = "8"
							sImage = "start.off"
						Else
							drawWidth = "17"
							If bLastOn Then
								bLastOn = False
								sImage = "on.off"
							Else
								sImage = "off.off"
							End If
						End If
					End If
						%>
							<TD nowrap><img id="<% =ControlName %>_TabSpacer<%=i%>" border="0" src="../../images/tabs/menu.<%=sImage%>.gif"></TD>
							<TD ID="<% =ControlName %>_t<%=i%>" CLASS="<%=drawclass%>" HEIGHT=21 onmouseover="TabOver('<% =ControlName %>',this)" onmouseout="TabOver('<% =ControlName %>',this)" onclick="changeTabs('<% =ControlName %>',this);" nowrap><% If Not TabList(2,i) = "" Then %><img style="display:none;" class="TabIcon<% =TabList(3,i) %>" src="<%=TabList(2,i) %>" border="0"><% End If %><%=TabList(0, i)%></TD>
						<%

				next
			%>
				<script type="text/javascript">
					var <% =ControlName %>_firstFlag = false;
					var <% =ControlName %>_currentTab = self.document.getElementById('<% =ControlName %>_t<%=iTabOn%>');

					if (top.hideSubTabbarIcons == false)
					{
					    if (top.fraTopic.refreshcurrentsubtabs)
					    {
					        top.fraTopic.refreshcurrentsubtabs();
					    }
					    if (self.refreshcurrentsubtabs)
					    {
					        self.refreshcurrentsubtabs();
					    }
                    }
				</script>
				<% If iTabOn = ubound(TabList,2) Then %>
				<TD style="background-color:transparent;" background="../../images/tabs/menu.nothing.bg.gif" nowrap><img id="<% =ControlName %>_TabSpacer<%=i%>" border="0" src="../../images/tabs/menu.on.end.gif"></TD>
				<% Else %>
				<TD style="background-color:transparent;" background="../../images/tabs/menu.nothing.bg.gif" nowrap><img id="<% =ControlName %>_TabSpacer<%=i%>" border="0" src="../../images/tabs/menu.off.end.gif"></TD>
				<% End If %>
				<TD style="background-color:transparent;" background="../../images/tabs/menu.nothing.bg.gif" width="100%" nowrap>&nbsp;</td>
			</TR>
			</table>
			<% ' COLSPAN="%=(ubound(Tablist,2)*2)+2%"  %>
			<TABLE id="<%=ControlName%>_table" bgcolor="<%=BGColor%>" width="<%=TabWidth%>" height="<%=TabHeight%>" CELLPADDING="0" CELLSPACING="0" BORDER="0">
			<TR>
				<TD HEIGHT="100%" valign="top" ID="<% =ControlName %>_tabContents" class="<%=TCClass%>">
<%
		end if
	End Sub

	Private Sub DrawStatic()
    Response.End
		dim i
		dim drawclass
		dim drawclassStart
		dim drawWidth
		dim bLastOn
		dim sImage%>
		<%if debug = 0 then %>
		<TABLE id="<%=ControlName%>_table" bgcolor='<%=BGColor%>' width=<%=TabWidth%> height=<%=TabHeight%> CELLPADDING=0 CELLSPACING=0 BORDER=0>
			<TR onselectStart="return false;">

			<%
				For i = 1 to ubound(TabList,2)

					If TabList(0,i) = sSelected Then
						drawclass="SelectedTab"
						bLastOn = True

						If i=1 Then
							drawWidth = "8"
							sImage = "start.on"
						Else
							sImage = "off.on"
						End If
					Else
						drawclass="Tab"
						If i=1 Then
							drawWidth = "8"
							sImage = "start.off"
						Else
							drawWidth = "17"
							If bLastOn Then
								bLastOn = False
								sImage = "on.off"
							Else
								sImage = "off.off"
							End If
						End If
					End If
						%>
							<TD nowrap><img id="<% =ControlName %>_TabSpacer<%=i%>" border="0" src="../../images/tabs/menu.<%=sImage%>.gif"></TD>
							<TD ID="<% =ControlName %>_t<%=i%>" CLASS="<%=drawclass%>" HEIGHT=21 onmouseover="TabOver('<% =ControlName %>',this)" onmouseout="TabOver('<% =ControlName %>',this)" onclick="<%=TabList(1, i)%>" nowrap><% If Not TabList(2,i) = "" Then %><img style="display:none;" class="TabIcon" src="<%=TabList(2,i) %>" border="0"><% End If %><%=TabList(0, i)%></TD>
						<%
				next

			If bLastOn Then sImage = "on.end" Else sImage = "off.end"

			%>
				<TD style="background-color:transparent;" background="../../images/tabs/menu.nothing.bg.gif" nowrap><img id="<% =ControlName %>_TabSpacer<%=i%>" border="0" src="../../images/tabs/menu.<%=sImage%>.gif"></TD>
				<TD style="background-color:transparent;" background="../../images/tabs/menu.nothing.bg.gif" width=100% nowrap>&nbsp;</td>
			</TR>
			<TR>
				<TD HEIGHT="100%" valign="top" align="left" COLSPAN="<%=(ubound(Tablist,2)*2)+2%>" ID="<% =ControlName %>_tabContents" class="<% =TCClass %>">&nbsp;
<%
		end if
	End Sub

	Public Sub EndAdd

		if debug then
			Response.Write "<BR><BR>END<BR><BR>"
		else
			Response.Write "</DIV>"
		end if
	End Sub

	Public Sub EndTabs

		If bStatic Then
			Response.Write "</TD></TR></TABLE>"
			If TabHeight = "100%" Then
			%>
			<script language="javascript">
			    self.document.getElementById('<%=ControlName%>_table').height = self.innerHeight-<% =HeightSubtract %>+'px';
			</script>
			<%
			End If
		Else
			%>
					</TD>
				</TR>
			</TABLE>
			</DIV>
			<script language="javascript">			    
			    if (self.<% =ControlName %>_currentTab)
			    {
			        var <% =ControlName %>_currentTabContent = document.getElementById('<% =ControlName %>_TAB_' + <% =ControlName %>_currentTab.innerText);
			    }
			    else
			    {
			        var <% =ControlName %>_currentTabContent = null;
			    }
			    if (self.<% =ControlName %>_currentTabContent)
			    {
			        self.<% =ControlName %>_currentTabContent.style.display = '';
			    }
			    <% If TabHeight = "100%" Then %>
					try {
					self.document.getElementById('<%=ControlName%>_table').height = self.innerHeight-<% =HeightSubtract %> + 'px';
			    } catch(e) {}
			    <% End If %>
			</script>
			<%
		End If

	End Sub

	Public Sub EndTabsOnly

		If bStatic Then
			Response.Write "</TD></TR></TABLE>"
		Else
			%>
					</TD>
				</TR>
			</TABLE>
			</DIV>
			<%
		End If

	End Sub

	Public Sub EndTabsJSOnly

		If bStatic Then
			If TabHeight = "100%" Then
			%>
			<script language="javascript">
				self.document.getElementById('<%=ControlName%>_table').height = self.innerHeight-<% =HeightSubtract %>+'px';
			</script>
			<%
			End If
		Else
			%>
			<script language="javascript">
			    if (self.<% =ControlName %>_currentTab)
			    {
			        var <% =ControlName %>_currentTabContent = document.getElementById('<% =ControlName %>_TAB_' + <% =ControlName %>_currentTab.innerText);
			    }
			    else
			    {
			        var <% =ControlName %>_currentTabContent = null;
			    }
			    if (self.<% =ControlName %>_currentTabContent)
			    {
			        self.<% =ControlName %>_currentTabContent.style.display = '';
			    }
			    <% If TabHeight = "100%" Then %>
					try {
					self.document.getElementById('<%=ControlName%>_table').height = self.innerHeight - <% =HeightSubtract %> + 'px';
			    } catch(e) {}
			    <% End If %>
			</script>
			<%
		End If

	End Sub

End Class
%>