SET XACT_ABORT OFF
SET ARITHABORT OFF

/* DECLARE GLOBAL VARIABLES */
DECLARE @rptPK int
DECLARE @rptGroupPK int
DECLARE @tempPK int
DECLARE @errorcount int
DECLARE @dsql varchar(8000)
DECLARE @version decimal(9,2)

SELECT @version = CAST(REPLACE(LOWER(schema_version), 'sp', '') AS decimal(9,2)) FROM _schema

/* =====================================================
UPDATE REPORTS STYLES AND GROUP HEADERS - (Sandvik_FtpErrors)
===================================================== */

IF @version >= 3.0 
BEGIN 
EXEC('PRINT ''Version 3.0 or greater detected. Importing Report Styles and Group Headers''
DECLARE @tempPK int
IF NOT EXISTS(SELECT ReportStyleName FROM ReportStyle WHERE ReportStyleName = ''Gradient - Blue'')
BEGIN 
INSERT INTO ReportStyle (ReportStyleName, ReportStyleDesc, ReportStyleCSS, IsDefault, IsBase, RowVersionIPAddress, RowVersionUserPK, RowVersionInitials, RowVersionAction, RowVersionDate) VALUES(''Gradient - Blue'', null, '' .pageselect{FONT-SIZE: 9pt; COLOR: #333333; FONT-FAMILY: Arial}
 .heading {background-color:#ffffff; cursor:pointer; FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: royalblue; FONT-FAMILY: Arial; z-index: 2500;} 
 .legendHeader {FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: #333333; FONT-FAMILY: Arial}
 .normaltext {FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial}
 .labels {FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: royalblue; FONT-FAMILY: Arial}
 .assetUP {FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: green; FONT-FAMILY: Arial}
 .assetDOWN {FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #DD0000; FONT-FAMILY: Arial}
 .asset {FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Arial}
 .data {FONT-SIZE: 12px; COLOR: #494949; FONT-FAMILY: Arial}
 .data_underline {BORDER-RIGHT: medium none; BORDER-TOP: medium none; FONT-SIZE: 12px; BORDER-LEFT: medium none; COLOR: #494949; BORDER-BOTTOM: #333333 1px solid; FONT-FAMILY: Arial}
 .bottomline {BORDER-RIGHT: medium none; BORDER-TOP: medium none; BORDER-LEFT: medium none; BORDER-BOTTOM: #333333 1px solid}
 .buttons {FONT-SIZE: 12px; WIDTH: 80px; cursor: pointer; COLOR: #333333; FONT-FAMILY: Arial}
 .subtotal {BORDER-RIGHT: medium none; BORDER-TOP: #C0C0C0 1px solid; FONT-SIZE: 12px; BORDER-LEFT: medium none; COLOR: #333333; BORDER-BOTTOM: medium none; FONT-FAMILY: Arial}
 .bodyclasspreview {background-color:#ffffff; padding:10px; scrollbar-base-color: #FBFBFB; font-size:8pt; font-family:Arial; color:#000000;}
 .bodyclasspreviewinwo {background-color:#ffffff; padding-right:10px; scrollbar-base-color: #FBFBFB; font-size:8pt; font-family:Arial; color:#000000;}
 .bodyclassprint {background-color:#ffffff; PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; PADDING-TOP: 0px; font-size:8pt; font-family:Arial; color:#000000;}
 .bodyclassemail {background-color:#ffffff; PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 20px; PADDING-TOP: 0px; font-size:8pt; font-family:Arial; color:#000000;}
 .group1 {padding-left:0px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; FONT-WEIGHT: Bold; BACKGROUND-COLOR: #acc5e7;}
 .group2 {padding-left:10px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; FONT-WEIGHT: Bold; BACKGROUND-COLOR: #c7d7ed;}
 .group3 {padding-left:20px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; FONT-WEIGHT: Bold; BACKGROUND-COLOR: #dce8f4;}
 .group4 {padding-left:30px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; FONT-WEIGHT: Bold; BACKGROUND-COLOR: #ecf1fb;}
 .group5 {padding-left:40px; FONT-SIZE: 12px; COLOR: royalblue; FONT-FAMILY: Arial; FONT-WEIGHT: Bold; BACKGROUND-COLOR: #ffffff;}
 .groupheader {FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; FONT-WEIGHT: Bold;}
 .normalright {BORDER-RIGHT: #c0c0c0 1px solid; BORDER-TOP: #c0c0c0 1px solid; PADDING-LEFT: 1px; FONT-WEIGHT: normal; FONT-SIZE: 8pt; MARGIN-BOTTOM: 1px; BORDER-LEFT: #c0c0c0 1px solid; COLOR: #000000; BORDER-BOTTOM: #c0c0c0 1px solid; FONT-FAMILY: Arial; BACKGROUND-COLOR: #ffffff; TEXT-ALIGN: right}
 .clsBtnUp {cursor: pointer; color: black; font-weight: normal; border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #B4B4B4;border-bottom:1px solid #B4B4B4;padding-right:2px;}
 .clsBtnDown {cursor: pointer; color: black; font-weight: normal; border-right:1px solid #ffffff;border-bottom:1px solid #ffffff;border-top:1px solid #B4B4B4;border-left:1px solid #B4B4B4;padding-right:2px;}
 .clsBtnOff {color: black; font-weight: normal; tab-index: 0; border:1px solid transparent; padding-right:2px;}
 .actionbarlabel {float:right;margin-top:6px;padding-left:5px;padding-right:5px;font-size:8pt;font-family:Arial;color:#000000;}
 INPUT {padding-left:3px;}
 A:link {FONT-SIZE: 8pt; cursor: pointer; COLOR: #315aad; FONT-FAMILY: Arial; BACKGROUND-COLOR: transparent;}
 A:visited {FONT-SIZE: 8pt; cursor: pointer; COLOR: #315aad; FONT-FAMILY: Arial; BACKGROUND-COLOR: transparent;}
 A:active {FONT-SIZE: 8pt; cursor: pointer; COLOR: #315aad; FONT-FAMILY: Arial; BACKGROUND-COLOR: transparent;}
 A:hover {COLOR: red;}
 fieldset {border: 1px solid #AAAAAB;}
 .buttonsdisabled { display: static; opacity: 0.4; cursor:pointer;	}
 .buttonsenabled { display:	; opacity: 1; cursor:pointer; }
 .normalrow {FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial }
 .tb {width:100%; PADDING-LEFT: 1px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; border: 1px solid #AAAAAB;}
 .tbf {BACKGROUND-COLOR: #ffffcc; width:100%; PADDING-LEFT: 1px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; border: 1px solid #AAAAAB;}
 .ta {width:200px; PADDING-LEFT: 1px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; border: 1px solid #AAAAAB;}
 .taf {BACKGROUND-COLOR: #ffffcc; width:200px; PADDING-LEFT: 1px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; border: 1px solid #AAAAAB;}
 .cb {COLOR: #333333;}
 .HeaderRight {font-family:Arial;font-size:16px;color:#333333;font-weight:bold}
 .SubHeaderRight {font-family:Arial;font-size:11px;font-weight:normal}
 .SRInstructions {margin-top:5px;font-family:Arial;font-size:8pt;color:green;font-weight:bold}
 .verticalcolumn {border:1px solid #CCCCCC;}
 .mcpagebreak {page-break-before: always;}
 .ReportRow1 {background-color:#FFFFFF; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial }
 .ReportRow2 {background-color:#EFEFEF; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial }
 .ReportRowCrit1 {background-color:#FFFFFF; FONT-SIZE: 8pt; COLOR: #333333; FONT-FAMILY: Arial }
 .ReportRowCrit2 {background-color:#EFEFEF; FONT-SIZE: 8pt; COLOR: #333333; FONT-FAMILY: Arial }
 .SmartRow {background-color:#FFDF84; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; cursor:pointer; }
 .SubReportRow {background-color:#DEEFC6; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial }
 .ExpandCollapse {font-family:Arial;font-size:8pt;color:#333333;cursor:pointer;}'', 0, 0, '''', '''', ''_MC'', '''', getdate())
END 
')
END 
/* ============================================================
ReportID: Sandvik_FtpErrors
Report Name: FTP Import Errors
============================================================ */

IF EXISTS (
SELECT ReportName FROM Reports WITH (NOLOCK)
WHERE ReportID = 'Sandvik_FtpErrors')

BEGIN

PRINT '*******************************************************'
PRINT 'Report Exists - Updating...'
PRINT 'Report: FTP Import Errors'
PRINT '*******************************************************'

/* ================================================
UPDATE REPORT RECORDS - (Sandvik_FtpErrors)
================================================ */
/* Set ReportPK for this Report */
SELECT @rptPK = ReportPK FROM Reports WITH (NOLOCK) WHERE ReportID='Sandvik_FtpErrors'

/* Update Main Report Fields */
UPDATE Reports
SET	[ReportIDPriorToCopy]='WOList', [ReportDesc]=null, [Sort1]='MC_InterfaceLog.ProcessDate', [Sort2]=null, [Sort3]=null, [Sort4]=null, [Sort5]=null, [Sort1DESC]=1, [Sort2DESC]=0, [Sort3DESC]=0, [Sort4DESC]=0, [Sort5DESC]=0, [Group1]=0, [Group2]=0, [Group3]=0, [Group4]=0, [Group5]=0, [Header1]=0, [Header2]=0, [Header3]=0, [Header4]=0, [Header5]=0, [GroupHeader1]=null, [GroupHeader2]=null, [GroupHeader3]=null, [GroupHeader4]=null, [GroupHeader5]=null, [Total1]=0, [Total2]=0, [Total3]=0, [Total4]=0, [Total5]=0, [Chart]=null, [ChartName]='Work Order Status', [ChartField]='AUTO_SORT1', [ChartSize]='L', [ReportFile]='rpt_generic1.asp', [FromSQL]='FROM MC_InterfaceLog', [JoinSQL]=null, [WhereSQL]=null, [GroupBy]=0, [hits]=5, [Sequence]=0, [Layout]='hor', [VertCols]=1, [PageBreakEachRecord]=0, [Custom]=0, [ReportCopy]=1, [MCRegistrationDB]=0, [PrintCriteria]=0, [Active]=1, [UDFChar1]=null, [UDFChar2]=null, [UDFChar3]=null, [UDFChar4]='N', [UDFChar5]=null, [UDFDate1]=null, [UDFDate2]=null, [DemoLaborPK]=null, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionAction]='EDIT', [RowVersionDate]=getdate() , [ChartFunction]='C', [ChartFunctionField]='NONE', [NoDetail]=0, [PB1]=0, [PB2]=0, [PB3]=0, [PB4]=0, [PB5]=0, [SLDefault]=0, [SLType]=' ', [SLAction]='PW', [SLModuleID]='WO', [SLPKField]=null, [SLReportID]=null, [SLCustomAction]=null, [SLTooltip]=null, [SDDisplay]='     ', [SDModuleID]='  ', [SDPKField]=null, [SmartEmail]=0, [ChartPosition]='T', [ChartFormat]='F', [ChartSQL]=null, [Chart2]=null, [ChartName2]=null, [ChartField2]=null, [ChartSize2]=null, [ChartFormat2]='I', [ChartFunction2]=null, [ChartFunctionField2]=null, [ChartPosition2]=null, [ChartSQL2]=null, [Chart3]=null, [ChartName3]=null, [ChartField3]=null, [ChartSize3]=null, [ChartFormat3]='I', [ChartFunction3]=null, [ChartFunctionField3]=null, [ChartPosition3]=null, [ChartSQL3]=null, [ChartOnly]=0, [NoHeader]=0, [SRID1]=null, [SRPKField1]=null, [SRID2]=null, [SRPKField2]=null, [SRID3]=null, [SRPKField3]=null, [SRID4]=null, [SRPKField4]=null, [SRID5]=null, [SRPKField5]=null, [ReportPageSize]='Default', [ReportWidth]='80%', [PhotoCriteria]=0, [ReportStyleName]='Gradient - Blue', [UsedFor]='REPORTS', [SmartEmailLaborPK]=0, [SCDefault]='H', [SCField1]=null, [SCField2]=null, [SCField3]=null, [ReportStyleFontSize]=null, [ReportStyleFontColor]=null, [ReportStyleFontFamily]=null 
WHERE ReportPK = @rptPK
IF (@@error > 0) SET @errorcount = @errorcount + 1

IF @version >= 3.0 
BEGIN
SET @dSQL = 'UPDATE Reports SET [HavingSQL]=null, [DisplayPivotBar]=0, [DisplayColumnLines]=0, [DisplayTitleonPageBreak]=0, [DisplayFormatCriteria]=1, [R1T]='' '', [R1O]=null, [R1V1]=null, [R1V2]=null, [R1A]=0, [R1L]=0, [R1F]=0, [R1CS]=''font-family: Arial;font-size: 8pt;color: #000000;text-align: left;'', [R1AF]=''C'', [R2T]='' '', [R2O]=null, [R2V1]=null, [R2V2]=null, [R2A]=0, [R2L]=0, [R2F]=0, [R2CS]=''border: #0066CC 2px solid;'', [R2AF]=''C'', [R3T]='' '', [R3O]=null, [R3V1]=null, [R3V2]=null, [R3A]=0, [R3L]=0, [R3F]=0, [R3CS]=''border: #0066CC 2px solid;'', [R3AF]=''C'' WHERE ReportPK = ' + CAST(@rptPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1

END

IF @version >= 4.2 
BEGIN
SET @dSQL = 'UPDATE Reports SET [DisplayDescription]=0 WHERE ReportPK = ' + CAST(@rptPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1

END

PRINT 'Updating Report - Sandvik_FtpErrors'

/* ==================================================
DELETE AND INSERT REPORT GROUPS - (Sandvik_FtpErrors)
=================================================== */
DELETE FROM Report_ReportGroup WHERE ReportPK = @rptPK

PRINT 'Deleting Report_ReportGroup Rows - Sandvik_FtpErrors'

IF (@@error > 0) SET @errorcount = @errorcount + 1

/* Make sure the report group actually exists */
IF NOT EXISTS (SELECT ReportGroupPK FROM ReportGroup WHERE ReportGroupPK = 12)
BEGIN
INSERT INTO ReportGroup ([ReportGroupID], [ReportGroupName], [ModuleID], [Sequence], [Icon], [RepairCenterPK], [IsUserGroup], [IsBatchGroup], [UDFChar1], [UDFChar2], [UDFChar3], [UDFChar4], [UDFChar5], [UDFDate1], [UDFDate2], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionAction], [RowVersionDate]) 
VALUES ('AS', 'Asset Reports', 'AS', 99, null, null, 0, 0, null, null, null, null, null, null, null, null, '', '', '_MC', 'CREATE', getdate())
PRINT 'Inserting ReportGroup Row - ' + 'Asset Reports'

IF (@@error > 0) SET @errorcount = @errorcount + 1

SET @rptGroupPK = @@IDENTITY
END
ELSE
BEGIN
UPDATE	ReportGroup
SET	[ReportGroupID]='AS', [ReportGroupName]='Asset Reports', [ModuleID]='AS', [Sequence]=99, [Icon]=null, [RepairCenterPK]=null, [IsUserGroup]=0, [IsBatchGroup]=0, [UDFChar1]=null, [UDFChar2]=null, [UDFChar3]=null, [UDFChar4]=null, [UDFChar5]=null, [UDFDate1]=null, [UDFDate2]=null, [DemoLaborPK]=null, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionAction]='EDIT', [RowVersionDate]=getdate()
WHERE	ReportGroupPK = 12
PRINT 'Updating ReportGroup Row - ' + 'Asset Reports'

SET @rptGroupPK = 12
END

IF (@@error > 0) SET @errorcount = @errorcount + 1

INSERT INTO Report_ReportGroup ([ReportPK], [ReportGroupPK], [DemoLaborPK], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, @rptGroupPK, '', '', '_MC', getdate())
PRINT 'Inserting Report_ReportGroup Row - ' + 'Asset Reports'
IF (@@error > 0) SET @errorcount = @errorcount + 1

/* Make sure the report group actually exists */
IF NOT EXISTS (SELECT ReportGroupPK FROM ReportGroup WHERE ReportGroupPK = 23)
BEGIN
INSERT INTO ReportGroup ([ReportGroupID], [ReportGroupName], [ModuleID], [Sequence], [Icon], [RepairCenterPK], [IsUserGroup], [IsBatchGroup], [UDFChar1], [UDFChar2], [UDFChar3], [UDFChar4], [UDFChar5], [UDFDate1], [UDFDate2], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionAction], [RowVersionDate]) 
VALUES ('FA', 'Failure Reports', 'FA', 99, null, null, 0, 0, null, null, null, null, null, null, null, null, '', '', '_MC', 'CREATE', getdate())
PRINT 'Inserting ReportGroup Row - ' + 'Failure Reports'

IF (@@error > 0) SET @errorcount = @errorcount + 1

SET @rptGroupPK = @@IDENTITY
END
ELSE
BEGIN
UPDATE	ReportGroup
SET	[ReportGroupID]='FA', [ReportGroupName]='Failure Reports', [ModuleID]='FA', [Sequence]=99, [Icon]=null, [RepairCenterPK]=null, [IsUserGroup]=0, [IsBatchGroup]=0, [UDFChar1]=null, [UDFChar2]=null, [UDFChar3]=null, [UDFChar4]=null, [UDFChar5]=null, [UDFDate1]=null, [UDFDate2]=null, [DemoLaborPK]=null, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionAction]='EDIT', [RowVersionDate]=getdate()
WHERE	ReportGroupPK = 23
PRINT 'Updating ReportGroup Row - ' + 'Failure Reports'

SET @rptGroupPK = 23
END

IF (@@error > 0) SET @errorcount = @errorcount + 1

INSERT INTO Report_ReportGroup ([ReportPK], [ReportGroupPK], [DemoLaborPK], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, @rptGroupPK, '', '', '_MC', getdate())
PRINT 'Inserting Report_ReportGroup Row - ' + 'Failure Reports'
IF (@@error > 0) SET @errorcount = @errorcount + 1


/* ==================================================
DELETE AND INSERT REPORT TABLES - (Sandvik_FtpErrors)
=================================================== */
DELETE FROM ReportTables WHERE ReportPK = @rptPK

PRINT 'Deleting ReportTables Rows - Sandvik_FtpErrors'
IF (@@error > 0) SET @errorcount = @errorcount + 1

INSERT INTO ReportTables ([ReportPK], [RFTable], [Alias], [DisplayOrder], [LabelOverride], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, 'MC_InterfaceLog', null, 0, null, null, '', '', '_MC', getdate())
PRINT 'Inserting ReportTables Row - MC_InterfaceLog'
IF (@@error > 0) SET @errorcount = @errorcount + 1


/* ==================================================
UPDATE OR INSERT REPORT CRITERIA - (Sandvik_FtpErrors)
=================================================== */	
IF EXISTS (
SELECT ReportCriteriaPK FROM ReportCriteria WITH (NOLOCK) WHERE ReportPK = @rptPK AND DisplayTable = 'MC_InterfaceLog' AND DisplayField = 'ProcessDate')
BEGIN
UPDATE ReportCriteria
SET	[SQLWhereTable]='MC_InterfaceLog', [SQLWhereField]='ProcessDate', [CritName]=null, [Operator]='is within', [isMulti]=0, [AskLater]=1, [LabelOverride]=null, [DisplayOrder]=1, [FK_LookupOverride]=null, [DemoLaborPK]=0, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionDate]=getdate()
WHERE	ReportPK = @rptPK AND DisplayTable = 'MC_InterfaceLog' AND DisplayField = 'ProcessDate' 
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Updating ReportCriteria Row - MC_InterfaceLog.ProcessDate'
END
ELSE
BEGIN
INSERT INTO ReportCriteria ([ReportPK], [DisplayTable], [DisplayField], [SQLWhereTable], [SQLWhereField], [DefaultCritValue], [CritName], [Operator], [isMulti], [AskLater], [LabelOverride], [DisplayOrder], [FK_LookupOverride], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, 'MC_InterfaceLog', 'ProcessDate', 'MC_InterfaceLog', 'ProcessDate', 'CW', null, 'is within', 0, 1, null, 1, null, 0, null, 0, '_MC', getdate())
PRINT 'Inserting ReportCriteria Row - MC_InterfaceLog.ProcessDate'
IF (@@error > 0) SET @errorcount = @errorcount + 1

END

IF EXISTS (
SELECT ReportCriteriaPK FROM ReportCriteria WITH (NOLOCK) WHERE ReportPK = @rptPK AND DisplayTable = 'MC_InterfaceLog' AND DisplayField = 'Processed')
BEGIN
UPDATE ReportCriteria
SET	[SQLWhereTable]='MC_InterfaceLog', [SQLWhereField]='Processed', [CritName]=null, [Operator]='is', [isMulti]=0, [AskLater]=1, [LabelOverride]=null, [DisplayOrder]=0, [FK_LookupOverride]=null, [DemoLaborPK]=null, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionDate]=getdate()
WHERE	ReportPK = @rptPK AND DisplayTable = 'MC_InterfaceLog' AND DisplayField = 'Processed' 
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Updating ReportCriteria Row - MC_InterfaceLog.Processed'
END
ELSE
BEGIN
INSERT INTO ReportCriteria ([ReportPK], [DisplayTable], [DisplayField], [SQLWhereTable], [SQLWhereField], [DefaultCritValue], [CritName], [Operator], [isMulti], [AskLater], [LabelOverride], [DisplayOrder], [FK_LookupOverride], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, 'MC_InterfaceLog', 'Processed', 'MC_InterfaceLog', 'Processed', 'N', null, 'is', 0, 1, null, 0, null, null, null, 0, '_MC', getdate())
PRINT 'Inserting ReportCriteria Row - MC_InterfaceLog.Processed'
IF (@@error > 0) SET @errorcount = @errorcount + 1

END

/* ==================================================
UPDATE OR INSERT REPORT FIELDS - (Sandvik_FtpErrors)
=================================================== */	
IF EXISTS (
SELECT ReportFieldPK FROM ReportFields WITH (NOLOCK) WHERE ReportPK = @rptPK AND RFTable = 'MC_InterfaceLog' AND RFField = 'ErrorMessage')
BEGIN
SELECT @tempPK = ReportFieldPK FROM ReportFields WITH (NOLOCK) WHERE ReportPK = @rptPK AND RFTable = 'MC_InterfaceLog' AND RFField = 'ErrorMessage'
UPDATE ReportFields
SET [ReportPK]=@rptPK, [DataDictPK]=10464, [AGFunction]=null, [Alias]=null, [DisplayOrder]=2, [Display]=1, [NotUserSelectable]=0, [LabelOverride]='Error description', [TotalIfSelected]=0, [BlankLineIfSelected]=0, [UseCustomExpression]=0, [CustomExpression]=null, [DemoLaborPK]=0, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionDate]=getdate(), [SLAction]='  ', [SLModuleID]=null, [SLPKField]=null, [SLReportID]=null, [SLCustomAction]=null, [SLToolTip]=null 
WHERE [ReportPK]=@rptPK AND [RFTable]='MC_InterfaceLog' AND [RFField]='ErrorMessage'
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Updating ReportFields Row - MC_InterfaceLog.ErrorMessage'
END
ELSE
BEGIN
INSERT INTO ReportFields ([ReportPK], [DataDictPK], [AGFunction], [RFTable], [RFField], [Alias], [DisplayOrder], [Display], [NotUserSelectable], [LabelOverride], [TotalIfSelected], [BlankLineIfSelected], [UseCustomExpression], [CustomExpression], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate], [SLAction], [SLModuleID], [SLPKField], [SLReportID], [SLCustomAction], [SLTooltip])
VALUES(@rptPK, 10464, null, 'MC_InterfaceLog', 'ErrorMessage', null, 2, 1, 0, 'Error description', 0, 0, 0, null, 0, '', '', '_MC', getdate(), '  ', null, null, null, null, null)
PRINT 'Inserting ReportFields Row - MC_InterfaceLog.ErrorMessage'
IF (@@error > 0) 
BEGIN
SET @errorcount = @errorcount + 1
SET @tempPK = -1
END
ELSE
BEGIN
SET @tempPK = @@IDENTITY
END
END

IF @version >= 4.2 
BEGIN
SET @dSQL = 'UPDATE ReportFields SET [Alignment] = null, [AdditionalWidth] = null, [PivotSetup] = ''  '', [AddPivotColumnsWithNoDataFrom] = null, [AddPivotColumnsWithNoDataFromCustom] = null, [Data_Type_Override] = null, [ColumnFormat] = null, [ColumnCS] = null, [GroupByCustomExpression] = 1 WHERE ReportFieldPK = ' + CAST(@tempPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1
END
IF EXISTS (
SELECT ReportFieldPK FROM ReportFields WITH (NOLOCK) WHERE ReportPK = @rptPK AND RFTable = 'MC_InterfaceLog' AND RFField = 'FileName')
BEGIN
SELECT @tempPK = ReportFieldPK FROM ReportFields WITH (NOLOCK) WHERE ReportPK = @rptPK AND RFTable = 'MC_InterfaceLog' AND RFField = 'FileName'
UPDATE ReportFields
SET [ReportPK]=@rptPK, [DataDictPK]=9786, [AGFunction]=null, [Alias]=null, [DisplayOrder]=0, [Display]=1, [NotUserSelectable]=0, [LabelOverride]='File Name', [TotalIfSelected]=0, [BlankLineIfSelected]=0, [UseCustomExpression]=0, [CustomExpression]=null, [DemoLaborPK]=0, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionDate]=getdate(), [SLAction]='  ', [SLModuleID]=null, [SLPKField]=null, [SLReportID]=null, [SLCustomAction]=null, [SLToolTip]=null 
WHERE [ReportPK]=@rptPK AND [RFTable]='MC_InterfaceLog' AND [RFField]='FileName'
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Updating ReportFields Row - MC_InterfaceLog.FileName'
END
ELSE
BEGIN
INSERT INTO ReportFields ([ReportPK], [DataDictPK], [AGFunction], [RFTable], [RFField], [Alias], [DisplayOrder], [Display], [NotUserSelectable], [LabelOverride], [TotalIfSelected], [BlankLineIfSelected], [UseCustomExpression], [CustomExpression], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate], [SLAction], [SLModuleID], [SLPKField], [SLReportID], [SLCustomAction], [SLTooltip])
VALUES(@rptPK, 9786, null, 'MC_InterfaceLog', 'FileName', null, 0, 1, 0, 'File Name', 0, 0, 0, null, 0, '', '', '_MC', getdate(), '  ', null, null, null, null, null)
PRINT 'Inserting ReportFields Row - MC_InterfaceLog.FileName'
IF (@@error > 0) 
BEGIN
SET @errorcount = @errorcount + 1
SET @tempPK = -1
END
ELSE
BEGIN
SET @tempPK = @@IDENTITY
END
END

IF @version >= 4.2 
BEGIN
SET @dSQL = 'UPDATE ReportFields SET [Alignment] = null, [AdditionalWidth] = null, [PivotSetup] = ''  '', [AddPivotColumnsWithNoDataFrom] = null, [AddPivotColumnsWithNoDataFromCustom] = null, [Data_Type_Override] = null, [ColumnFormat] = null, [ColumnCS] = null, [GroupByCustomExpression] = 1 WHERE ReportFieldPK = ' + CAST(@tempPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1
END
IF EXISTS (
SELECT ReportFieldPK FROM ReportFields WITH (NOLOCK) WHERE ReportPK = @rptPK AND RFTable = 'MC_InterfaceLog' AND RFField = 'ProcessDate')
BEGIN
SELECT @tempPK = ReportFieldPK FROM ReportFields WITH (NOLOCK) WHERE ReportPK = @rptPK AND RFTable = 'MC_InterfaceLog' AND RFField = 'ProcessDate'
UPDATE ReportFields
SET [ReportPK]=@rptPK, [DataDictPK]=9788, [AGFunction]=null, [Alias]=null, [DisplayOrder]=1, [Display]=1, [NotUserSelectable]=0, [LabelOverride]=null, [TotalIfSelected]=0, [BlankLineIfSelected]=0, [UseCustomExpression]=0, [CustomExpression]=null, [DemoLaborPK]=null, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionDate]=getdate(), [SLAction]=null, [SLModuleID]=null, [SLPKField]=null, [SLReportID]=null, [SLCustomAction]=null, [SLToolTip]=null 
WHERE [ReportPK]=@rptPK AND [RFTable]='MC_InterfaceLog' AND [RFField]='ProcessDate'
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Updating ReportFields Row - MC_InterfaceLog.ProcessDate'
END
ELSE
BEGIN
INSERT INTO ReportFields ([ReportPK], [DataDictPK], [AGFunction], [RFTable], [RFField], [Alias], [DisplayOrder], [Display], [NotUserSelectable], [LabelOverride], [TotalIfSelected], [BlankLineIfSelected], [UseCustomExpression], [CustomExpression], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate], [SLAction], [SLModuleID], [SLPKField], [SLReportID], [SLCustomAction], [SLTooltip])
VALUES(@rptPK, 9788, null, 'MC_InterfaceLog', 'ProcessDate', null, 1, 1, 0, null, 0, 0, 0, null, null, '', '', '_MC', getdate(), null, null, null, null, null, null)
PRINT 'Inserting ReportFields Row - MC_InterfaceLog.ProcessDate'
IF (@@error > 0) 
BEGIN
SET @errorcount = @errorcount + 1
SET @tempPK = -1
END
ELSE
BEGIN
SET @tempPK = @@IDENTITY
END
END

IF @version >= 4.2 
BEGIN
SET @dSQL = 'UPDATE ReportFields SET [Alignment] = null, [AdditionalWidth] = null, [PivotSetup] = null, [AddPivotColumnsWithNoDataFrom] = null, [AddPivotColumnsWithNoDataFromCustom] = null, [Data_Type_Override] = null, [ColumnFormat] = null, [ColumnCS] = null, [GroupByCustomExpression] = 1 WHERE ReportFieldPK = ' + CAST(@tempPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1
END
/* =====================================================
UPDATE REPORTS COLS IN DATA DICT - (Sandvik_FtpErrors)
===================================================== */
UPDATE DataDict 
SET	[REPORT_NOSELECT]=null, [REPORT_EDIT]=null, [REPORT_LABEL]=null, [TOTALIFSELECTED]=0 
WHERE	TABLE_NAME = 'MC_InterfaceLog' AND COLUMN_NAME = 'ErrorMessage' 

PRINT 'Updating DataDict Row - MC_InterfaceLog.ErrorMessage'
IF (@@error > 0) SET @errorcount = @errorcount + 1

UPDATE DataDict 
SET	[REPORT_NOSELECT]=null, [REPORT_EDIT]=null, [REPORT_LABEL]=null, [TOTALIFSELECTED]=0 
WHERE	TABLE_NAME = 'MC_InterfaceLog' AND COLUMN_NAME = 'ProcessDate' 

PRINT 'Updating DataDict Row - MC_InterfaceLog.ProcessDate'
IF (@@error > 0) SET @errorcount = @errorcount + 1

END
/* REPORT DID NOT EXIST, CREATE ALL RECORDS */
ELSE

BEGIN
PRINT '*******************************************************'
PRINT 'Report Does Not Exist - Inserting...'
PRINT 'Report: FTP Import Errors'
PRINT '*******************************************************'

/* =====================================================
INSERT REPORT RECORDS - (Sandvik_FtpErrors)
===================================================== */
/* INSERT Main Report */
INSERT INTO Reports ([ReportIDPriorToCopy], [ReportID], [ReportName], [ReportDesc], [RepairCenterPK], [Sort1], [Sort2], [Sort3], [Sort4], [Sort5], [Sort1DESC], [Sort2DESC], [Sort3DESC], [Sort4DESC], [Sort5DESC], [Group1], [Group2], [Group3], [Group4], [Group5], [Header1], [Header2], [Header3], [Header4], [Header5], [GroupHeader1], [GroupHeader2], [GroupHeader3], [GroupHeader4], [GroupHeader5], [Total1], [Total2], [Total3], [Total4], [Total5], [Chart], [ChartName], [ChartField], [ChartSize], [ReportFile], [FromSQL], [JoinSQL], [WhereSQL], [GroupBy], [hits], [Sequence], [Layout], [VertCols], [PageBreakEachRecord], [Custom], [ReportCopy], [MCRegistrationDB], [PrintCriteria], [Active], [UDFChar1], [UDFChar2], [UDFChar3], [UDFChar4], [UDFChar5], [UDFDate1], [UDFDate2], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionAction], [RowVersionDate], [ChartFunction], [ChartFunctionField], [NoDetail], [PB1], [PB2], [PB3], [PB4], [PB5], [SLDefault], [SLType], [SLAction], [SLModuleID], [SLPKField], [SLReportID], [SLCustomAction], [SLTooltip], [SDDisplay], [SDModuleID], [SDPKField], [SmartEmail], [ChartPosition], [ChartFormat], [ChartSQL], [Chart2], [ChartName2], [ChartField2], [ChartSize2], [ChartFormat2], [ChartFunction2], [ChartFunctionField2], [ChartPosition2], [ChartSQL2], [Chart3], [ChartName3], [ChartField3], [ChartSize3], [ChartFormat3], [ChartFunction3], [ChartFunctionField3], [ChartPosition3], [ChartSQL3], [ChartOnly], [NoHeader], [HavingSQL], [SRID1], [SRPKField1], [SRID2], [SRPKField2], [SRID3], [SRPKField3], [SRID4], [SRPKField4], [SRID5], [SRPKField5], [ReportPageSize], [ReportWidth], [PhotoCriteria], [ReportStyleName], [UsedFor], [SmartEmailLaborPK], [SCDefault], [SCField1], [SCField2], [SCField3], [ReportStyleFontSize], [ReportStyleFontColor], [ReportStyleFontFamily] )
VALUES('WOList', 'Sandvik_FtpErrors', 'FTP Import Errors', null, null, 'MC_InterfaceLog.ProcessDate', null, null, null, null, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, null, null, null, null, null, 0, 0, 0, 0, 0, null, 'Work Order Status', 'AUTO_SORT1', 'L', 'rpt_generic1.asp', 'FROM MC_InterfaceLog', null, null, 0, 0, 0, 'hor', 1, 0, 0, 1, 0, 0, 1, null, null, null, 'N', null, null, null, null, '', '', '_MC', 'CREATE', getdate(), 'C', 'NONE', 0, 0, 0, 0, 0, 0, 0, ' ', 'PW', 'WO', null, null, null, null, '     ', '  ', null, 0, 'T', 'F', null, null, null, null, null, 'I', null, null, null, null, null, null, null, null, 'I', null, null, null, null, 0, 0, null, null, null, null, null, null, null, null, null, null, null, 'Default', '80%', 0, 'Gradient - Blue', 'REPORTS', 0, 'H', null, null, null, null, null, null)

PRINT 'Inserting Report - Sandvik_FtpErrors'
IF (@@error > 0) SET @errorcount = @errorcount + 1

SET @rptPK = @@Identity

IF @version >= 3.0 
BEGIN
SET @dSQL = 'UPDATE Reports SET [HavingSQL]=null, [DisplayPivotBar]=0, [DisplayColumnLines]=0, [DisplayTitleonPageBreak]=0, [DisplayFormatCriteria]=1, [R1T]='' '', [R1O]=null, [R1V1]=null, [R1V2]=null, [R1A]=0, [R1L]=0, [R1F]=0, [R1CS]=''font-family: Arial;font-size: 8pt;color: #000000;text-align: left;'', [R1AF]=''C'', [R2T]='' '', [R2O]=null, [R2V1]=null, [R2V2]=null, [R2A]=0, [R2L]=0, [R2F]=0, [R2CS]=''border: #0066CC 2px solid;'', [R2AF]=''C'', [R3T]='' '', [R3O]=null, [R3V1]=null, [R3V2]=null, [R3A]=0, [R3L]=0, [R3F]=0, [R3CS]=''border: #0066CC 2px solid;'', [R3AF]=''C'' WHERE ReportPK = ' + CAST(@rptPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1

END

IF @version >= 4.2 
BEGIN
SET @dSQL = 'UPDATE Reports SET [DisplayDescription]=0 WHERE ReportPK = ' + CAST(@rptPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1

END

/* ==================================================
UPDATE OR INSERT REPORT GROUPS - (Sandvik_FtpErrors)
=================================================== */

/* Make sure the report group actually exists */
IF NOT EXISTS (SELECT ReportGroupPK FROM ReportGroup WHERE ReportGroupPK = 12)
BEGIN
INSERT INTO ReportGroup ([ReportGroupID], [ReportGroupName], [ModuleID], [Sequence], [Icon], [RepairCenterPK], [IsUserGroup], [IsBatchGroup], [UDFChar1], [UDFChar2], [UDFChar3], [UDFChar4], [UDFChar5], [UDFDate1], [UDFDate2], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionAction], [RowVersionDate]) 
VALUES ('AS', 'Asset Reports', 'AS', 99, null, null, 0, 0, null, null, null, null, null, null, null, null, '', '', '_MC', 'CREATE', getdate())
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Inserting ReportGroup Row - ' + 'Asset Reports'

SET @rptGroupPK = @@IDENTITY
END
ELSE
BEGIN
UPDATE	ReportGroup
SET	[ReportGroupID]='AS', [ReportGroupName]='Asset Reports', [ModuleID]='AS', [Sequence]=99, [Icon]=null, [RepairCenterPK]=null, [IsUserGroup]=0, [IsBatchGroup]=0, [UDFChar1]=null, [UDFChar2]=null, [UDFChar3]=null, [UDFChar4]=null, [UDFChar5]=null, [UDFDate1]=null, [UDFDate2]=null, [DemoLaborPK]=null, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionAction]='EDIT', [RowVersionDate]=getdate()
WHERE	ReportGroupPK = 12
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Updating ReportGroup Row - ' + 'Asset Reports'

SET @rptGroupPK = 12
END

INSERT INTO Report_ReportGroup ([ReportPK], [ReportGroupPK], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, @rptGroupPK, null, '', '', '_MC', getdate())
PRINT 'Inserting Report_ReportGroup Row - ' + 'Asset Reports'
IF (@@error > 0) SET @errorcount = @errorcount + 1

/* Make sure the report group actually exists */
IF NOT EXISTS (SELECT ReportGroupPK FROM ReportGroup WHERE ReportGroupPK = 23)
BEGIN
INSERT INTO ReportGroup ([ReportGroupID], [ReportGroupName], [ModuleID], [Sequence], [Icon], [RepairCenterPK], [IsUserGroup], [IsBatchGroup], [UDFChar1], [UDFChar2], [UDFChar3], [UDFChar4], [UDFChar5], [UDFDate1], [UDFDate2], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionAction], [RowVersionDate]) 
VALUES ('FA', 'Failure Reports', 'FA', 99, null, null, 0, 0, null, null, null, null, null, null, null, null, '', '', '_MC', 'CREATE', getdate())
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Inserting ReportGroup Row - ' + 'Failure Reports'

SET @rptGroupPK = @@IDENTITY
END
ELSE
BEGIN
UPDATE	ReportGroup
SET	[ReportGroupID]='FA', [ReportGroupName]='Failure Reports', [ModuleID]='FA', [Sequence]=99, [Icon]=null, [RepairCenterPK]=null, [IsUserGroup]=0, [IsBatchGroup]=0, [UDFChar1]=null, [UDFChar2]=null, [UDFChar3]=null, [UDFChar4]=null, [UDFChar5]=null, [UDFDate1]=null, [UDFDate2]=null, [DemoLaborPK]=null, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionAction]='EDIT', [RowVersionDate]=getdate()
WHERE	ReportGroupPK = 23
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Updating ReportGroup Row - ' + 'Failure Reports'

SET @rptGroupPK = 23
END

INSERT INTO Report_ReportGroup ([ReportPK], [ReportGroupPK], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, @rptGroupPK, null, '', '', '_MC', getdate())
PRINT 'Inserting Report_ReportGroup Row - ' + 'Failure Reports'
IF (@@error > 0) SET @errorcount = @errorcount + 1

/* ==================================================
INSERT REPORT TABLES- (Sandvik_FtpErrors)
=================================================== */	
INSERT INTO ReportTables ([ReportPK], [RFTable], [Alias], [DisplayOrder], [LabelOverride], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, 'MC_InterfaceLog', null, 0, null, null, '', '', '_MC', getdate())
PRINT 'Inserting ReportTables Row - MC_InterfaceLog'
IF (@@error > 0) SET @errorcount = @errorcount + 1


/* ==================================================
INSERT REPORT CRITERIA - (Sandvik_FtpErrors)
=================================================== */	
INSERT INTO ReportCriteria ([ReportPK], [DisplayTable], [DisplayField], [SQLWhereTable], [SQLWhereField], [DefaultCritValue], [CritName], [Operator], [isMulti], [AskLater], [LabelOverride], [DisplayOrder], [FK_LookupOverride], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, 'MC_InterfaceLog', 'ProcessDate', 'MC_InterfaceLog', 'ProcessDate', 'CW', null, 'is within', 0, 1, null, 1, null, 0, '', '', '_MC', getdate())
PRINT 'Inserting ReportCriteria Row - MC_InterfaceLog.ProcessDate'
IF (@@error > 0) SET @errorcount = @errorcount + 1

INSERT INTO ReportCriteria ([ReportPK], [DisplayTable], [DisplayField], [SQLWhereTable], [SQLWhereField], [DefaultCritValue], [CritName], [Operator], [isMulti], [AskLater], [LabelOverride], [DisplayOrder], [FK_LookupOverride], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, 'MC_InterfaceLog', 'Processed', 'MC_InterfaceLog', 'Processed', 'N', null, 'is', 0, 1, null, 0, null, null, '', '', '_MC', getdate())
PRINT 'Inserting ReportCriteria Row - MC_InterfaceLog.Processed'
IF (@@error > 0) SET @errorcount = @errorcount + 1

/* ==================================================
INSERT REPORT FIELDS - (Sandvik_FtpErrors)
=================================================== */	
INSERT INTO ReportFields ([ReportPK], [DataDictPK], [AGFunction], [RFTable], [RFField], [Alias], [DisplayOrder], [Display], [NotUserSelectable], [LabelOverride], [TotalIfSelected], [BlankLineIfSelected], [UseCustomExpression], [CustomExpression], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate], [SLAction], [SLModuleID], [SLPKField], [SLReportID], [SLCustomAction], [SLTooltip])
VALUES(@rptPK, 10464, null, 'MC_InterfaceLog', 'ErrorMessage', null, 2, 1, 0, 'Error description', 0, 0, 0, null, 0, '', '', '_MC', getdate(), '  ', null, null, null, null, null)
PRINT 'Inserting ReportFields Row - MC_InterfaceLog.ErrorMessage'
IF (@@error > 0) SET @errorcount = @errorcount + 1

IF @version >= 4.2 
BEGIN
SET @tempPK = @@IDENTITY
SET @dSQL = 'UPDATE ReportFields SET [Alignment] = null, [AdditionalWidth] = null, [PivotSetup] = ''  '', [AddPivotColumnsWithNoDataFrom] = null, [AddPivotColumnsWithNoDataFromCustom] = null, [Data_Type_Override] = null, [ColumnFormat] = null, [ColumnCS] = null, [GroupByCustomExpression] = 1 WHERE ReportFieldPK = ' + CAST(@tempPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1
END
INSERT INTO ReportFields ([ReportPK], [DataDictPK], [AGFunction], [RFTable], [RFField], [Alias], [DisplayOrder], [Display], [NotUserSelectable], [LabelOverride], [TotalIfSelected], [BlankLineIfSelected], [UseCustomExpression], [CustomExpression], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate], [SLAction], [SLModuleID], [SLPKField], [SLReportID], [SLCustomAction], [SLTooltip])
VALUES(@rptPK, 9786, null, 'MC_InterfaceLog', 'FileName', null, 0, 1, 0, 'File Name', 0, 0, 0, null, 0, '', '', '_MC', getdate(), '  ', null, null, null, null, null)
PRINT 'Inserting ReportFields Row - MC_InterfaceLog.FileName'
IF (@@error > 0) SET @errorcount = @errorcount + 1

IF @version >= 4.2 
BEGIN
SET @tempPK = @@IDENTITY
SET @dSQL = 'UPDATE ReportFields SET [Alignment] = null, [AdditionalWidth] = null, [PivotSetup] = ''  '', [AddPivotColumnsWithNoDataFrom] = null, [AddPivotColumnsWithNoDataFromCustom] = null, [Data_Type_Override] = null, [ColumnFormat] = null, [ColumnCS] = null, [GroupByCustomExpression] = 1 WHERE ReportFieldPK = ' + CAST(@tempPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1
END
INSERT INTO ReportFields ([ReportPK], [DataDictPK], [AGFunction], [RFTable], [RFField], [Alias], [DisplayOrder], [Display], [NotUserSelectable], [LabelOverride], [TotalIfSelected], [BlankLineIfSelected], [UseCustomExpression], [CustomExpression], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate], [SLAction], [SLModuleID], [SLPKField], [SLReportID], [SLCustomAction], [SLTooltip])
VALUES(@rptPK, 9788, null, 'MC_InterfaceLog', 'ProcessDate', null, 1, 1, 0, null, 0, 0, 0, null, null, '', '', '_MC', getdate(), null, null, null, null, null, null)
PRINT 'Inserting ReportFields Row - MC_InterfaceLog.ProcessDate'
IF (@@error > 0) SET @errorcount = @errorcount + 1

IF @version >= 4.2 
BEGIN
SET @tempPK = @@IDENTITY
SET @dSQL = 'UPDATE ReportFields SET [Alignment] = null, [AdditionalWidth] = null, [PivotSetup] = null, [AddPivotColumnsWithNoDataFrom] = null, [AddPivotColumnsWithNoDataFromCustom] = null, [Data_Type_Override] = null, [ColumnFormat] = null, [ColumnCS] = null, [GroupByCustomExpression] = 1 WHERE ReportFieldPK = ' + CAST(@tempPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1
END
/* =====================================================
UPDATE REPORTS COLS IN DATA DICT - (Sandvik_FtpErrors)
===================================================== */

UPDATE DataDict 
SET	[REPORT_NOSELECT]=null, [REPORT_EDIT]=null, [REPORT_LABEL]=null, [TOTALIFSELECTED]=0 
WHERE	TABLE_NAME = 'MC_InterfaceLog' AND COLUMN_NAME = 'ErrorMessage' 

PRINT 'Updating DataDict Row - MC_InterfaceLog.ErrorMessage'
if (@@error > 0) Set @errorcount = @errorcount + 1

UPDATE DataDict 
SET	[REPORT_NOSELECT]=null, [REPORT_EDIT]=null, [REPORT_LABEL]=null, [TOTALIFSELECTED]=0 
WHERE	TABLE_NAME = 'MC_InterfaceLog' AND COLUMN_NAME = 'ProcessDate' 

PRINT 'Updating DataDict Row - MC_InterfaceLog.ProcessDate'
if (@@error > 0) Set @errorcount = @errorcount + 1

END

PRINT '*******************************************************'
IF (@errorcount > 0) PRINT @errorcount + ' Error(s) Occurred: Sandvik_FtpErrors'
ELSE PRINT 'No Errors Occurred: Sandvik_FtpErrors'

GO

SET XACT_ABORT OFF
SET ARITHABORT OFF

/* DECLARE GLOBAL VARIABLES */
DECLARE @rptPK int
DECLARE @rptGroupPK int
DECLARE @tempPK int
DECLARE @errorcount int
DECLARE @dsql varchar(8000)
DECLARE @version decimal(9,2)

SELECT @version = CAST(REPLACE(LOWER(schema_version), 'sp', '') AS decimal(9,2)) FROM _schema

/* =====================================================
UPDATE REPORTS STYLES AND GROUP HEADERS - (Sandvik_FtpSuccess)
===================================================== */

IF @version >= 3.0 
BEGIN 
EXEC('PRINT ''Version 3.0 or greater detected. Importing Report Styles and Group Headers''
DECLARE @tempPK int
IF NOT EXISTS(SELECT ReportStyleName FROM ReportStyle WHERE ReportStyleName = ''Gradient - Blue'')
BEGIN 
INSERT INTO ReportStyle (ReportStyleName, ReportStyleDesc, ReportStyleCSS, IsDefault, IsBase, RowVersionIPAddress, RowVersionUserPK, RowVersionInitials, RowVersionAction, RowVersionDate) VALUES(''Gradient - Blue'', null, '' .pageselect{FONT-SIZE: 9pt; COLOR: #333333; FONT-FAMILY: Arial}
 .heading {background-color:#ffffff; cursor:pointer; FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: royalblue; FONT-FAMILY: Arial; z-index: 2500;} 
 .legendHeader {FONT-WEIGHT: bold; FONT-SIZE: 14px; COLOR: #333333; FONT-FAMILY: Arial}
 .normaltext {FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial}
 .labels {FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: royalblue; FONT-FAMILY: Arial}
 .assetUP {FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: green; FONT-FAMILY: Arial}
 .assetDOWN {FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #DD0000; FONT-FAMILY: Arial}
 .asset {FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Arial}
 .data {FONT-SIZE: 12px; COLOR: #494949; FONT-FAMILY: Arial}
 .data_underline {BORDER-RIGHT: medium none; BORDER-TOP: medium none; FONT-SIZE: 12px; BORDER-LEFT: medium none; COLOR: #494949; BORDER-BOTTOM: #333333 1px solid; FONT-FAMILY: Arial}
 .bottomline {BORDER-RIGHT: medium none; BORDER-TOP: medium none; BORDER-LEFT: medium none; BORDER-BOTTOM: #333333 1px solid}
 .buttons {FONT-SIZE: 12px; WIDTH: 80px; cursor: pointer; COLOR: #333333; FONT-FAMILY: Arial}
 .subtotal {BORDER-RIGHT: medium none; BORDER-TOP: #C0C0C0 1px solid; FONT-SIZE: 12px; BORDER-LEFT: medium none; COLOR: #333333; BORDER-BOTTOM: medium none; FONT-FAMILY: Arial}
 .bodyclasspreview {background-color:#ffffff; padding:10px; scrollbar-base-color: #FBFBFB; font-size:8pt; font-family:Arial; color:#000000;}
 .bodyclasspreviewinwo {background-color:#ffffff; padding-right:10px; scrollbar-base-color: #FBFBFB; font-size:8pt; font-family:Arial; color:#000000;}
 .bodyclassprint {background-color:#ffffff; PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; PADDING-TOP: 0px; font-size:8pt; font-family:Arial; color:#000000;}
 .bodyclassemail {background-color:#ffffff; PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 20px; PADDING-TOP: 0px; font-size:8pt; font-family:Arial; color:#000000;}
 .group1 {padding-left:0px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; FONT-WEIGHT: Bold; BACKGROUND-COLOR: #acc5e7;}
 .group2 {padding-left:10px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; FONT-WEIGHT: Bold; BACKGROUND-COLOR: #c7d7ed;}
 .group3 {padding-left:20px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; FONT-WEIGHT: Bold; BACKGROUND-COLOR: #dce8f4;}
 .group4 {padding-left:30px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; FONT-WEIGHT: Bold; BACKGROUND-COLOR: #ecf1fb;}
 .group5 {padding-left:40px; FONT-SIZE: 12px; COLOR: royalblue; FONT-FAMILY: Arial; FONT-WEIGHT: Bold; BACKGROUND-COLOR: #ffffff;}
 .groupheader {FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; FONT-WEIGHT: Bold;}
 .normalright {BORDER-RIGHT: #c0c0c0 1px solid; BORDER-TOP: #c0c0c0 1px solid; PADDING-LEFT: 1px; FONT-WEIGHT: normal; FONT-SIZE: 8pt; MARGIN-BOTTOM: 1px; BORDER-LEFT: #c0c0c0 1px solid; COLOR: #000000; BORDER-BOTTOM: #c0c0c0 1px solid; FONT-FAMILY: Arial; BACKGROUND-COLOR: #ffffff; TEXT-ALIGN: right}
 .clsBtnUp {cursor: pointer; color: black; font-weight: normal; border-top:1px solid #ffffff;border-left:1px solid #ffffff;border-right:1px solid #B4B4B4;border-bottom:1px solid #B4B4B4;padding-right:2px;}
 .clsBtnDown {cursor: pointer; color: black; font-weight: normal; border-right:1px solid #ffffff;border-bottom:1px solid #ffffff;border-top:1px solid #B4B4B4;border-left:1px solid #B4B4B4;padding-right:2px;}
 .clsBtnOff {color: black; font-weight: normal; tab-index: 0; border:1px solid transparent; padding-right:2px;}
 .actionbarlabel {float:right;margin-top:6px;padding-left:5px;padding-right:5px;font-size:8pt;font-family:Arial;color:#000000;}
 INPUT {padding-left:3px;}
 A:link {FONT-SIZE: 8pt; cursor: pointer; COLOR: #315aad; FONT-FAMILY: Arial; BACKGROUND-COLOR: transparent;}
 A:visited {FONT-SIZE: 8pt; cursor: pointer; COLOR: #315aad; FONT-FAMILY: Arial; BACKGROUND-COLOR: transparent;}
 A:active {FONT-SIZE: 8pt; cursor: pointer; COLOR: #315aad; FONT-FAMILY: Arial; BACKGROUND-COLOR: transparent;}
 A:hover {COLOR: red;}
 fieldset {border: 1px solid #AAAAAB;}
 .buttonsdisabled { display: static; opacity: 0.4; cursor:pointer;	}
 .buttonsenabled { display:	; opacity: 1; cursor:pointer; }
 .normalrow {FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial }
 .tb {width:100%; PADDING-LEFT: 1px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; border: 1px solid #AAAAAB;}
 .tbf {BACKGROUND-COLOR: #ffffcc; width:100%; PADDING-LEFT: 1px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; border: 1px solid #AAAAAB;}
 .ta {width:200px; PADDING-LEFT: 1px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; border: 1px solid #AAAAAB;}
 .taf {BACKGROUND-COLOR: #ffffcc; width:200px; PADDING-LEFT: 1px; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; border: 1px solid #AAAAAB;}
 .cb {COLOR: #333333;}
 .HeaderRight {font-family:Arial;font-size:16px;color:#333333;font-weight:bold}
 .SubHeaderRight {font-family:Arial;font-size:11px;font-weight:normal}
 .SRInstructions {margin-top:5px;font-family:Arial;font-size:8pt;color:green;font-weight:bold}
 .verticalcolumn {border:1px solid #CCCCCC;}
 .mcpagebreak {page-break-before: always;}
 .ReportRow1 {background-color:#FFFFFF; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial }
 .ReportRow2 {background-color:#EFEFEF; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial }
 .ReportRowCrit1 {background-color:#FFFFFF; FONT-SIZE: 8pt; COLOR: #333333; FONT-FAMILY: Arial }
 .ReportRowCrit2 {background-color:#EFEFEF; FONT-SIZE: 8pt; COLOR: #333333; FONT-FAMILY: Arial }
 .SmartRow {background-color:#FFDF84; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial; cursor:pointer; }
 .SubReportRow {background-color:#DEEFC6; FONT-SIZE: 12px; COLOR: #333333; FONT-FAMILY: Arial }
 .ExpandCollapse {font-family:Arial;font-size:8pt;color:#333333;cursor:pointer;}'', 0, 0, '''', '''', ''_MC'', '''', getdate())
END 
')
END 
/* ============================================================
ReportID: Sandvik_FtpSuccess
Report Name: FTP Import Success
============================================================ */

IF EXISTS (
SELECT ReportName FROM Reports WITH (NOLOCK)
WHERE ReportID = 'Sandvik_FtpSuccess')

BEGIN

PRINT '*******************************************************'
PRINT 'Report Exists - Updating...'
PRINT 'Report: FTP Import Success'
PRINT '*******************************************************'

/* ================================================
UPDATE REPORT RECORDS - (Sandvik_FtpSuccess)
================================================ */
/* Set ReportPK for this Report */
SELECT @rptPK = ReportPK FROM Reports WITH (NOLOCK) WHERE ReportID='Sandvik_FtpSuccess'

/* Update Main Report Fields */
UPDATE Reports
SET	[ReportIDPriorToCopy]='NavGeoTabImpErrors', [ReportDesc]=null, [Sort1]='MC_InterfaceLog.PK', [Sort2]=null, [Sort3]=null, [Sort4]=null, [Sort5]=null, [Sort1DESC]=0, [Sort2DESC]=0, [Sort3DESC]=0, [Sort4DESC]=0, [Sort5DESC]=0, [Group1]=0, [Group2]=0, [Group3]=0, [Group4]=0, [Group5]=0, [Header1]=0, [Header2]=0, [Header3]=0, [Header4]=0, [Header5]=0, [GroupHeader1]=null, [GroupHeader2]=null, [GroupHeader3]=null, [GroupHeader4]=null, [GroupHeader5]=null, [Total1]=0, [Total2]=0, [Total3]=0, [Total4]=0, [Total5]=0, [Chart]=null, [ChartName]='Work Order Status', [ChartField]='AUTO_SORT1', [ChartSize]='L', [ReportFile]='rpt_generic1.asp', [FromSQL]='FROM MC_InterfaceLog', [JoinSQL]=null, [WhereSQL]=null, [GroupBy]=0, [hits]=2, [Sequence]=0, [Layout]='hor', [VertCols]=1, [PageBreakEachRecord]=0, [Custom]=0, [ReportCopy]=1, [MCRegistrationDB]=0, [PrintCriteria]=0, [Active]=1, [UDFChar1]=null, [UDFChar2]=null, [UDFChar3]=null, [UDFChar4]='N', [UDFChar5]=null, [UDFDate1]=null, [UDFDate2]=null, [DemoLaborPK]=null, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionAction]='EDIT', [RowVersionDate]=getdate() , [ChartFunction]='C', [ChartFunctionField]='NONE', [NoDetail]=0, [PB1]=0, [PB2]=0, [PB3]=0, [PB4]=0, [PB5]=0, [SLDefault]=1, [SLType]=' ', [SLAction]='PW', [SLModuleID]='WO', [SLPKField]=null, [SLReportID]=null, [SLCustomAction]=null, [SLTooltip]=null, [SDDisplay]='     ', [SDModuleID]='  ', [SDPKField]=null, [SmartEmail]=0, [ChartPosition]='T', [ChartFormat]='F', [ChartSQL]=null, [Chart2]=null, [ChartName2]=null, [ChartField2]=null, [ChartSize2]=null, [ChartFormat2]='I', [ChartFunction2]=null, [ChartFunctionField2]=null, [ChartPosition2]=null, [ChartSQL2]=null, [Chart3]=null, [ChartName3]=null, [ChartField3]=null, [ChartSize3]=null, [ChartFormat3]='I', [ChartFunction3]=null, [ChartFunctionField3]=null, [ChartPosition3]=null, [ChartSQL3]=null, [ChartOnly]=0, [NoHeader]=0, [SRID1]=null, [SRPKField1]=null, [SRID2]=null, [SRPKField2]=null, [SRID3]=null, [SRPKField3]=null, [SRID4]=null, [SRPKField4]=null, [SRID5]=null, [SRPKField5]=null, [ReportPageSize]='Default', [ReportWidth]='80%', [PhotoCriteria]=1, [ReportStyleName]='Gradient - Blue', [UsedFor]='REPORTS', [SmartEmailLaborPK]=0, [SCDefault]='H', [SCField1]=null, [SCField2]=null, [SCField3]=null, [ReportStyleFontSize]=null, [ReportStyleFontColor]=null, [ReportStyleFontFamily]=null 
WHERE ReportPK = @rptPK
IF (@@error > 0) SET @errorcount = @errorcount + 1

IF @version >= 3.0 
BEGIN
SET @dSQL = 'UPDATE Reports SET [HavingSQL]=null, [DisplayPivotBar]=0, [DisplayColumnLines]=0, [DisplayTitleonPageBreak]=0, [DisplayFormatCriteria]=1, [R1T]='' '', [R1O]=null, [R1V1]=null, [R1V2]=null, [R1A]=0, [R1L]=0, [R1F]=0, [R1CS]=''font-family: Arial;font-size: 8pt;color: #000000;text-align: left;'', [R1AF]=''C'', [R2T]='' '', [R2O]=null, [R2V1]=null, [R2V2]=null, [R2A]=0, [R2L]=0, [R2F]=0, [R2CS]=''border: #0066CC 2px solid;'', [R2AF]=''C'', [R3T]='' '', [R3O]=null, [R3V1]=null, [R3V2]=null, [R3A]=0, [R3L]=0, [R3F]=0, [R3CS]=''border: #0066CC 2px solid;'', [R3AF]=''C'' WHERE ReportPK = ' + CAST(@rptPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1

END

IF @version >= 4.2 
BEGIN
SET @dSQL = 'UPDATE Reports SET [DisplayDescription]=0 WHERE ReportPK = ' + CAST(@rptPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1

END

PRINT 'Updating Report - Sandvik_FtpSuccess'

/* ==================================================
DELETE AND INSERT REPORT GROUPS - (Sandvik_FtpSuccess)
=================================================== */
DELETE FROM Report_ReportGroup WHERE ReportPK = @rptPK

PRINT 'Deleting Report_ReportGroup Rows - Sandvik_FtpSuccess'

IF (@@error > 0) SET @errorcount = @errorcount + 1

/* Make sure the report group actually exists */
IF NOT EXISTS (SELECT ReportGroupPK FROM ReportGroup WHERE ReportGroupPK = 12)
BEGIN
INSERT INTO ReportGroup ([ReportGroupID], [ReportGroupName], [ModuleID], [Sequence], [Icon], [RepairCenterPK], [IsUserGroup], [IsBatchGroup], [UDFChar1], [UDFChar2], [UDFChar3], [UDFChar4], [UDFChar5], [UDFDate1], [UDFDate2], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionAction], [RowVersionDate]) 
VALUES ('AS', 'Asset Reports', 'AS', 99, null, null, 0, 0, null, null, null, null, null, null, null, null, '', '', '_MC', 'CREATE', getdate())
PRINT 'Inserting ReportGroup Row - ' + 'Asset Reports'

IF (@@error > 0) SET @errorcount = @errorcount + 1

SET @rptGroupPK = @@IDENTITY
END
ELSE
BEGIN
UPDATE	ReportGroup
SET	[ReportGroupID]='AS', [ReportGroupName]='Asset Reports', [ModuleID]='AS', [Sequence]=99, [Icon]=null, [RepairCenterPK]=null, [IsUserGroup]=0, [IsBatchGroup]=0, [UDFChar1]=null, [UDFChar2]=null, [UDFChar3]=null, [UDFChar4]=null, [UDFChar5]=null, [UDFDate1]=null, [UDFDate2]=null, [DemoLaborPK]=null, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionAction]='EDIT', [RowVersionDate]=getdate()
WHERE	ReportGroupPK = 12
PRINT 'Updating ReportGroup Row - ' + 'Asset Reports'

SET @rptGroupPK = 12
END

IF (@@error > 0) SET @errorcount = @errorcount + 1

INSERT INTO Report_ReportGroup ([ReportPK], [ReportGroupPK], [DemoLaborPK], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, @rptGroupPK, '', '', '_MC', getdate())
PRINT 'Inserting Report_ReportGroup Row - ' + 'Asset Reports'
IF (@@error > 0) SET @errorcount = @errorcount + 1

/* Make sure the report group actually exists */
IF NOT EXISTS (SELECT ReportGroupPK FROM ReportGroup WHERE ReportGroupPK = 23)
BEGIN
INSERT INTO ReportGroup ([ReportGroupID], [ReportGroupName], [ModuleID], [Sequence], [Icon], [RepairCenterPK], [IsUserGroup], [IsBatchGroup], [UDFChar1], [UDFChar2], [UDFChar3], [UDFChar4], [UDFChar5], [UDFDate1], [UDFDate2], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionAction], [RowVersionDate]) 
VALUES ('FA', 'Failure Reports', 'FA', 99, null, null, 0, 0, null, null, null, null, null, null, null, null, '', '', '_MC', 'CREATE', getdate())
PRINT 'Inserting ReportGroup Row - ' + 'Failure Reports'

IF (@@error > 0) SET @errorcount = @errorcount + 1

SET @rptGroupPK = @@IDENTITY
END
ELSE
BEGIN
UPDATE	ReportGroup
SET	[ReportGroupID]='FA', [ReportGroupName]='Failure Reports', [ModuleID]='FA', [Sequence]=99, [Icon]=null, [RepairCenterPK]=null, [IsUserGroup]=0, [IsBatchGroup]=0, [UDFChar1]=null, [UDFChar2]=null, [UDFChar3]=null, [UDFChar4]=null, [UDFChar5]=null, [UDFDate1]=null, [UDFDate2]=null, [DemoLaborPK]=null, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionAction]='EDIT', [RowVersionDate]=getdate()
WHERE	ReportGroupPK = 23
PRINT 'Updating ReportGroup Row - ' + 'Failure Reports'

SET @rptGroupPK = 23
END

IF (@@error > 0) SET @errorcount = @errorcount + 1

INSERT INTO Report_ReportGroup ([ReportPK], [ReportGroupPK], [DemoLaborPK], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, @rptGroupPK, '', '', '_MC', getdate())
PRINT 'Inserting Report_ReportGroup Row - ' + 'Failure Reports'
IF (@@error > 0) SET @errorcount = @errorcount + 1


/* ==================================================
DELETE AND INSERT REPORT TABLES - (Sandvik_FtpSuccess)
=================================================== */
DELETE FROM ReportTables WHERE ReportPK = @rptPK

PRINT 'Deleting ReportTables Rows - Sandvik_FtpSuccess'
IF (@@error > 0) SET @errorcount = @errorcount + 1

INSERT INTO ReportTables ([ReportPK], [RFTable], [Alias], [DisplayOrder], [LabelOverride], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, 'MC_InterfaceLog', null, 0, null, null, '', '', '_MC', getdate())
PRINT 'Inserting ReportTables Row - MC_InterfaceLog'
IF (@@error > 0) SET @errorcount = @errorcount + 1


/* ==================================================
UPDATE OR INSERT REPORT CRITERIA - (Sandvik_FtpSuccess)
=================================================== */	
IF EXISTS (
SELECT ReportCriteriaPK FROM ReportCriteria WITH (NOLOCK) WHERE ReportPK = @rptPK AND DisplayTable = 'MC_InterfaceLog' AND DisplayField = 'ProcessDate')
BEGIN
UPDATE ReportCriteria
SET	[SQLWhereTable]='MC_InterfaceLog', [SQLWhereField]='ProcessDate', [CritName]=null, [Operator]='is within', [isMulti]=0, [AskLater]=1, [LabelOverride]=null, [DisplayOrder]=1, [FK_LookupOverride]=null, [DemoLaborPK]=null, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionDate]=getdate()
WHERE	ReportPK = @rptPK AND DisplayTable = 'MC_InterfaceLog' AND DisplayField = 'ProcessDate' 
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Updating ReportCriteria Row - MC_InterfaceLog.ProcessDate'
END
ELSE
BEGIN
INSERT INTO ReportCriteria ([ReportPK], [DisplayTable], [DisplayField], [SQLWhereTable], [SQLWhereField], [DefaultCritValue], [CritName], [Operator], [isMulti], [AskLater], [LabelOverride], [DisplayOrder], [FK_LookupOverride], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, 'MC_InterfaceLog', 'ProcessDate', 'MC_InterfaceLog', 'ProcessDate', 'CW', null, 'is within', 0, 1, null, 1, null, null, null, 0, '_MC', getdate())
PRINT 'Inserting ReportCriteria Row - MC_InterfaceLog.ProcessDate'
IF (@@error > 0) SET @errorcount = @errorcount + 1

END

IF EXISTS (
SELECT ReportCriteriaPK FROM ReportCriteria WITH (NOLOCK) WHERE ReportPK = @rptPK AND DisplayTable = 'MC_InterfaceLog' AND DisplayField = 'Processed')
BEGIN
UPDATE ReportCriteria
SET	[SQLWhereTable]='MC_InterfaceLog', [SQLWhereField]='Processed', [CritName]=null, [Operator]='is', [isMulti]=0, [AskLater]=1, [LabelOverride]=null, [DisplayOrder]=0, [FK_LookupOverride]=null, [DemoLaborPK]=null, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionDate]=getdate()
WHERE	ReportPK = @rptPK AND DisplayTable = 'MC_InterfaceLog' AND DisplayField = 'Processed' 
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Updating ReportCriteria Row - MC_InterfaceLog.Processed'
END
ELSE
BEGIN
INSERT INTO ReportCriteria ([ReportPK], [DisplayTable], [DisplayField], [SQLWhereTable], [SQLWhereField], [DefaultCritValue], [CritName], [Operator], [isMulti], [AskLater], [LabelOverride], [DisplayOrder], [FK_LookupOverride], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, 'MC_InterfaceLog', 'Processed', 'MC_InterfaceLog', 'Processed', 'Y', null, 'is', 0, 1, null, 0, null, null, null, 0, '_MC', getdate())
PRINT 'Inserting ReportCriteria Row - MC_InterfaceLog.Processed'
IF (@@error > 0) SET @errorcount = @errorcount + 1

END

/* ==================================================
UPDATE OR INSERT REPORT FIELDS - (Sandvik_FtpSuccess)
=================================================== */	
IF EXISTS (
SELECT ReportFieldPK FROM ReportFields WITH (NOLOCK) WHERE ReportPK = @rptPK AND RFTable = 'MC_InterfaceLog' AND RFField = 'FileName')
BEGIN
SELECT @tempPK = ReportFieldPK FROM ReportFields WITH (NOLOCK) WHERE ReportPK = @rptPK AND RFTable = 'MC_InterfaceLog' AND RFField = 'FileName'
UPDATE ReportFields
SET [ReportPK]=@rptPK, [DataDictPK]=9786, [AGFunction]=null, [Alias]=null, [DisplayOrder]=0, [Display]=1, [NotUserSelectable]=0, [LabelOverride]='File Name', [TotalIfSelected]=0, [BlankLineIfSelected]=0, [UseCustomExpression]=0, [CustomExpression]=null, [DemoLaborPK]=0, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionDate]=getdate(), [SLAction]='  ', [SLModuleID]=null, [SLPKField]=null, [SLReportID]=null, [SLCustomAction]=null, [SLToolTip]=null 
WHERE [ReportPK]=@rptPK AND [RFTable]='MC_InterfaceLog' AND [RFField]='FileName'
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Updating ReportFields Row - MC_InterfaceLog.FileName'
END
ELSE
BEGIN
INSERT INTO ReportFields ([ReportPK], [DataDictPK], [AGFunction], [RFTable], [RFField], [Alias], [DisplayOrder], [Display], [NotUserSelectable], [LabelOverride], [TotalIfSelected], [BlankLineIfSelected], [UseCustomExpression], [CustomExpression], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate], [SLAction], [SLModuleID], [SLPKField], [SLReportID], [SLCustomAction], [SLTooltip])
VALUES(@rptPK, 9786, null, 'MC_InterfaceLog', 'FileName', null, 0, 1, 0, 'File Name', 0, 0, 0, null, 0, '', '', '_MC', getdate(), '  ', null, null, null, null, null)
PRINT 'Inserting ReportFields Row - MC_InterfaceLog.FileName'
IF (@@error > 0) 
BEGIN
SET @errorcount = @errorcount + 1
SET @tempPK = -1
END
ELSE
BEGIN
SET @tempPK = @@IDENTITY
END
END

IF @version >= 4.2 
BEGIN
SET @dSQL = 'UPDATE ReportFields SET [Alignment] = null, [AdditionalWidth] = null, [PivotSetup] = ''  '', [AddPivotColumnsWithNoDataFrom] = null, [AddPivotColumnsWithNoDataFromCustom] = null, [Data_Type_Override] = null, [ColumnFormat] = null, [ColumnCS] = null, [GroupByCustomExpression] = 1 WHERE ReportFieldPK = ' + CAST(@tempPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1
END
IF EXISTS (
SELECT ReportFieldPK FROM ReportFields WITH (NOLOCK) WHERE ReportPK = @rptPK AND RFTable = 'MC_InterfaceLog' AND RFField = 'ProcessDate')
BEGIN
SELECT @tempPK = ReportFieldPK FROM ReportFields WITH (NOLOCK) WHERE ReportPK = @rptPK AND RFTable = 'MC_InterfaceLog' AND RFField = 'ProcessDate'
UPDATE ReportFields
SET [ReportPK]=@rptPK, [DataDictPK]=9788, [AGFunction]=null, [Alias]=null, [DisplayOrder]=1, [Display]=1, [NotUserSelectable]=0, [LabelOverride]=null, [TotalIfSelected]=0, [BlankLineIfSelected]=0, [UseCustomExpression]=0, [CustomExpression]=null, [DemoLaborPK]=0, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionDate]=getdate(), [SLAction]='  ', [SLModuleID]=null, [SLPKField]=null, [SLReportID]=null, [SLCustomAction]=null, [SLToolTip]=null 
WHERE [ReportPK]=@rptPK AND [RFTable]='MC_InterfaceLog' AND [RFField]='ProcessDate'
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Updating ReportFields Row - MC_InterfaceLog.ProcessDate'
END
ELSE
BEGIN
INSERT INTO ReportFields ([ReportPK], [DataDictPK], [AGFunction], [RFTable], [RFField], [Alias], [DisplayOrder], [Display], [NotUserSelectable], [LabelOverride], [TotalIfSelected], [BlankLineIfSelected], [UseCustomExpression], [CustomExpression], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate], [SLAction], [SLModuleID], [SLPKField], [SLReportID], [SLCustomAction], [SLTooltip])
VALUES(@rptPK, 9788, null, 'MC_InterfaceLog', 'ProcessDate', null, 1, 1, 0, null, 0, 0, 0, null, 0, '', '', '_MC', getdate(), '  ', null, null, null, null, null)
PRINT 'Inserting ReportFields Row - MC_InterfaceLog.ProcessDate'
IF (@@error > 0) 
BEGIN
SET @errorcount = @errorcount + 1
SET @tempPK = -1
END
ELSE
BEGIN
SET @tempPK = @@IDENTITY
END
END

IF @version >= 4.2 
BEGIN
SET @dSQL = 'UPDATE ReportFields SET [Alignment] = null, [AdditionalWidth] = null, [PivotSetup] = ''  '', [AddPivotColumnsWithNoDataFrom] = null, [AddPivotColumnsWithNoDataFromCustom] = null, [Data_Type_Override] = null, [ColumnFormat] = null, [ColumnCS] = null, [GroupByCustomExpression] = 1 WHERE ReportFieldPK = ' + CAST(@tempPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1
END
IF EXISTS (
SELECT ReportFieldPK FROM ReportFields WITH (NOLOCK) WHERE ReportPK = @rptPK AND RFTable = 'MC_InterfaceLog' AND RFField = 'ErrorMessage')
BEGIN
SELECT @tempPK = ReportFieldPK FROM ReportFields WITH (NOLOCK) WHERE ReportPK = @rptPK AND RFTable = 'MC_InterfaceLog' AND RFField = 'ErrorMessage'
UPDATE ReportFields
SET [ReportPK]=@rptPK, [DataDictPK]=9785, [AGFunction]=null, [Alias]=null, [DisplayOrder]=2, [Display]=1, [NotUserSelectable]=0, [LabelOverride]=null, [TotalIfSelected]=0, [BlankLineIfSelected]=0, [UseCustomExpression]=0, [CustomExpression]=null, [DemoLaborPK]=null, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionDate]=getdate(), [SLAction]=null, [SLModuleID]=null, [SLPKField]=null, [SLReportID]=null, [SLCustomAction]=null, [SLToolTip]=null 
WHERE [ReportPK]=@rptPK AND [RFTable]='MC_InterfaceLog' AND [RFField]='ErrorMessage'
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Updating ReportFields Row - MC_InterfaceLog.ErrorMessage'
END
ELSE
BEGIN
INSERT INTO ReportFields ([ReportPK], [DataDictPK], [AGFunction], [RFTable], [RFField], [Alias], [DisplayOrder], [Display], [NotUserSelectable], [LabelOverride], [TotalIfSelected], [BlankLineIfSelected], [UseCustomExpression], [CustomExpression], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate], [SLAction], [SLModuleID], [SLPKField], [SLReportID], [SLCustomAction], [SLTooltip])
VALUES(@rptPK, 9785, null, 'MC_InterfaceLog', 'ErrorMessage', null, 2, 1, 0, null, 0, 0, 0, null, null, '', '', '_MC', getdate(), null, null, null, null, null, null)
PRINT 'Inserting ReportFields Row - MC_InterfaceLog.ErrorMessage'
IF (@@error > 0) 
BEGIN
SET @errorcount = @errorcount + 1
SET @tempPK = -1
END
ELSE
BEGIN
SET @tempPK = @@IDENTITY
END
END

IF @version >= 4.2 
BEGIN
SET @dSQL = 'UPDATE ReportFields SET [Alignment] = null, [AdditionalWidth] = null, [PivotSetup] = null, [AddPivotColumnsWithNoDataFrom] = null, [AddPivotColumnsWithNoDataFromCustom] = null, [Data_Type_Override] = null, [ColumnFormat] = null, [ColumnCS] = null, [GroupByCustomExpression] = 1 WHERE ReportFieldPK = ' + CAST(@tempPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1
END
/* =====================================================
UPDATE REPORTS COLS IN DATA DICT - (Sandvik_FtpSuccess)
===================================================== */
UPDATE DataDict 
SET	[REPORT_NOSELECT]=null, [REPORT_EDIT]=null, [REPORT_LABEL]=null, [TOTALIFSELECTED]=0 
WHERE	TABLE_NAME = 'MC_InterfaceLog' AND COLUMN_NAME = 'ErrorMessage' 

PRINT 'Updating DataDict Row - MC_InterfaceLog.ErrorMessage'
IF (@@error > 0) SET @errorcount = @errorcount + 1

UPDATE DataDict 
SET	[REPORT_NOSELECT]=null, [REPORT_EDIT]=null, [REPORT_LABEL]=null, [TOTALIFSELECTED]=0 
WHERE	TABLE_NAME = 'MC_InterfaceLog' AND COLUMN_NAME = 'ProcessDate' 

PRINT 'Updating DataDict Row - MC_InterfaceLog.ProcessDate'
IF (@@error > 0) SET @errorcount = @errorcount + 1

END
/* REPORT DID NOT EXIST, CREATE ALL RECORDS */
ELSE

BEGIN
PRINT '*******************************************************'
PRINT 'Report Does Not Exist - Inserting...'
PRINT 'Report: FTP Import Success'
PRINT '*******************************************************'

/* =====================================================
INSERT REPORT RECORDS - (Sandvik_FtpSuccess)
===================================================== */
/* INSERT Main Report */
INSERT INTO Reports ([ReportIDPriorToCopy], [ReportID], [ReportName], [ReportDesc], [RepairCenterPK], [Sort1], [Sort2], [Sort3], [Sort4], [Sort5], [Sort1DESC], [Sort2DESC], [Sort3DESC], [Sort4DESC], [Sort5DESC], [Group1], [Group2], [Group3], [Group4], [Group5], [Header1], [Header2], [Header3], [Header4], [Header5], [GroupHeader1], [GroupHeader2], [GroupHeader3], [GroupHeader4], [GroupHeader5], [Total1], [Total2], [Total3], [Total4], [Total5], [Chart], [ChartName], [ChartField], [ChartSize], [ReportFile], [FromSQL], [JoinSQL], [WhereSQL], [GroupBy], [hits], [Sequence], [Layout], [VertCols], [PageBreakEachRecord], [Custom], [ReportCopy], [MCRegistrationDB], [PrintCriteria], [Active], [UDFChar1], [UDFChar2], [UDFChar3], [UDFChar4], [UDFChar5], [UDFDate1], [UDFDate2], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionAction], [RowVersionDate], [ChartFunction], [ChartFunctionField], [NoDetail], [PB1], [PB2], [PB3], [PB4], [PB5], [SLDefault], [SLType], [SLAction], [SLModuleID], [SLPKField], [SLReportID], [SLCustomAction], [SLTooltip], [SDDisplay], [SDModuleID], [SDPKField], [SmartEmail], [ChartPosition], [ChartFormat], [ChartSQL], [Chart2], [ChartName2], [ChartField2], [ChartSize2], [ChartFormat2], [ChartFunction2], [ChartFunctionField2], [ChartPosition2], [ChartSQL2], [Chart3], [ChartName3], [ChartField3], [ChartSize3], [ChartFormat3], [ChartFunction3], [ChartFunctionField3], [ChartPosition3], [ChartSQL3], [ChartOnly], [NoHeader], [HavingSQL], [SRID1], [SRPKField1], [SRID2], [SRPKField2], [SRID3], [SRPKField3], [SRID4], [SRPKField4], [SRID5], [SRPKField5], [ReportPageSize], [ReportWidth], [PhotoCriteria], [ReportStyleName], [UsedFor], [SmartEmailLaborPK], [SCDefault], [SCField1], [SCField2], [SCField3], [ReportStyleFontSize], [ReportStyleFontColor], [ReportStyleFontFamily] )
VALUES('NavGeoTabImpErrors', 'Sandvik_FtpSuccess', 'FTP Import Success', null, null, 'MC_InterfaceLog.PK', null, null, null, null, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, null, null, null, null, null, 0, 0, 0, 0, 0, null, 'Work Order Status', 'AUTO_SORT1', 'L', 'rpt_generic1.asp', 'FROM MC_InterfaceLog', null, null, 0, 0, 0, 'hor', 1, 0, 0, 1, 0, 0, 1, null, null, null, 'N', null, null, null, null, '', '', '_MC', 'CREATE', getdate(), 'C', 'NONE', 0, 0, 0, 0, 0, 0, 1, ' ', 'PW', 'WO', null, null, null, null, '     ', '  ', null, 0, 'T', 'F', null, null, null, null, null, 'I', null, null, null, null, null, null, null, null, 'I', null, null, null, null, 0, 0, null, null, null, null, null, null, null, null, null, null, null, 'Default', '80%', 1, 'Gradient - Blue', 'REPORTS', 0, 'H', null, null, null, null, null, null)

PRINT 'Inserting Report - Sandvik_FtpSuccess'
IF (@@error > 0) SET @errorcount = @errorcount + 1

SET @rptPK = @@Identity

IF @version >= 3.0 
BEGIN
SET @dSQL = 'UPDATE Reports SET [HavingSQL]=null, [DisplayPivotBar]=0, [DisplayColumnLines]=0, [DisplayTitleonPageBreak]=0, [DisplayFormatCriteria]=1, [R1T]='' '', [R1O]=null, [R1V1]=null, [R1V2]=null, [R1A]=0, [R1L]=0, [R1F]=0, [R1CS]=''font-family: Arial;font-size: 8pt;color: #000000;text-align: left;'', [R1AF]=''C'', [R2T]='' '', [R2O]=null, [R2V1]=null, [R2V2]=null, [R2A]=0, [R2L]=0, [R2F]=0, [R2CS]=''border: #0066CC 2px solid;'', [R2AF]=''C'', [R3T]='' '', [R3O]=null, [R3V1]=null, [R3V2]=null, [R3A]=0, [R3L]=0, [R3F]=0, [R3CS]=''border: #0066CC 2px solid;'', [R3AF]=''C'' WHERE ReportPK = ' + CAST(@rptPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1

END

IF @version >= 4.2 
BEGIN
SET @dSQL = 'UPDATE Reports SET [DisplayDescription]=0 WHERE ReportPK = ' + CAST(@rptPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1

END

/* ==================================================
UPDATE OR INSERT REPORT GROUPS - (Sandvik_FtpSuccess)
=================================================== */

/* Make sure the report group actually exists */
IF NOT EXISTS (SELECT ReportGroupPK FROM ReportGroup WHERE ReportGroupPK = 12)
BEGIN
INSERT INTO ReportGroup ([ReportGroupID], [ReportGroupName], [ModuleID], [Sequence], [Icon], [RepairCenterPK], [IsUserGroup], [IsBatchGroup], [UDFChar1], [UDFChar2], [UDFChar3], [UDFChar4], [UDFChar5], [UDFDate1], [UDFDate2], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionAction], [RowVersionDate]) 
VALUES ('AS', 'Asset Reports', 'AS', 99, null, null, 0, 0, null, null, null, null, null, null, null, null, '', '', '_MC', 'CREATE', getdate())
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Inserting ReportGroup Row - ' + 'Asset Reports'

SET @rptGroupPK = @@IDENTITY
END
ELSE
BEGIN
UPDATE	ReportGroup
SET	[ReportGroupID]='AS', [ReportGroupName]='Asset Reports', [ModuleID]='AS', [Sequence]=99, [Icon]=null, [RepairCenterPK]=null, [IsUserGroup]=0, [IsBatchGroup]=0, [UDFChar1]=null, [UDFChar2]=null, [UDFChar3]=null, [UDFChar4]=null, [UDFChar5]=null, [UDFDate1]=null, [UDFDate2]=null, [DemoLaborPK]=null, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionAction]='EDIT', [RowVersionDate]=getdate()
WHERE	ReportGroupPK = 12
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Updating ReportGroup Row - ' + 'Asset Reports'

SET @rptGroupPK = 12
END

INSERT INTO Report_ReportGroup ([ReportPK], [ReportGroupPK], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, @rptGroupPK, null, '', '', '_MC', getdate())
PRINT 'Inserting Report_ReportGroup Row - ' + 'Asset Reports'
IF (@@error > 0) SET @errorcount = @errorcount + 1

/* Make sure the report group actually exists */
IF NOT EXISTS (SELECT ReportGroupPK FROM ReportGroup WHERE ReportGroupPK = 23)
BEGIN
INSERT INTO ReportGroup ([ReportGroupID], [ReportGroupName], [ModuleID], [Sequence], [Icon], [RepairCenterPK], [IsUserGroup], [IsBatchGroup], [UDFChar1], [UDFChar2], [UDFChar3], [UDFChar4], [UDFChar5], [UDFDate1], [UDFDate2], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionAction], [RowVersionDate]) 
VALUES ('FA', 'Failure Reports', 'FA', 99, null, null, 0, 0, null, null, null, null, null, null, null, null, '', '', '_MC', 'CREATE', getdate())
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Inserting ReportGroup Row - ' + 'Failure Reports'

SET @rptGroupPK = @@IDENTITY
END
ELSE
BEGIN
UPDATE	ReportGroup
SET	[ReportGroupID]='FA', [ReportGroupName]='Failure Reports', [ModuleID]='FA', [Sequence]=99, [Icon]=null, [RepairCenterPK]=null, [IsUserGroup]=0, [IsBatchGroup]=0, [UDFChar1]=null, [UDFChar2]=null, [UDFChar3]=null, [UDFChar4]=null, [UDFChar5]=null, [UDFDate1]=null, [UDFDate2]=null, [DemoLaborPK]=null, [RowVersionIPAddress]='', [RowVersionUserPK]='', [RowVersionInitials]='_MC', [RowVersionAction]='EDIT', [RowVersionDate]=getdate()
WHERE	ReportGroupPK = 23
IF (@@error > 0) SET @errorcount = @errorcount + 1

PRINT 'Updating ReportGroup Row - ' + 'Failure Reports'

SET @rptGroupPK = 23
END

INSERT INTO Report_ReportGroup ([ReportPK], [ReportGroupPK], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, @rptGroupPK, null, '', '', '_MC', getdate())
PRINT 'Inserting Report_ReportGroup Row - ' + 'Failure Reports'
IF (@@error > 0) SET @errorcount = @errorcount + 1

/* ==================================================
INSERT REPORT TABLES- (Sandvik_FtpSuccess)
=================================================== */	
INSERT INTO ReportTables ([ReportPK], [RFTable], [Alias], [DisplayOrder], [LabelOverride], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, 'MC_InterfaceLog', null, 0, null, null, '', '', '_MC', getdate())
PRINT 'Inserting ReportTables Row - MC_InterfaceLog'
IF (@@error > 0) SET @errorcount = @errorcount + 1


/* ==================================================
INSERT REPORT CRITERIA - (Sandvik_FtpSuccess)
=================================================== */	
INSERT INTO ReportCriteria ([ReportPK], [DisplayTable], [DisplayField], [SQLWhereTable], [SQLWhereField], [DefaultCritValue], [CritName], [Operator], [isMulti], [AskLater], [LabelOverride], [DisplayOrder], [FK_LookupOverride], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, 'MC_InterfaceLog', 'ProcessDate', 'MC_InterfaceLog', 'ProcessDate', 'CW', null, 'is within', 0, 1, null, 1, null, null, '', '', '_MC', getdate())
PRINT 'Inserting ReportCriteria Row - MC_InterfaceLog.ProcessDate'
IF (@@error > 0) SET @errorcount = @errorcount + 1

INSERT INTO ReportCriteria ([ReportPK], [DisplayTable], [DisplayField], [SQLWhereTable], [SQLWhereField], [DefaultCritValue], [CritName], [Operator], [isMulti], [AskLater], [LabelOverride], [DisplayOrder], [FK_LookupOverride], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate]) VALUES (@rptPK, 'MC_InterfaceLog', 'Processed', 'MC_InterfaceLog', 'Processed', 'Y', null, 'is', 0, 1, null, 0, null, null, '', '', '_MC', getdate())
PRINT 'Inserting ReportCriteria Row - MC_InterfaceLog.Processed'
IF (@@error > 0) SET @errorcount = @errorcount + 1

/* ==================================================
INSERT REPORT FIELDS - (Sandvik_FtpSuccess)
=================================================== */	
INSERT INTO ReportFields ([ReportPK], [DataDictPK], [AGFunction], [RFTable], [RFField], [Alias], [DisplayOrder], [Display], [NotUserSelectable], [LabelOverride], [TotalIfSelected], [BlankLineIfSelected], [UseCustomExpression], [CustomExpression], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate], [SLAction], [SLModuleID], [SLPKField], [SLReportID], [SLCustomAction], [SLTooltip])
VALUES(@rptPK, 9786, null, 'MC_InterfaceLog', 'FileName', null, 0, 1, 0, 'File Name', 0, 0, 0, null, 0, '', '', '_MC', getdate(), '  ', null, null, null, null, null)
PRINT 'Inserting ReportFields Row - MC_InterfaceLog.FileName'
IF (@@error > 0) SET @errorcount = @errorcount + 1

IF @version >= 4.2 
BEGIN
SET @tempPK = @@IDENTITY
SET @dSQL = 'UPDATE ReportFields SET [Alignment] = null, [AdditionalWidth] = null, [PivotSetup] = ''  '', [AddPivotColumnsWithNoDataFrom] = null, [AddPivotColumnsWithNoDataFromCustom] = null, [Data_Type_Override] = null, [ColumnFormat] = null, [ColumnCS] = null, [GroupByCustomExpression] = 1 WHERE ReportFieldPK = ' + CAST(@tempPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1
END
INSERT INTO ReportFields ([ReportPK], [DataDictPK], [AGFunction], [RFTable], [RFField], [Alias], [DisplayOrder], [Display], [NotUserSelectable], [LabelOverride], [TotalIfSelected], [BlankLineIfSelected], [UseCustomExpression], [CustomExpression], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate], [SLAction], [SLModuleID], [SLPKField], [SLReportID], [SLCustomAction], [SLTooltip])
VALUES(@rptPK, 9788, null, 'MC_InterfaceLog', 'ProcessDate', null, 1, 1, 0, null, 0, 0, 0, null, 0, '', '', '_MC', getdate(), '  ', null, null, null, null, null)
PRINT 'Inserting ReportFields Row - MC_InterfaceLog.ProcessDate'
IF (@@error > 0) SET @errorcount = @errorcount + 1

IF @version >= 4.2 
BEGIN
SET @tempPK = @@IDENTITY
SET @dSQL = 'UPDATE ReportFields SET [Alignment] = null, [AdditionalWidth] = null, [PivotSetup] = ''  '', [AddPivotColumnsWithNoDataFrom] = null, [AddPivotColumnsWithNoDataFromCustom] = null, [Data_Type_Override] = null, [ColumnFormat] = null, [ColumnCS] = null, [GroupByCustomExpression] = 1 WHERE ReportFieldPK = ' + CAST(@tempPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1
END
INSERT INTO ReportFields ([ReportPK], [DataDictPK], [AGFunction], [RFTable], [RFField], [Alias], [DisplayOrder], [Display], [NotUserSelectable], [LabelOverride], [TotalIfSelected], [BlankLineIfSelected], [UseCustomExpression], [CustomExpression], [DemoLaborPK], [RowVersionIPAddress], [RowVersionUserPK], [RowVersionInitials], [RowVersionDate], [SLAction], [SLModuleID], [SLPKField], [SLReportID], [SLCustomAction], [SLTooltip])
VALUES(@rptPK, 9785, null, 'MC_InterfaceLog', 'ErrorMessage', null, 2, 1, 0, null, 0, 0, 0, null, null, '', '', '_MC', getdate(), null, null, null, null, null, null)
PRINT 'Inserting ReportFields Row - MC_InterfaceLog.ErrorMessage'
IF (@@error > 0) SET @errorcount = @errorcount + 1

IF @version >= 4.2 
BEGIN
SET @tempPK = @@IDENTITY
SET @dSQL = 'UPDATE ReportFields SET [Alignment] = null, [AdditionalWidth] = null, [PivotSetup] = null, [AddPivotColumnsWithNoDataFrom] = null, [AddPivotColumnsWithNoDataFromCustom] = null, [Data_Type_Override] = null, [ColumnFormat] = null, [ColumnCS] = null, [GroupByCustomExpression] = 1 WHERE ReportFieldPK = ' + CAST(@tempPK AS varchar(20)) + '
'
EXEC(@dSQL)
IF (@@error > 0) SET @errorcount = @errorcount + 1
END
/* =====================================================
UPDATE REPORTS COLS IN DATA DICT - (Sandvik_FtpSuccess)
===================================================== */

UPDATE DataDict 
SET	[REPORT_NOSELECT]=null, [REPORT_EDIT]=null, [REPORT_LABEL]=null, [TOTALIFSELECTED]=0 
WHERE	TABLE_NAME = 'MC_InterfaceLog' AND COLUMN_NAME = 'ErrorMessage' 

PRINT 'Updating DataDict Row - MC_InterfaceLog.ErrorMessage'
if (@@error > 0) Set @errorcount = @errorcount + 1

UPDATE DataDict 
SET	[REPORT_NOSELECT]=null, [REPORT_EDIT]=null, [REPORT_LABEL]=null, [TOTALIFSELECTED]=0 
WHERE	TABLE_NAME = 'MC_InterfaceLog' AND COLUMN_NAME = 'ProcessDate' 

PRINT 'Updating DataDict Row - MC_InterfaceLog.ProcessDate'
if (@@error > 0) Set @errorcount = @errorcount + 1

END

PRINT '*******************************************************'
IF (@errorcount > 0) PRINT @errorcount + ' Error(s) Occurred: Sandvik_FtpSuccess'
ELSE PRINT 'No Errors Occurred: Sandvik_FtpSuccess'

GO

