/*
###############################################################################################################################
#   Date     #     Author       #                         Comments
# 10/20/2016 # Remi G Grandsire # Original script to set the reports to the new ASP page
###############################################################################################################################

Make sure to set the correct database prior to running the script...
*/

UPDATE Reports SET ReportFile = 'rpt_WO1_nogroups_SacCity_Rev2.asp' WHERE ReportID = 'WorkOrder';
UPDATE Reports SET ReportFile = 'rpt_WO1_nogroups_SacCity_Rev2.asp' WHERE ReportID = 'WorkOrderC';
UPDATE Reports SET ReportFile = 'rpt_WO1_nogroups_SacCity_Rev2.asp' WHERE ReportID = 'WorkOrderNoGroups';
UPDATE Reports SET ReportFile = 'rpt_WO1_nogroups_SacCity_Rev2.asp' WHERE ReportID = 'WorkOrderCNoGroups';

