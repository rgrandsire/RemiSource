/*****************************************************************************************************************
#################################################################################################################
#     Date     #   version   #      Author      #                   Comments 
#---------------------------------------------------------------------------------------------------------------- 
#  12/08/2016  #   1.0.0.1   # Remi G Grandsire # Original version 
#################################################################################################################
*****************************************************************************************************************/

-- Make sure the correct database is selected prior to running this script

-- Adding Columns to Asset table

ALTER TABLE Asset
ADD UDFChar11 VARCHAR(75) NULL
   ,UDFChar12 VARCHAR(75) NULL
   ,UDFChar13 VARCHAR(75) NULL
   ,UDFChar14 VARCHAR(75) NULL
   ,UDFChar15 VARCHAR(75) NULL
   ,UDFChar16 VARCHAR(75) NULL
   ,UDFChar17 VARCHAR(75) NULL
   ,UDFChar18 VARCHAR(75) NULL
   ,UDFChar19 VARCHAR(75) NULL
   ,UDFChar20 VARCHAR(75) NULL
   ,UDFChar21 VARCHAR(75) NULL
   ,UDFChar22 VARCHAR(75) NULL
   ,UDFChar23 VARCHAR(75) NULL
   ,UDFChar24 VARCHAR(75) NULL
   ,UDFChar25 VARCHAR(75) NULL
   ,UDFChar26 VARCHAR(75) NULL
   ,UDFChar27 VARCHAR(75) NULL
   ,UDFChar28 VARCHAR(75) NULL
   ,UDFChar29 VARCHAR(75) NULL
   ,UDFChar30 VARCHAR(75) NULL
   ,UDFChar31 VARCHAR(75) NULL
   ,UDFChar32 VARCHAR(75) NULL
   ,UDFChar33 VARCHAR(75) NULL
   ,UDFChar34 VARCHAR(75) NULL
   ,UDFChar35 VARCHAR(75) NULL
   ,UDFChar36 VARCHAR(75) NULL
   ,UDFChar37 VARCHAR(75) NULL
   ,UDFChar38 VARCHAR(75) NULL
   ,UDFChar39 VARCHAR(75) NULL
   ,UDFChar40 VARCHAR(75) NULL
   ,UDFChar41 VARCHAR(75) NULL
   ,UDFChar42 VARCHAR(75) NULL
   ,UDFChar43 VARCHAR(75) NULL
   ,UDFChar44 VARCHAR(75) NULL
   ,UDFChar45 VARCHAR(75) NULL
   ,UDFChar46 VARCHAR(75) NULL
   ,UDFChar47 VARCHAR(75) NULL
   ,UDFChar48 VARCHAR(75) NULL
   ,UDFChar49 VARCHAR(75) NULL
   ,UDFChar50 VARCHAR(75) NULL
   ,UDFChar51 VARCHAR(75) NULL
   ,UDFChar52 VARCHAR(75) NULL;
-- Script 1 complete
   
/*****************************************************************************************************************
#################################################################################################################
#     Date     #   version   #      Author      #                   Comments 
#---------------------------------------------------------------------------------------------------------------- 
#  12/08/2016  #   1.0.0.1   # Remi G Grandsire # Original version 
#################################################################################################################
*****************************************************************************************************************/

-- Make sure the correct database is selected prior to running this script

-- This script inserts the custom caption for the new AEM Risk Tab


INSERT INTO specification
  ([SpecificationName]
  ,[Description]
  ,[TrackHistory]
  ,[Active]
  ,[Comments]
  ,[UseLookupTable]
  ,[LookupTable]
  )
VALUES
  ('Consequences'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Likely consequences of equipment failure or malfunction including seriousness of and prevalence of harm'
  ,'True'
  ,'Consequences'
  ),
  ('Degraded'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Does the AEM result in degraded performance of equipment?'
  ,'False'
  ,'NULL'
  ),
  ('Director'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Director or Assistant Director Decision'
  ,'False'
  ,'NULL'
  ),
  ('Effectiveness'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Effectiveness of AEM Program methods this equipment is based on. Check all that apply'
  ,'False'
  ,'NULL'
  ),
  ('Equipment'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'How is the equipment/ component used, Check one'
  ,'False'
  ,'NULL'
  ),
  ('Evaluation'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Equipment / component evaluation to determine if AEM program needs to be altered.'
  ,'False'
  ,'NULL'
  ),
  ('Failure'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Previous service and failure request?'
  ,'False'
  ,'NULL'
  ),
  ('History'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Maintenance History from MC History tab'
  ,'False'
  ,'NULL'
  ),
  ('Hospital'
  ,'RIsk_Tab'
  ,'False'
  ,'True'
  ,'How does hospital assess if AEM Program uses appropriate maintenance strategies?'
  ,'False'
  ,'NULL'
  ),
  ('Included'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Included in the AEM program'
  ,'False'
  ,'NULL'
  ),
  ('Maintenance'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Are maintenance requirements simple or complex'
  ,'False'
  ,'NULL'
  ),
  ('Malfunction'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Could the malfunction have been prevented?'
  ,'False'
  ,'NULL'
  ),
  ('Modified'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Why are manufacture requirements being modified? Check one'
  ,'False'
  ,'NULL'
  ),
  ('Remove'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Remove equipment /component from service when no longer safe or suitable for intended service if warranted.'
  ,'False'
  ,'NULL'
  ),
  ('Requirements'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'How much of manufactures requirements are being changed? Check one'
  ,'False'
  ,'NULL'
  ),
  ('Review'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Reviewed by:'
  ,'True'
  ,'Labor'
  ),
  ('Score'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Score of 8 is Classed as Critical/ High Risk'
  ,'False'
  ,'NULL'
  ),
  ('Seriousness'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Seriousness and prevalence of harm during normal use:'
  ,'True'
  ,'RemiRisk1'
  ),
  ('Timeliness'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Timeliness of alternate or Back-up equipment in the event of a failure. Check one'
  ,'False'
  ,'NULL'
  ),
  ('Why'
  ,'Risk_Tab'
  ,'False'
  ,'True'
  ,'Why or How?'
  ,'False'
  ,'NULL'
  )
   
-- Script 2 complete


/*****************************************************************************************************************
#################################################################################################################
#     Date     #   version   #      Author      #                   Comments 
#---------------------------------------------------------------------------------------------------------------- 
#  12/08/2016  #   1.0.0.1   # Remi G Grandsire # Original version 
#################################################################################################################
*****************************************************************************************************************/

-- Make sure the correct database is selected prior to running this script

-- This script adds the Looku[p tables for the cpations

INSERT INTO LookupTable
  ([LookupTable]
  ,[Description]
  ,[CodeWidth]
  ,[DescriptionWidth]
  ,[Internal]
  ,[Enabled]
  ,[CanModify]
  ,[SkipValidation]
  ,[NoCascadeUpdate]
  ,[RowVersionDate]
  ,[NotViewableInLookupTableMgr]
  )
VALUES
  ('R1_RISK1'
  ,'R1_RISK1 TABLE'
  ,25
  ,50
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R10_RISK'
  ,'R10_RISK TABLE'
  ,25
  ,100
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R11_RISK'
  ,'R11_RISK TABLE'
  ,25
  ,100
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R12_RISK'
  ,'R12_RISK TABLE'
  ,25
  ,100
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R13_RISK'
  ,'R13_RISK TABLE'
  ,25
  ,100
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R14_RISK'
  ,'R14_RISK TABLE'
  ,25
  ,100
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R15_RISK'
  ,'R15_RISK TABLE'
  ,25
  ,100
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R16_RISK'
  ,'R16_RISK TABLE'
  ,25
  ,100
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R17_RISK'
  ,'R17_RISK TABLE'
  ,25
  ,100
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R19_RISK'
  ,'R19_RISK TABLE'
  ,25
  ,100
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R2_RISK'
  ,'R2_RISK TABLE'
  ,25
  ,100
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R21_RISK'
  ,'R21_RISK TABLE'
  ,25
  ,100
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R22_RISK'
  ,'R22_RISK TABLE'
  ,25
  ,50
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R3_RISK'
  ,'R3_RISK TABLE'
  ,25
  ,100
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R6_RISK'
  ,'R6_RISK TABLE'
  ,25
  ,100
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R7_RISK'
  ,'R7_RISK TABLE'
  ,25
  ,100
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('R9_RISK'
  ,'R9_RISK TABLE'
  ,25
  ,100
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('REMIRISK1'
  ,'REMIRISK1 TABLE'
  ,25
  ,50
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  ),
  ('RISK1'
  ,'RISK1 TABLE'
  ,25
  ,50
  ,'False'
  ,'True'
  ,'True'
  ,'False'
  ,'False'
  ,GETDATE()
  ,'False'
  )

-- Script 3 complete

/*****************************************************************************************************************
#################################################################################################################
#     Date     #   version   #      Author      #                   Comments 
#---------------------------------------------------------------------------------------------------------------- 
#  12/08/2016  #   1.0.0.1   # Remi G Grandsire # Original version 
#################################################################################################################
*****************************************************************************************************************/

-- Make sure the correct database is selected prior to running this script
-- This script will the ASP pages in the database to run the custom Asset Risk page(s)


UPDATE MCModule
SET     
   FormPage = '_asset_healthcare_Tanner.asp'
  ,ActionPage = '_asset_healthcare_Tanner.asp'
WHERE  ModuleID = 'AS'
   
   
-- script 4 complete   
   
   
   
-- Update Data Dic:
EXEC ADM_UpdateDataDict;

-- Complete