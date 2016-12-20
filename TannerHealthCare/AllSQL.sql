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
   ,UDFChar52 VARCHAR(75) NULL

-- Complete

/*****************************************************************************************************************
#################################################################################################################
#     Date     #   version   #      Author      #                   Comments 
#---------------------------------------------------------------------------------------------------------------- 
#  12/08/2016  #   1.0.0.1   # Remi G Grandsire # Original version 
#################################################################################################################
*****************************************************************************************************************/

-- Make sure the correct database is selected prior to running this script

-- This script inserts the custom caption for the new AEM Risk Tab


INSERT INTO Specification
  ([SpecificationName]
  ,[Description]
  ,[Comments]
  ,[TrackHistory]
  ,[Active]
  )
VALUES
  ('Consequences'
  ,'Risk_Tab'
  ,'Likely consequences of equipment failure or malfunction including seriousness of and prevalence of harm'
  ,'False'
  ,'True'
  ),
  ('Degraded'
  ,'Risk_Tab'
  ,'Does the AEM result in degraded performance of equipment?'
  ,'False'
  ,'True'
  ),
  ('Effectiveness'
  ,'Risk_Tab'
  ,'Effectiveness of AEM Program methods this equipment is based on. Check all that apply'
  ,'False'
  ,'True'
  ),
  ('Equipment'
  ,'Risk_Tab'
  ,'How is the equipment/ component used, Check one'
  ,'False'
  ,'True'
  ),
  ('Evaluation'
  ,'Risk_Tab'
  ,'Equipment / component evaluation to determine if AEM program needs to be altered.'
  ,'False'
  ,'True'
  ),
  ('Failure'
  ,'Risk_Tab'
  ,'Previous service and failure request?'
  ,'False'
  ,'True'
  ),
  ('Hospital'
  ,'Risk_Tab'
  ,'How does hospital assess if AEM Program uses appropriate maintenance strategies?'
  ,'False'
  ,'True'
  ),
  ('Maintenance'
  ,'Risk_Tab'
  ,'Are maintenance requirements simple or complex'
  ,'False'
  ,'True'
  ),
  ('Malfunction'
  ,'Risk_Tab'
  ,'Could the malfunction have been prevented?'
  ,'False'
  ,'True'
  ),
  ('Modified'
  ,'Risk_Tab'
  ,'Why are manufacture requirements being modified? Check one'
  ,'False'
  ,'True'
  ),
  ('Remove'
  ,'Risk_Tab'
  ,'Remove equipment /component from service when no longer safe or suitable for intended service if warranted.'
  ,'False'
  ,'True'
  ),
  ('Requirements'
  ,'Risk_Tab'
  ,'How much of manufactures requirements are being changed? Check one'
  ,'False'
  ,'True'
  ),
  ('Review'
  ,'Risk_Tab'
  ,'Reviewed by:'
  ,'False'
  ,'True'
  ),
  ('Score'
  ,'Risk_Tab'
  ,'Score of 8 is Classed as Critical/ High Risk'
  ,'False'
  ,'True'
  ),
  ('Seriousness'
  ,'Risk_Tab'
  ,'Seriousness and prevalence of harm during normal use:'
  ,'False'
  ,'True'
  ),
  ('Timeliness'
  ,'Risk_Tab'
  ,'Timeliness of alternate or Back-up equipment in the event of a failure. Check one'
  ,'False'
  ,'True'
  ),
  ('Why'
  ,'Risk_Tab'
  ,'Why or How?'
  ,'False'
  ,'True'
  )

-- Complete


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
  ,[Enabled]
  ,[CanModify]
  )
VALUES
  ('R1_Risk'
  ,'SERIOUSNESS'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('R10_RISK'
  ,'R10_RISK TABLE'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('R11_RISK'
  ,'R11_RISK TABLE'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('R12_RISK'
  ,'R12_RISK TABLE'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('R13_RISK'
  ,'R13_RISK TABLE'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('R14_RISK'
  ,'R14_RISK TABLE'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('R15_RISK'
  ,'R15_RISK TABLE'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('R16_RISK'
  ,'R16_RISK TABLE'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('R17_RISK'
  ,'R17_RISK TABLE'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('R19_RISK'
  ,'R19_RISK TABLE'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('R2_RISK'
  ,'R2_RISK TABLE'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('R21_RISK'
  ,'R21_RISK TABLE'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('R3_RISK'
  ,'R3_RISK TABLE'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('R6_RISK'
  ,'R6_RISK TABLE'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('R7_RISK'
  ,'R7_RISK TABLE'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('R9_RISK'
  ,'R9_RISK TABLE'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('Risk2_Risk'
  ,'CONSEQUENCES'
  ,25
  ,100
  ,'True'
  ,'True'
  ),
  ('Risk3_Risk'
  ,'CONSEQUENCES'
  ,25
  ,100
  ,'True'
  ,'True'
  )

-- Done


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

-- Done need to update the Data dic

EXEC ADM_UpdateDataDict;

-- @end