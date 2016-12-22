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
  )



-- Update Data Dic:
EXEC ADM_UpdateDataDict;

-- Complete