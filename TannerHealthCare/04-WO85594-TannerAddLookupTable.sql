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