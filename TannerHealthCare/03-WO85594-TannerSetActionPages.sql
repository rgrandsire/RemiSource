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