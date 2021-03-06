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