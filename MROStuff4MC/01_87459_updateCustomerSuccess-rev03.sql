/*
#####################################################################################################################
#    Date     # Version #      Author       #           Comments
#      ??     #  1.0.1  #      ??           # Original select script created for CSM  
#  09/2016    #  1.0.2  # Remi G Grandsire  # Created the update from the select statement
# 10/26/2016  #  1.0.3  #                   # Added the RowVersionIPAddress to the statement
#################################################################################################################### 
  */

UPDATE assetspecification 
SET    assetspecification.valuetext = NULL, 
       assetspecification.valuedate = NULL, 
       assetspecification.valuenumeric = NULL, 
       rowversioninitials = '_MC', 
       rowversionipaddress = 'CSM Assignment Removed' 
FROM   asset 
       INNER JOIN assetspecification 
               ON assetspecification.assetpk = asset.assetpk 
       INNER JOIN labor 
               ON labor.laborid = assetspecification.valuetext 
WHERE  assetspecification.specificationpk = 320 
       AND valuetext IS NOT NULL 
       AND ( ( ( (SELECT Datediff(dd, (SELECT TOP 1 notedate 
                                       FROM   assetnote WITH (nolock) 
                                              LEFT JOIN labor 
                                                     ON assetnote.laborpk = 
                                                        labor.laborpk 
                                       WHERE  assetnote.assetpk = asset.assetpk 
                                              AND labor.jobtitle LIKE 
                                                  '%Success%' 
                                       ORDER  BY notedate DESC), Getdate())) > 
                 90 ) 
                OR ( NOT EXISTS((SELECT notedate 
                                 FROM   assetnote WITH (nolock) 
                                        LEFT JOIN labor 
                                               ON assetnote.laborpk = 
                                                  labor.laborpk 
                                 WHERE  assetnote.assetpk = asset.assetpk 
                                        AND labor.jobtitle LIKE '%Success%')) ) 
             ) 
             AND ( ( (SELECT Datediff(dd, (SELECT TOP 1 workdate 
                                           FROM   wolabor WITH (nolock) 
                                                  INNER JOIN wo 
                                                          ON wo.assetpk = 
                                                             asset.assetpk 
                                                  LEFT JOIN labor 
                                                         ON wolabor.laborpk = 
                                                            labor.laborpk 
                                           WHERE  wolabor.wopk = wo.wopk 
                                                  AND labor.jobtitle LIKE 
                                                      '%Success%' 
                                           ORDER  BY workdate DESC), Getdate())) 
                     > 90 ) 
                    OR ( NOT EXISTS((SELECT workdate 
                                     FROM   wolabor WITH (nolock) 
                                            INNER JOIN wo 
                                                    ON wo.assetpk = 
                                                       asset.assetpk 
                                            LEFT JOIN labor 
                                                   ON wolabor.laborpk = 
                                                      labor.laborpk 
                                     WHERE  wolabor.wopk = wo.wopk 
                                            AND labor.jobtitle LIKE '%Success%') 
                                   ) ) ) ) 