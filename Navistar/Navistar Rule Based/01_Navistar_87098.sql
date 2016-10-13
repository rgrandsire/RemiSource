/* 
---------------------------------------------------------------------------------------------------------- 
| Version | Author         | Date       | Comments   
----------------------------------------------------------------------------------------------------------  
| 1.0.1   | Remi Grandsire | 09/26/2016 | Original Script  
| 1.0.2   |                | 10/12/2016 | Simplify and optimize queries
---------------------------------------------------------------------------------------------------------- 
*/
-- Setting the POPK for the whole task 
DECLARE @zPOPK AS INTEGER

SET @zPOPK= Replace(
'WO.WOPK = WO.WOPK = WO.WOPK = WO.WOPK = WO.WOPK = [EVENTRECORD]',
    'WO.WOPK = ', '')

-- Change task # before inserting Specifications --> insert 10 and 20 */ 
UPDATE WOTask
SET    TaskNo += 20
WHERE  WOPK = @zPOPK
       AND TaskNo > 9

-- Adding or updating the Mileage specification
IF EXISTS (SELECT 1
           FROM   WOTask WITH (nolock)
           WHERE  WOPK = @zPOPK
                  AND SpecificationPK = 100)
  BEGIN
      UPDATE WOTask WITH ( rowlock )
      SET    TaskNo = 10
      WHERE  WOPK = @zPOPK
             AND SpecificationPK = 100
  END
ELSE
  BEGIN
      INSERT INTO WOTask WITH ( rowlock )
                  (WOPK
                   ,TaskNo
                   ,TaskAction
                   ,Spec
                   ,SpecificationPK
                   ,AssetSpecificationPK)
      VALUES      ( @zPOPK
                    ,10
                    ,'Mileage'
                    ,'1'
                    ,100
                    ,(SELECT SpecificationPK
                      FROM   AssetSpecification WITH (nolock)
                      WHERE  PK = 100))
  END

IF EXISTS (SELECT 1
           FROM   WOTask WITH (nolock)
           WHERE  WOPK = @zPOPK
                  AND specificationpk = 271)
  BEGIN
      UPDATE WOTask WITH ( rowlock )
      SET    TaskNo = 20
      WHERE  WOPK = @zPOPK
             AND SpecificationPK = 271
  END
ELSE
  BEGIN
      INSERT INTO WOTask WITH ( rowlock )
                  (WOPK
                   ,TaskNo
                   ,TaskAction
                   ,Spec
                   ,SpecificationPK
                   ,AssetSpecificationPK)
      VALUES      (@zPOPK
                   ,20
                   ,'Engine Hours'
                   ,'1'
                   ,271
                   ,(SELECT SpecificationPK
                     FROM   AssetSpecification WITH (nolock)
                     WHERE  PK = 271))
  END

-- Adding specifications to the task list if it doesn't exist already 
IF NOT EXISTS (SELECT 1
               FROM   AssetSpecification As1 WITH (nolock)
                      INNER JOIN WO WO1
                              ON As1.assetpk = WO1.AssetPK
                      LEFT OUTER JOIN Asset As2
                                   ON As1.AssetPK = As2.AssetPK
               WHERE  SpecificationPK = 271
                      AND WO1.WOPK = @zPOPK)
  BEGIN
      INSERT INTO AssetSpecification WITH (rowlock)
                  (SpecificationPK
                   ,RowVersionInitials
                   ,AssetPK
                   ,ValueDate
                   ,TrackHistory
                   ,ValueLow
                   ,ValueHi
                   ,SpecificationName)
      VALUES      (271
                   ,'_MC'
                   ,(SELECT AssetPK
                     FROM   WO WITH (nolock)
                     WHERE  WOPK = @zPOPK)
                   ,Getdate()
                   ,1
                   ,NULL
                   ,NULL
                   ,'Engine Hours')
  END

IF NOT EXISTS (SELECT 1
               FROM   AssetSpecification As1 WITH (nolock)
                      INNER JOIN WO WO1
                              ON As1.AssetPK = WO1.AssetPK
                      LEFT OUTER JOIN Asset As2
                                   ON As1.AssetPK = As2.AssetPK
               WHERE  SpecificationPK = 100
                      AND WO1.WOPK = @zPOPK)
  BEGIN
      INSERT INTO AssetSpecification WITH (rowlock)
                  (SpecificationPK
                   ,RowVersionInitials
                   ,AssetPK
                   ,ValueDate
                   ,TrackHistory
                   ,ValueLow
                   ,ValueHi
                   ,SpecificationName)
      VALUES      (100
                   ,'_MC'
                   ,(SELECT AssetPK
                     FROM   WO WITH (nolock)
                     WHERE  WOPK = @zPOPK)
                   ,Getdate()
                   ,1
                   ,NULL
                   ,NULL
                   ,'Mileage')
  END 