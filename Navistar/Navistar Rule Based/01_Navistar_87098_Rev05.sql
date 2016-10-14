/* 
---------------------------------------------------------------------------------------------------------- 
| Version | Author         | Date       | Comments   
----------------------------------------------------------------------------------------------------------  
| 1.0.1   | Remi Grandsire | 09/26/2016 | Original Script  
| 1.0.2   |                | 10/12/2016 | Simplify and optimize queries
| 1.0.3   |                | 10/14/2016 | Added the @AssetPK variable
| 1.0.4   | Calvin Beck    |            | Removed "joins" to use @AssetPK
| 1.0.5   | Remi Grandsire |            | Add Spes to asset before adding to the WO per Randy's suggestion
---------------------------------------------------------------------------------------------------------- 
*/

DECLARE @WOPK AS INTEGER

DECLARE @AssetPK as INTEGER
-- Setting the WOPK for the whole task 
SET @WOPK= CONVERT(VARCHAR(15),SUBSTRING('[EVENTRECORD]',CHARINDEX('=','[EVENTRECORD]')+2,(LEN('[EVENTRECORD]')-CHARINDEX('=','[EVENTRECORD]')+2)))

-- Set the AssetPK for the whole task
SELECT @AssetPK= AssetPK FROM WO WITH (nolock) WHERE WOPK= @WOPK

-- Adding specifications to the Asset if it doesn't exist already 
IF NOT EXISTS (SELECT 1
               FROM   AssetSpecification WITH (nolock)
               WHERE  SpecificationPK = 271
                      AND AssetPK = @AssetPK)
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
                   ,@AssetPK
                   ,Getdate()
                   ,1
                   ,NULL
                   ,NULL
                   ,'Engine Hours')
  END

IF NOT EXISTS (SELECT 1
               FROM   AssetSpecification WITH (nolock)
               WHERE  SpecificationPK = 100
                      AND AssetPK = @AssetPK)
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
                   ,@AssetPK
                   ,Getdate()
                   ,1
                   ,NULL
                   ,NULL
                   ,'Mileage')
  END

-- Change task # before inserting Specifications --> insert 10 and 20 */ 
UPDATE WOTask
SET    TaskNo += 20
WHERE  WOPK = @WOPK
       AND TaskNo > 9

-- Adding or updating the Mileage specification
IF EXISTS (SELECT 1
           FROM   WOTask WITH (nolock)
           WHERE  WOPK = @WOPK
                  AND SpecificationPK = 100)
  BEGIN
      UPDATE WOTask WITH ( rowlock )
      SET    TaskNo = 10
      WHERE  WOPK = @WOPK
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
      VALUES      ( @WOPK
                    ,10
                    ,'Mileage'
                    ,'1'
                    ,100
                    ,(SELECT PK 
					  FROM 	 AssetSpecification WITH (NOLOCK)
					  WHERE  SpecificationPK = 100
							 AND AssetPK = @AssetPK))
  END

IF EXISTS (SELECT 1
           FROM   WOTask WITH (nolock)
           WHERE  WOPK = @WOPK
                  AND specificationpk = 271)
  BEGIN
      UPDATE WOTask WITH (rowlock)
      SET    TaskNo = 20
      WHERE  WOPK = @WOPK
             AND SpecificationPK = 271
  END
ELSE
  BEGIN
      INSERT INTO WOTask WITH (rowlock)
                  (WOPK
                   ,TaskNo
                   ,TaskAction
                   ,Spec
                   ,SpecificationPK
                   ,AssetSpecificationPK)
      VALUES      (@WOPK
                   ,20
                   ,'Engine Hours'
                   ,'1'
                   ,271
                   ,(SELECT PK 
					  FROM 	 AssetSpecification WITH (NOLOCK)
					  WHERE  SpecificationPK = 271
							 AND AssetPK = @AssetPK))
  END
