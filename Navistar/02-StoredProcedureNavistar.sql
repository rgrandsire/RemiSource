USE [entnavistar]
GO
/****** Object:  StoredProcedure [dbo].[CSTM_Navistar_GeoTabToMC]    Script Date: 11/16/2016 2:39:32 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


ALTER PROCEDURE [dbo].[CSTM_Navistar_GeoTabToMC] 
AS
BEGIN

  DECLARE
  @AssetPK INT=NULL
  ,@VehicleID VARCHAR(150)
  ,@Miles VARCHAR(25)
  ,@Hours VARCHAR(25)
  ,@ErrorMessage VARCHAR(7000)
  ,@RecordPK INT

  /************************************/
  /** Process Asset Data **************/
  /************************************/
	PRINT 'BEGIN Processing Asset data'
	DECLARE OdometerReadings CURSOR FAST_FORWARD FOR
    SELECT 
      VehicleID
      , Miles
	  , Hours
      , PK
    FROM 
      MC_InterfaceLog WITH (NOLOCK)
	WHERE
	    Processed IS NULL 
	    AND ProcessDate IS NULL
	ORDER BY 
        VehicleID ASC
  FOR READ ONLY 
  OPEN OdometerReadings 
 
  FETCH NEXT FROM OdometerReadings INTO @VehicleID, @Miles, @Hours, @RecordPK
  WHILE @@FETCH_STATUS = 0 
  BEGIN
    --Check for Asset first - error if not exist
	set @ErrorMessage= NULL
	set @AssetPK= Null
    SELECT @AssetPK = AssetPK FROM Asset WITH (NOLOCK) WHERE AssetID = @VehicleID
    
    IF @AssetPK IS NOT NULL
    BEGIN
      --Next Check current reading to make sure it is less than the new reading
	  IF (SELECT Meter1Reading FROM Asset WITH (NOLOCK) WHERE AssetPK = @AssetPK) > cast(@Miles as float) 
	  BEGIN
		set @ErrorMessage= 'MC mileage > GeoTab, original: ' + str((SELECT Meter1Reading FROM Asset WITH (NOLOCK) WHERE AssetPK = @AssetPK)) +', new: '+ @Miles+', difference= ' + str((SELECT Meter1Reading FROM Asset WITH (NOLOCK) WHERE AssetPK = @AssetPK)-cast(@Miles as float))
	  END
	  ELSE SET @ErrorMessage = NULL
	  IF (SELECT Meter2Reading FROM Asset WITH (NOLOCK) WHERE AssetPK = @AssetPK) > cast(@Hours as float) 
	  BEGIN
	   set @ErrorMessage = @ErrorMessage +' | '+ 'MC Hours > GeoTab, original: '+ str((SELECT Meter2Reading FROM Asset WITH (NOLOCK) WHERE AssetPK = @AssetPK))+' new: '+@Hours+', difference= ' + str((SELECT Meter2Reading FROM Asset WITH (NOLOCK) WHERE AssetPK = @AssetPK)-cast(@Hours as float))
	  END
	  if @ErrorMessage is null 
	  BEGIN
        UPDATE Asset WITH (ROWLOCK) SET Meter1Reading = @Miles, Meter2Reading = @Hours, RowVersionInitials='_MC', RowVersionDate= GetDate() WHERE AssetPK = @AssetPK
        UPDATE MC_InterfaceLog WITH (ROWLOCK) SET Processed = 'Y', ProcessDate = GETDATE(), MCRecordPK = @AssetPK WHERE PK = @RecordPK
      END
      ELSE
      BEGIN
        UPDATE MC_InterfaceLog WITH (ROWLOCK) SET Processed = 'N', ProcessDate = GETDATE(), ErrorMessage = ('Asset (' + @VehicleID + ') ' + @ErrorMessage) WHERE PK = @RecordPK
	  END
	END
    ELSE
    BEGIN
      UPDATE MC_InterfaceLog WITH (ROWLOCK) SET Processed = 'N', ProcessDate = GETDATE(), ErrorMessage = ('Asset (' + @VehicleID + ') does not exist') WHERE PK = @RecordPK
    END
    
    FETCH NEXT FROM OdometerReadings INTO @VehicleID, @Miles, @Hours, @RecordPK
  END 
  
  CLOSE OdometerReadings 
  DEALLOCATE OdometerReadings
  
END

