UPDATE AssetSpecification
SET AssetSpecification.ValueText = NULL, 
AssetSpecification.ValueDate = NULL, 
AssetSpecification.ValueNumeric = NULL,
RowVersionInitials = '_MC' from Asset
inner join assetspecification on  assetspecification.assetpk= asset.assetpk
inner join labor on labor.laborID = assetspecification.valuetext
where assetspecification.specificationpk = 320 
AND valuetext is not null 
AND 
(
	(
		((SELECT 
			DATEDIFF(dd,
				(SELECT TOP 1 NoteDate 
				FROM AssetNote WITH (NOLOCK) 
				LEFT JOIN Labor on AssetNote.LaborPK = Labor.LaborPK
				WHERE AssetNote.AssetPK = Asset.AssetPK AND 
				Labor.JobTitle like '%Success%'
				ORDER BY NoteDate DESC)
			,GETDATE() ) 
		) > 90
          )
		OR
		(
			NOT EXISTS(
				(SELECT NoteDate 
				FROM AssetNote WITH (NOLOCK) 
				LEFT JOIN Labor on AssetNote.LaborPK = Labor.LaborPK
				WHERE AssetNote.AssetPK = Asset.AssetPK AND 
				Labor.JobTitle like '%Success%')
			)
		)
	)

	AND

	(
		((SELECT 
			DATEDIFF(dd,
				(SELECT TOP 1 WorkDate 
				FROM WOLabor WITH (NOLOCK) 
				INNER JOIN WO ON WO.AssetPK = Asset.AssetPK 
				LEFT JOIN Labor on WOLabor.LaborPK = Labor.LaborPK
				WHERE WOLabor.WOPK = WO.WOPK AND 
				Labor.JobTitle like '%Success%'
				ORDER BY WorkDate DESC)
			,GETDATE() )
		  ) > 90 
          )
		OR
		(
			NOT EXISTS(
				(SELECT WorkDate 
				FROM WOLabor WITH (NOLOCK) 
				INNER JOIN WO ON WO.AssetPK = Asset.AssetPK 
				LEFT JOIN Labor on WOLabor.LaborPK = Labor.LaborPK
				WHERE WOLabor.WOPK = WO.WOPK AND 
				Labor.JobTitle like '%Success%')
			)
		)
	)
)