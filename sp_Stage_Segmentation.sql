
CREATE PROCEDURE [Stage].[sp_Stage_Segmentation]
AS
	SET NOCOUNT ON;
	--Variables
	DECLARE
		@NumPerformMnths					INT = 6						--Perfmance Window (in months) initially set by requirements
		,@HighPerform						INT = 2500					--Revenue amount for High Performance
		,@MedPerform						INT = 600					--Revenue amount for Medium Performance
		,@HighPotentialRX					INT = 500					--Minimum of UHC and Medicare claims for High Potential
		,@HighPotentialRXPerc				DECIMAL(8,2) = .3			--Percentage of UHC and Medicare claims to total claims for High Potential
		,@HistoryPeriodStartDateKey			BIGINT
		,@HistoryPeriodEndDateKey			BIGINT
		,@CurrentQuarterDateKey				BIGINT
		,@CurrentQuarterDate				DATE

	--Dates
	SELECT
		@HistoryPeriodEndDateKey	= REPLACE(CONVERT(NVARCHAR(25), DATEADD(DAY, -1, [FirstDayInQuarter])), N'-', N'')
		,@HistoryPeriodStartDateKey	= REPLACE(CONVERT(NVARCHAR(25), DATEADD(MONTH, @NumPerformMnths * -1, [FirstDayInQuarter])), N'-', N'')
		,@CurrentQuarterDateKey		= REPLACE(CONVERT(NVARCHAR(25), FirstDayInQuarter), N'-', N'')
		,@CurrentQuarterDate		= FirstDayInQuarter
	FROM
		OPENQUERY([ASL-EDPRDDW-01], N'
		SELECT
			FirstDayInQuarter
		FROM
			MED.dbo.DimDate
		WHERE [Date] = CONVERT(DATE, SYSDATETIME())
		')

	DROP TABLE IF EXISTS #DecileSummary;
	DROP TABLE IF EXISTS #NPIRevenueDaily;

	CREATE TABLE #NPIRevenueDaily
	(
		NPI			NVARCHAR(10)
		,Orders		BIGINT
		,Revenue	FLOAT
	)

	INSERT INTO #NPIRevenueDaily
	(
		NPI
		,Orders
		,Revenue
	)
	(
		SELECT
			NPI			= NPI
			,Orders		= ISNULL(SUM(Orders), 0)
			,Revenue	= ISNULL(SUM(Revenue), 0.00)
		FROM
			dbo.NPIRevenueDaily

		WHERE ISNULL(RevenueDateKey, @HistoryPeriodEndDateKey) BETWEEN @HistoryPeriodStartDateKey AND @HistoryPeriodEndDateKey

		GROUP BY
			NPI
	)

	CREATE TABLE #DecileSummary
	(
		NPI				NVARCHAR(10)
		,TerritoryName	NVARCHAR(512)
		,TotalClaims	FLOAT
		,AD_RX			BIGINT
		,Ter_Perc		FLOAT
		,Ter_Rnk_Value	BIGINT
		,rnk			INT
	)

		INSERT INTO #DecileSummary
	(
		NPI
		,TerritoryName
		,TotalClaims
		,AD_RX
		,Ter_Perc
		,Ter_Rnk_Value
		,rnk
	)
	(
		SELECT
			NPI				= NPI
			,TerritoryName	= TerritoryName
			,TotalClaims	= TotalClaims
			,AD_RX			= AD_RX
			,Ter_Perc		= Ter_Perc
			,Ter_Rnk_Value	= Ter_Rnk_Value
			,rnk			= ROW_NUMBER() OVER
							(
								-- Ranking to find which territory a provider generated the most revenue in (if NPI is active accross territories)
								PARTITION BY NPI
								ORDER BY TotalClaims DESC, TerritoryName DESC
							)
		FROM
		(
			SELECT
				NPI					= NPI
				,TerritoryName		= TerritoryName
				,TotalClaims		= CAST(SUM(ISNULL(RX_CLAIMS, 0)) AS FLOAT)
				,AD_RX				= SUM(ISNULL(UhcTotalRx, 0)) + SUM(ISNULL(MedicareTotalRx, 0))
				,Ter_Perc			= CASE WHEN SUM(ISNULL(RX_CLAIMS, 0)) = 0 THEN 0 ELSE
										(SUM(ISNULL(UhcTotalRx, 0)) + SUM(ISNULL(MedicareTotalRx, 0))) / CAST(SUM(ISNULL(RX_CLAIMS, 0)) AS FLOAT)
									END
				,Ter_Rnk_Value		= CASE WHEN SUM(ISNULL(RX_CLAIMS, 0)) = 0 THEN 0 ELSE
										(SUM(ISNULL(UhcTotalRx, 0)) + SUM(ISNULL(MedicareTotalRx, 0))) / CAST(SUM(ISNULL(RX_CLAIMS, 0)) AS FLOAT)
									END * (SUM(ISNULL(UhcTotalRx, 0)) + SUM(ISNULL(MedicareTotalRx, 0)))
			FROM
				dbo.vw_DecileSummary

			GROUP BY
				NPI
				,TerritoryName
		) AS DS_1
	)

		INSERT INTO Stage.Segmentation
	(
		NPI
		,Orders
		,Revenue
		,RevenueLevel
		,TotalClaims
		,UHCMedicarePerc
		,Potential
		,SegmentType
		,ProspectRank
		,ClinicianSegment
	)
	(
		SELECT
			NPI					= NPI				
			,Orders				= Orders			
			,Revenue			= Revenue		
			,RevenueLevel		= RevenueLevel			
			,TotalClaims		= TotalClaims	
			,UHCMedicarePerc	= UHCMedicarePerc
			,Potential			= Potential
			,SegmentType		= SegmentType	
			,ProspectRank		= ProspectRank	
			,ClinicianSegment	= 
									CASE WHEN SegmentType = 'Customer' THEN
										CASE WHEN RevenueLevel = 'High Revenue' THEN 'High Potential - Thrive' ELSE
											CASE WHEN Potential = 'High Potential' THEN 'High Potential - Grow' ELSE
												CASE WHEN RevenueLevel = 'Medium Revenue' THEN 'Low Potential - Maintain' ELSE 'Low Potential - Nuture'
												END
											END
										END 
									ELSE
										CASE WHEN Potential = 'High Potential' THEN 'High Potential - Convert' ELSE 'Low Potential - Convert' 
										END
									END
		FROM
		(
		SELECT
				NPI					= COALESCE(REV.NPI, DS.NPI)
				,Orders				= ISNULL(REV.Orders, 0)
				,Revenue			= ISNULL(REV.Revenue, 0)
				,RevenueLevel		=
									CASE
										WHEN REV.Revenue >= @HighPerform THEN 'High Revenue'
										WHEN REV.Revenue >= @MedPerform THEN 'Medium Revenue'
										ELSE 'Low Revenue'
									END
				,TotalClaims		= ISNULL(DS.TotalClaims, 0)
				,UHCMedicarePerc	= CASE WHEN ISNULL(DS.TotalClaims, 0) = 0 THEN 0 ELSE ISNULL(DS.AD_RX, 0) / DS.TotalClaims END
				,Potential			=
									CASE WHEN REV.Revenue >= @HighPerform THEN 'N/A' ELSE
										CASE WHEN ISNULL(DS.TotalClaims, 0) >= @HighPotentialRX
											AND CASE WHEN ISNULL(DS.TotalClaims, 0) = 0 THEN 0 ELSE ISNULL(DS.AD_RX, 0) / DS.TotalClaims END >= @HighPotentialRXPerc
											THEN 'High Potential' ELSE 'Low Potential'
										END
									END
				,SegmentType		= CASE WHEN ISNULL(REV.Orders, 0) = 0 THEN 'Prospect' ELSE 'Customer' END
				,ProspectRank		= CASE WHEN ISNULL(REV.Orders, 0) = 0 THEN ROW_NUMBER() OVER
									(
										PARTITION BY
											DS.TerritoryName
											,CASE WHEN ISNULL(DS.TotalClaims, 0) >= @HighPotentialRX
												AND CASE WHEN ISNULL(DS.TotalClaims, 0) = 0 THEN 0 ELSE ISNULL(DS.AD_RX, 0) / DS.TotalClaims END >= @HighPotentialRXPerc
													THEN 'High Potential'
													ELSE 'Low Potential'
											END
										ORDER BY DS.Ter_Rnk_Value DESC, DS.TotalClaims DESC
									) END
			FROM
				#NPIRevenueDaily AS REV

			FULL OUTER JOIN #DecileSummary AS DS
			ON DS.NPI = REV.NPI
			AND DS.rnk = 1
		) AS SUB1
	)