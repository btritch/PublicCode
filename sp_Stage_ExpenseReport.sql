CREATE PROCEDURE [dbo].[sp_Stage_ExpenseReport]
AS
BEGIN
	DECLARE @TransactionCounter BIGINT				= @@TRANCOUNT;
	DECLARE @TrackingId			UNIQUEIDENTIFIER	= NEWID();
	DECLARE @InsertedRows		BIGINT;
	SET NOCOUNT ON;

	--Audit

	IF @TransactionCounter > 0
		SAVE TRANSACTION StageExpenseReport
	ELSE 
		BEGIN TRANSACTION

	BEGIN TRY
		SET ANSI_WARNINGS OFF;

		--DROP ALL TEMP TABLES IF THEY EXIST
		IF OBJECT_ID('tempdb..#ExpenseReport') IS NOT NULL
		DROP TABLE #ExpenseReport

		IF OBJECT_ID('tempdb..#ExpenseReportDups') IS NOT NULL
		DROP TABLE #ExpenseReportDups

		--CREATE INITAL TEMP TABLE TO HOLD CONCUR, ORACLE, AND MANUAL  EXPENSE DATA
		CREATE TABLE #ExpenseReport
		(
			BusEntity					NVARCHAR (50)
			,REPORT_KEY					NVARCHAR (50)
			,ENTRY_KEY					NVARCHAR (100)	
			,ATTENDEE_KEY				NVARCHAR (50)	
			,PREPARER					NVARCHAR (240)	
			,EMAIL						NVARCHAR (240)	
			,START_EXPENSE_DATE			DATETIME		
			,DESCRIPTION				NVARCHAR (240)	
			,EXPENSE_TYPE				NVARCHAR (240)	
			,CAMPAIGN_ID				NVARCHAR (240)	
			,CAMPAIGN_NAME				NVARCHAR (240)	
			,ATTENDEE_TYPE				NVARCHAR (100)	
			,ATTENDEE_NAME				NVARCHAR (200)	
			,NPI						NVARCHAR (30)	
			,TOTAL						FLOAT (53)		
			,AMOUNT_CHARGED				FLOAT (53)		
			,DAILY_AMOUNT				FLOAT (53)		
			,NUMBER_ATTENDEES			FLOAT (53)		
			,LOCATION					NVARCHAR (150)	
			,ApprovalStatus				NVARCHAR (100)	
			,SourceSystem				NVARCHAR (100)	
			,UseCampaignMembers			BIT	
			,CREATION_DATE				DATETIME		
			,LAST_UPDATE_DATE			DATETIME		
		);

		INSERT INTO #ExpenseReport
		(
			BusEntity					
			,REPORT_KEY					
			,ENTRY_KEY						
			,ATTENDEE_KEY					
			,PREPARER						
			,EMAIL							
			,START_EXPENSE_DATE					
			,DESCRIPTION					
			,EXPENSE_TYPE					
			,CAMPAIGN_ID					
			,CAMPAIGN_NAME					
			,ATTENDEE_TYPE					
			,ATTENDEE_NAME					
			,NPI							
			,TOTAL								
			,AMOUNT_CHARGED						
			,DAILY_AMOUNT						
			,NUMBER_ATTENDEES				
			,LOCATION						
			,ApprovalStatus					
			,SourceSystem	
			,UseCampaignMembers						
			,CREATION_DATE									
			,LAST_UPDATE_DATE			
		)
		--DATA FROM CONCUR
		SELECT
			BUSENTITY				
			,REPORT_KEY
			,ENTRY_KEY
			,ATTENDEE_KEY			
			,PREPARER
			,EMAIL
			,START_EXPENSE_DATE
			,DESCRIPTION
			,EXPENSE_TYPE
			,CAMPAIGN_ID
			,CAMPAIGN_NAME			
			,ATTENDEE_TYPE			
			,ATTENDEE_NAME			
			,NPI					
			,TOTAL				
			,AMOUNT_CHARGED			
			,DAILY_AMOUNT
			,NUMBER_ATTENDEES
			,LOCATION
			,APPROVALSTATUS
			,SOURCESYSTEM			= N'CONCUR'
			,USECAMPAIGNMEMBERS
			,CREATION_DATE			= MIN(REPORT_DATE)
			,LAST_UPDATE_DATE		= MAX(LAST_UPDATE_DATE)
		FROM
		(
			SELECT
				BusEntity			=	ISNULL(CAST(SF.BusEntity AS varchar),
											CASE COMPANY_CODE
												WHEN 20		THEN 'MGL'	-- Oncology
												WHEN 31		THEN 'MWH'	-- Women's Health
												WHEN 80		THEN 'ARX'	-- Neuroscience
												WHEN 90		THEN 'CBI'	-- Autoimmune
															ELSE 'MGI'	-- Corp
											END
									)
		,REPORT_KEY				=	REPORT_KEY
		,ENTRY_KEY				=	ENTRY_KEY
		,ATTENDEE_KEY			=	CASE WHEN ((COMPANY_CODE = 80 OR SF.BusEntity = 'ARX') OR ((COMPANY_CODE = 90 OR SF.BusEntity = 'CBI') AND TRANSACTION_DATE > '2018-07-01') OR (TRANSACTION_DATE > '2018-09-01')) AND ISNULL(CAMPAIGN_ID, 'Crescendo') <> 'Crescendo'
										THEN null
										ELSE ATTENDEE_KEY
									END
		,PREPARER				=	EMPLOYEE_FULL_NAME
		,EMAIL					=	EMAIL_ADDRESS
		,START_EXPENSE_DATE		=	TRANSACTION_DATE
		,[DESCRIPTION]			=	REPORT_NAME
		,EXPENSE_TYPE			=	EXPENSE_TYPE
		,CAMPAIGN_ID			=	ISNULL(FLEX_VAL.ATTRIBUTE1,concur.CAMPAIGN_ID)
		,CAMPAIGN_NAME			=	CASE WHEN ((COMPANY_CODE = 80 OR SF.BusEntity = 'ARX') OR ((COMPANY_CODE = 90 OR SF.BusEntity = 'CBI') AND TRANSACTION_DATE > '2018-07-01') OR (TRANSACTION_DATE > '2018-09-01')) AND ISNULL(CAMPAIGN_ID, 'Crescendo') <> 'Crescendo'
										THEN null
										ELSE CAST(CAMPAIGN_NAME AS varchar(240))
									END
		,ATTENDEE_TYPE			=	CASE WHEN ((COMPANY_CODE = 80 OR SF.BusEntity = 'ARX') OR ((COMPANY_CODE = 90 OR SF.BusEntity = 'CBI') AND TRANSACTION_DATE > '2018-07-01') OR (TRANSACTION_DATE > '2018-09-01')) AND ISNULL(CAMPAIGN_ID, 'Crescendo') <> 'Crescendo'
										THEN null
										ELSE ATTENDEE_TYPE
									END
		,ATTENDEE_NAME			=	CAST(
									CASE WHEN ((COMPANY_CODE = 80 OR SF.BusEntity = 'ARX') OR ((COMPANY_CODE = 90 OR SF.BusEntity = 'CBI') AND TRANSACTION_DATE > '2018-07-01') OR (TRANSACTION_DATE > '2018-09-01')) AND ISNULL(CAMPAIGN_ID, 'Crescendo') <> 'Crescendo'
										THEN null
										ELSE (FIRST_NAME + ' ' + LAST_NAME)
									END AS varchar(200))
		,NPI					=	CASE WHEN ((COMPANY_CODE = 80 OR SF.BusEntity = 'ARX') OR ((COMPANY_CODE = 90 OR SF.BusEntity = 'CBI') AND TRANSACTION_DATE > '2018-07-01') OR (TRANSACTION_DATE > '2018-09-01')) AND ISNULL(CAMPAIGN_ID, 'Crescendo') <> 'Crescendo'
										THEN null
										ELSE NPI_NUMBER
									END
		,TOTAL					=	EXPENSE_AMOUNT_TOTAL
		,AMOUNT_CHARGED			=	CASE WHEN ((COMPANY_CODE = 80 OR SF.BusEntity = 'ARX') OR ((COMPANY_CODE = 90 OR SF.BusEntity = 'CBI') AND TRANSACTION_DATE > '2018-07-01') OR (TRANSACTION_DATE > '2018-09-01')) AND ISNULL(CAMPAIGN_ID, 'Crescendo') <> 'Crescendo'
										THEN null
										ELSE CASE WHEN TOTAL_EXPECTED_ATTENDEES IS NULL OR CAMPAIGN_ID IS NULL
													THEN	APPROVED_AMOUNT
												WHEN TOTAL_EXPECTED_ATTENDEES  = 0
													THEN	0
													ELSE	EXPENSE_AMOUNT_TOTAL/TOTAL_EXPECTED_ATTENDEES
											END
									END
		,DAILY_AMOUNT			=	EXPENSE_AMOUNT_TOTAL
		,NUMBER_ATTENDEES		=	TOTAL_EXPECTED_ATTENDEES
		,LOCATION				=	CITY_STATE_OF_PURCHASE
		,ApprovalStatus			=	CAST(APPROVAL_STATUS AS nvarchar(100))
		,SourceSystem			=	N'Concur'
		,REPORT_DATE			=	REPORT_DATE
		,LAST_UPDATE_DATE		=	concur.LAST_UPDATE_DATE
		,UseCampaignMembers		=	CASE WHEN ((COMPANY_CODE = 80 OR SF.BusEntity = 'ARX') OR ((COMPANY_CODE = 90 OR SF.BusEntity = 'CBI') AND TRANSACTION_DATE > '2018-07-01') OR (TRANSACTION_DATE > '2018-09-01')) AND ISNULL(CAMPAIGN_ID, 'Crescendo') <> 'Crescendo'
										THEN 1
										ELSE 0
									END
			FROM
				M1.MYXXMYR_CONCUR_INBOUND_EXP	AS CONCUR

			LEFT JOIN OD.dbo.vw_Campaigns AS SF
			ON	SF.CampaignID = CONCUR.CAMPAIGN_ID 

			LEFT JOIN M1.APPLSYS.FND_FLEX_VALUES AS FLEX_VAL
			ON CAST(FLEX_VAL.FLEX_VALUE_ID  AS VARCHAR) = CONCUR.CAMPAIGN_ID

			WHERE ISNULL(CONCUR.OPT_OUT,'NO') LIKE 'N%'
			AND ISNULL(CONCUR.SPEAKER,'NO') LIKE 'N%'
			AND 
			(
				ISNULL(FLEX_VAL.ATTRIBUTE1 ,CONCUR.CAMPAIGN_ID) IS NOT NULL
				OR CONCUR.NPI_NUMBER IS NOT NULL
			)
		) AS SUB1

		GROUP BY
			BUSENTITY
			,REPORT_KEY
			,ENTRY_KEY
			,ATTENDEE_KEY
			,PREPARER
			,EMAIL
			,START_EXPENSE_DATE
			,DESCRIPTION
			,EXPENSE_TYPE
			,CAMPAIGN_ID
			,CAMPAIGN_NAME
			,ATTENDEE_TYPE
			,ATTENDEE_NAME 
			,NPI
			,TOTAL				
			,AMOUNT_CHARGED
			,DAILY_AMOUNT
			,NUMBER_ATTENDEES
			,LOCATION
			,APPROVALSTATUS
			,USECAMPAIGNMEMBERS

		UNION ALL

		--DATA FROM ORACLE
		SELECT
			BUSENTITY =	
				CASE EX.SEGMENT1
					WHEN 3071 THEN N'ARX'
					WHEN 2051 THEN N'CBI'
					ELSE N'MGL'
				END
			,REPORT_KEY	= EX.INVOICE_NUM
			,ENTRY_KEY = CONVERT(NVARCHAR(100), EX.LINE_NUM)
			,ATTENDEE_KEY =	
				CASE WHEN EX.CAMPAIGN_ID IS NOT NULL
					THEN NULL
					ELSE NPI.NPI
				END
			,EX.PREPARER
			,EX.EMAIL
			,EX.START_EXPENSE_DATE
			,EX.DESCRIPTION
			,EXPENSE_TYPE =	EX.ITEM_DESCRIPTION
			,CAMPAIGN_ID =	
				CASE WHEN (EX.SEGMENT1 = 3071 AND EX.CAMPAIGN_ID IS NOT NULL) OR ISNULL(EX.CAMPAIGN_ID,'AA') < 'A'
					THEN EX.CAMPAIGN_ID
					ELSE CAMP.ATTRIBUTE1
				END
			,CAMPAIGN_NAME = NULL
			,ATTENDEE_TYPE = NULL
			,ATTENDEE_NAME =	
				CAST(CASE WHEN EX.CAMPAIGN_ID IS NOT NULL
					THEN NULL
					ELSE 
						CASE EX.SEGMENT1
						--	WHEN 3071 THEN NPI.NPI		--	WILL POPULATE LATER WITH A JOIN TO NPI TABLE
							WHEN 2051 THEN SF.NAME
							WHEN 8 THEN PART.PARTY_NAME
							ELSE NULL
						END
				END AS NVARCHAR(200))
			,NPI =	
				CAST(CASE WHEN EX.CAMPAIGN_ID IS NOT NULL	--EX.SEGMENT1 = 3071 AND 
					THEN NULL
					ELSE 
						CASE EX.SEGMENT1
							WHEN 3071 THEN NPI.NPI
							WHEN 2051 THEN SF.NPI_VOD
							WHEN 8 THEN PART.GLOBAL_ATTRIBUTE3
							ELSE NULL
						END
				END AS NVARCHAR(30))
			,TOTAL = EX.TOTAL
			,AMOUNT_CHARGED =	
				CONVERT(FLOAT, CASE WHEN EX.SEGMENT1 = 3071 AND EX.CAMPAIGN_ID IS NOT NULL
					THEN NULL
					ELSE DAILY_AMOUNT/NUMBER_ATTENDEES
				END)
			,DAILY_AMOUNT = EX.DAILY_AMOUNT
			,NUMBER_ATTENDEES = EX.NUMBER_ATTENDEES
			,LOCATION =	EX.LOCATION
			,APPROVALSTATUS	= N'APPROVED'
			,SOURCESYSTEM =	N'ORACLE'
			,USECAMPAIGNMEMBERS = 1
			,CREATION_DATE = EX.CREATION_DATE
			,LAST_UPDATE_DATE =	EX.LAST_UPDATE_DATE
		
		FROM		
			M1.APPS.vw_Consol_Exp AS EX

		LEFT JOIN
		(
			SELECT 		 
				P.INVOICE_NUM
				,P.LINE_NUM
				,P.SEGMENT1
				,P.NPI
			FROM		
			(
				SELECT INVOICE_NUM, LINE_NUM, SEGMENT1, ATTRIBUTE7 AS NPI FROM M1.APPS.vw_Consol_Exp UNION
				SELECT INVOICE_NUM, LINE_NUM, SEGMENT1, ATTRIBUTE8 AS NPI FROM M1.APPS.vw_Consol_Exp UNION
				SELECT INVOICE_NUM, LINE_NUM, SEGMENT1, ATTRIBUTE9 AS NPI FROM M1.APPS.vw_Consol_Exp UNION
				SELECT INVOICE_NUM, LINE_NUM, SEGMENT1, ATTRIBUTE10 AS NPI	FROM M1.APPS.vw_Consol_Exp UNION
				SELECT INVOICE_NUM, LINE_NUM, SEGMENT1, ATTRIBUTE11 AS NPI	FROM M1.APPS.vw_Consol_Exp UNION
				SELECT INVOICE_NUM, LINE_NUM, SEGMENT1, ATTRIBUTE12 AS NPI	FROM M1.APPS.vw_Consol_Exp UNION
				SELECT INVOICE_NUM, LINE_NUM, SEGMENT1, ATTRIBUTE13 AS NPI	FROM M1.APPS.vw_Consol_Exp UNION
				SELECT INVOICE_NUM, LINE_NUM, SEGMENT1, ATTRIBUTE14 AS NPI	FROM M1.APPS.vw_Consol_Exp UNION
				SELECT INVOICE_NUM, LINE_NUM, SEGMENT1, ATTRIBUTE15 AS NPI	FROM M1.APPS.vw_Consol_Exp 
			) AS P
			WHERE NPI IS NOT NULL
		) AS NPI
		ON	NPI.INVOICE_NUM = EX.INVOICE_NUM
		AND	NPI.LINE_NUM = EX.LINE_NUM

		LEFT JOIN SALESFORCE_CBI.DBO.SF_ACCOUNT AS SF
		ON NPI.NPI = SF.ID

		LEFT JOIN M1.AR.HZ_PARTIES AS PART
		ON NPI.NPI = CAST(PART.PARTY_ID AS NVARCHAR(150))

		LEFT JOIN M1.APPLSYS.FND_FLEX_VALUES AS CAMP
		ON EX.CAMPAIGN_ID =	CAMP.FLEX_VALUE

		WHERE EX.CAMPAIGN_ID IS NOT NULL
		OR NPI.NPI IS NOT NULL

		GROUP BY 
			CASE EX.SEGMENT1
				WHEN 3071 THEN N'ARX'
				WHEN 2051 THEN N'CBI'
				ELSE N'MGL'
			END
			,EX.INVOICE_NUM
			,EX.LINE_NUM
			,CASE WHEN EX.CAMPAIGN_ID IS NOT NULL
				THEN NULL
				ELSE NPI.NPI
			END
			,EX.PREPARER
			,EX.EMAIL
			,EX.START_EXPENSE_DATE
			,EX.DESCRIPTION
			,EX.ITEM_DESCRIPTION
			,CASE WHEN (EX.SEGMENT1 = 3071 AND EX.CAMPAIGN_ID IS NOT NULL) OR ISNULL(EX.CAMPAIGN_ID,'AA') < 'A'
				THEN EX.CAMPAIGN_ID
				ELSE CAMP.ATTRIBUTE1
			END
			,CAST(CASE WHEN EX.CAMPAIGN_ID IS NOT NULL
				THEN NULL
				ELSE 
					CASE EX.SEGMENT1
					--	WHEN 3071 THEN NPI.NPI		--	WILL POPULATE LATER WITH A JOIN TO NPI TABLE
						WHEN 2051 THEN SF.NAME
						WHEN 8 THEN PART.PARTY_NAME
						ELSE NULL
					END
			END AS NVARCHAR(200))
			,CAST(CASE WHEN EX.CAMPAIGN_ID IS NOT NULL	--EX.SEGMENT1 = 3071 AND 
				THEN NULL
				ELSE 
					CASE EX.SEGMENT1
						WHEN 3071 THEN NPI.NPI
						WHEN 2051 THEN SF.NPI_VOD
						WHEN 8 THEN PART.GLOBAL_ATTRIBUTE3
						ELSE NULL
					END
			END AS NVARCHAR(30))
			,EX.TOTAL
			,CASE WHEN EX.SEGMENT1 = 3071 AND EX.CAMPAIGN_ID IS NOT NULL
				THEN NULL
				ELSE DAILY_AMOUNT/NUMBER_ATTENDEES
			END
			,EX.DAILY_AMOUNT
			,EX.NUMBER_ATTENDEES
			,EX.LOCATION
			,EX.CREATION_DATE
			,EX.LAST_UPDATE_DATE

			UNION ALL

			--DATA FROM  CORPORATE EXPENSE (MANUAL)
			SELECT		 
				BUSENTITY					=	
												CASE LEFT(COMPANY_CODE,1)
													WHEN '2' THEN N'MGL'	-- ONCOLOGY
													WHEN '3' THEN N'WH'		-- WOMEN'S HEALTH
													WHEN '8' THEN N'ARX'	-- NEUROSCIENCE
													WHEN '9' THEN N'CBI'	-- AUTOIMMUNE
													ELSE N'MGI'				-- CORP
												END
				,REPORT_KEY					=	N'MGMANUALEXP_' + CONVERT(NVARCHAR, EVENT_DATE, 110)
				,ENTRY_KEY					=	EVENT_NAME
				,ATTENDEE_KEY				=	NPI
				,PREPARER					=	N'CORPORATE ENTRY'
				,EMAIL						=	NULL
				,START_EXPENSE_DATE			=	EVENT_DATE
				,DESCRIPTION				=	EVENT_NAME
				,EXPENSE_TYPE				=	N'MANUAL CORPORATE EXPENSE'
				,CAMPAIGN_ID				=	NULL
				,CAMPAIGN_NAME				=	NULL
				,ATTENDEE_TYPE				=	N'BUSINESS GUEST'
				,ATTENDEE_NAME				=	FIRST_NAME + ' ' + LAST_NAME
				,NPI						=	NPI
				,TOTAL						=	AMOUNT
				,AMOUNT_CHARGED				=	AMOUNT
				,DAILY_AMOUNT				=	AMOUNT
				,NUMBER_ATTENDEES			=	(SELECT COUNT(*) FROM M1.MYXXMYR__CORPORATE_EXPENSES WHERE EX.EVENT_NAME = EVENT_NAME)
				,LOCATION					=	NULL
				,APPROVALSTATUS				=	NULL
				,SOURCESYSTEM				=	N'CORPEXP'
				,USECAMPAIGNMEMBERS			=	0
				,CREATION_DATE				=	NULL
				,LAST_UPDATE_DATE			=	NULL
		FROM	
			M1.MYXXMYR__CORPORATE_EXPENSES AS EX;

		--SELECT TOP 100 * FROM #ExpenseReport

		--CREATE TEMP TABLE TO REMOVE DUPLICATE  DATA
		CREATE TABLE #ExpenseReportDups
		(
			BUSENTITY					NVARCHAR(50)	
			,REPORTKEY					NVARCHAR(50)	
			,ENTRYKEY					NVARCHAR(50)	
			,ATTENDEEKEY				NVARCHAR(50)	
			,EMPLOYEE					NVARCHAR(240)	
			,REPORTNAME					NVARCHAR(240)	
			,EXPENSETYPE				NVARCHAR(240)	
			,TRANSACTIONDATE			DATE			
			,ATTENDEEGUESTTYPE			NVARCHAR(240)	
			,ATTENDEENAME				NVARCHAR(360)	
			,NPI						NVARCHAR(150)	
			,ISNPIMATCH					NVARCHAR(3)		
			,AMOUNT						FLOAT(53)		
			,EXPENSETOTAL				FLOAT(53)		
			,LOCATION					NVARCHAR(240)	
			,CAMPAIGNID					NVARCHAR(240)	
			,SOURCESYSTEM				NVARCHAR(100)	
			,APPROVALSTATUS				NVARCHAR(100)	
			,DATEFIRSTSUBMITTED			DATE			
			,POSTEDDATE					DATE					
		);

		INSERT INTO #ExpenseReportDups
		(
			BusEntity
			,ReportKey
			,EntryKey
			,AttendeeKey
			,Employee
			,ReportName
			,ExpenseType
			,TransactionDate
			,AttendeeGuestType
			,AttendeeName
			,NPI
			,isNPIMatch
			,Amount
			,ExpenseTotal
			,Location
			,CampaignID
			,SourceSystem
			,ApprovalStatus
			,DateFirstSubmitted
			,PostedDate
		)

		SELECT		 
			BusEntity							=	ISNULL(SF.BusEntity,stg.BusEntity)
			,ReportKey							=	stg.REPORT_KEY
			,EntryKey							=	stg.ENTRY_KEY
			,AttendeeKey						=	COALESCE(SF.ContactId,stg.ATTENDEE_KEY)
			,Employee							=	ISNULL(stg.PREPARER,N'')
			,ReportName							=	ISNULL(stg.[DESCRIPTION],N'none')
			,ExpenseType						=	ISNULL(stg.EXPENSE_TYPE,N'none')
			,TransactionDate					=	ISNULL(stg.START_EXPENSE_DATE, N'1900-01-01')
			,AttendeeGuestType					=	COALESCE(SF.GuestType,stg.ATTENDEE_TYPE)
			,AttendeeName						=	COALESCE(SF.Name,hcp.Name,hcp2.Name,stg.ATTENDEE_NAME)
			,NPI								=	COALESCE(hcp.NPI,hcp2.NPI,SF.NPI,stg.NPI)
			,isNPIMatch							=	CASE WHEN hcp2.NPI IS NULL AND hcp.NPI IS NULL
														THEN N'No'
														ELSE N'Yes'
													END
			,Amount								=	CASE WHEN stg.UseCampaignMembers = 0
														THEN stg.AMOUNT_CHARGED 
														ELSE	
															CASE WHEN ISNULL(stg.NUMBER_ATTENDEES,0) = 0
																THEN 
																	CASE WHEN ISNULL((SELECT COUNT(ContactId) FROM OD.dbo.vw_CampaignMembers WHERE CampaignID = stg.CAMPAIGN_ID),0) = 0
																		THEN stg.AMOUNT_CHARGED
																		ELSE  stg.DAILY_AMOUNT/(SELECT COUNT(ContactId) FROM OD.dbo.vw_CampaignMembers WHERE CampaignID = stg.CAMPAIGN_ID)
																	END
																ELSE stg.DAILY_AMOUNT / stg.NUMBER_ATTENDEES
															END
													END
			,ExpenseTotal						=	stg.TOTAL
			,Location							=	stg.LOCATION
			,CampaignID							=	stg.CAMPAIGN_ID
			,SourceSystem						=	stg.SourceSystem
			,ApprovalStatus						=	stg.ApprovalStatus
			,DateFirstSubmitted					=	stg.CREATION_DATE
			,PostedDate							=	stg.LAST_UPDATE_DATE
		FROM
			#ExpenseReport AS stg
		
		LEFT JOIN OD.dbo.vw_CampaignMembers AS SF
		ON SF.CampaignID = stg.CAMPAIGN_ID
		AND	stg.UseCampaignMembers = 1			
		AND	ISNULL(SF.MemberStatus,'') <> N'Y'
		
		LEFT JOIN OD.dbo.NPI AS hcp -- validate NPI entered in SalesForce
		ON hcp.NPI = SF.NPI
		
		LEFT JOIN OD.dbo.NPI AS hcp2 -- for non-campaigns and older MGL/CBI campaign members listed on expense
		ON hcp2.NPI = stg.NPI
				
		WHERE COALESCE(SF.ContactId,stg.ATTENDEE_KEY) IS NOT NULL
		AND	stg.REPORT_KEY IS NOT NULL
		AND	stg.ENTRY_KEY IS NOT NULL
		AND	ISNULL(stg.ATTENDEE_TYPE,'') <>	N'No Shows';

		--INSERT DATA INTO STAGING TABLE
		INSERT INTO OD_Stage.dbo.ExpenseReport -- ORIGINAL STAGING TABLE
		(
			BusEntity
			,ReportKey
			,EntryKey
			,AttendeeKey
			,Employee
			,ReportName
			,ExpenseType
			,TransactionDate
			,AttendeeGuestType
			,AttendeeName
			,NPI
			,isNPIMatch
			,Amount
			,ExpenseTotal
			,Location
			,CampaignID
			,SourceSystem
			,ApprovalStatus
			,DateFirstSubmitted
			,PostedDate
		)
		SELECT
			2.BusEntity
			,2.ReportKey
			,2.EntryKey
			,2.AttendeeKey
			,2.Employee
			,2.ReportName
			,2.ExpenseType
			,2.TransactionDate
			,2.AttendeeGuestType
			,2.AttendeeName
			,2.NPI
			,2.isNPIMatch
			,2.Amount
			,2.ExpenseTotal
			,2.Location
			,2.CampaignID
			,2.SourceSystem
			,2.ApprovalStatus
			,2.DateFirstSubmitted
			,2.PostedDate
			FROM
			#ExpenseReportDups AS 2

		SET @InsertedRows = @@ROWCOUNT;

		-- LOAD MRAZEK EXPENSE DATA TO STAGE
		INSERT INTO OD_Stage.dbo.ExpenseReport
		(
			BusEntity
			,ReportKey
			,EntryKey
			,AttendeeKey
			,Employee
			,TransactionDate
			,ReportName
			,ExpenseType
			,CampaignID
			,AttendeeGuestType
			,AttendeeName
			,NPI
			,isNPIMatch
			,Amount
			,Location
			,ApprovalStatus
			,SourceSystem
			,DateFirstSubmitted
			,PostedDate
		)
		SELECT
			BusEntity
			,ReportKey							=	N'MrazekBook_' + CONVERT(NVARCHAR, TransactionDate, 110)
			,EntryKey							=	N'MrazekBook_' + AttendeeName
			,AttendeeKey						=	stg.NPI
			,Employee							=	Employee
			,TransactionDate					=	TransactionDate
			,ReportName							=	N'Mrazek Book'
			,ExpenseType						=	ExpenseType
			,CampaignID							=	NULL
			,AttendeeGuestType					=	AttendeeGuestType
			,AttendeeName						=	AttendeeName
			,NPI								=	stg.NPI
			,isNPIMatch							=	CASE WHEN hcp.NPI IS NULL
															THEN N'No'
															ELSE N'Yes'
													END
			,Amount								=	Amount
			,Location							=	NULL
			,ApprovalStatus						=	NULL
			,SourceSystem						=	N'MrazekBook'
			,DateFirstSubmitted					=	CreatedDate
			,PostedDate							=	LastUpdatedDate
		FROM		
			OD_Stage.dbo.ExpenseMrazekBook	AS stg

		LEFT JOIN OD.dbo.NPI AS hcp
		ON	hcp.NPI = stg.NPI

		WHERE InActiveDate IS NULL;

		SET @InsertedRows = @InsertedRows + @@ROWCOUNT;

			IF @TransactionCounter = 0
			COMMIT TRANSACTION;
		SET ANSI_WARNINGS ON;
	END TRY

	BEGIN CATCH
		-- add any rollback logic before we insert the error data
		IF @TransactionCounter = 0
			ROLLBACK TRANSACTION;
		ELSE
			IF XACT_STATE() <> -1
				ROLLBACK TRANSACTION StageExpenseReport;

		DECLARE @ErrorMessage	NVARCHAR(4000)	= ERROR_MESSAGE();
		DECLARE @ErrorSeverity	INT				= ERROR_SEVERITY();
		DECLARE @ErrorState		INT				= ERROR_STATE();

		RAISERROR
		(
			@ErrorMessage
			,@ErrorSeverity
			,@ErrorState
		);
	END CATCH
END