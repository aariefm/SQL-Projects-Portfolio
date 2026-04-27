-- Rev02 Update once more the formatting to be easily digestable for VBA Script

SELECT 
	 COALESCE(CAST(T0.OriginNum AS VARCHAR), 'MFG-Internal') AS "SalesOrder"	-- For Internal Manufacturing Tasks with no Sales Order No., call it MFG-Internal
	,COALESCE(T1.CardName, 'MFG-Internal') AS "CustomerName"					-- For Internal Manufacturing Tasks with no Customer., call it MFG-Internal
	,' ' AS "Task"
	,CAST(T0.PlannedQty AS INT) AS "QTY"				-- Simplify as just 'QTY'
	,T0.Comments AS "Description"     					-- The comments/remarks is mainly used as product description and quantity so heading should reflect this.
	,T0.DocNum AS "WO#"			-- Production Order Number or WO# column will be left of QTY and DESCRIPTION.
	,' ' AS "PartPulled?"		-- Empty placeholder for columns in the target spreadsheet
	,' ' AS "ShopDueDate"		-- Empty placeholder for columns in the target spreadsheet
	,' ' AS "Initials"			-- Empty placeholder for columns in the target spreadsheet
	,' ' AS "Completed",		-- Empty placeholder for columns in the target spreadsheet
	
	-- Use CASE Expression to allow sorting of work orders by product type.
	CASE 
		WHEN	T2.ItmsGrpCod LIKE '101%'
				OR T2.ItmsGrpCod LIKE '112%'
				OR T2.ItmsGrpCod LIKE '113%'
				THEN 'MachineShop'	
		WHEN T2.ItmsGrpCod LIKE '143%' THEN 'PumpBay'
		WHEN	T2.ItmsGrpCod LIKE '141%'
				OR T2.ItmsGrpCod LIKE '152%'
				OR T2.ItmsGrpCod LIKE '190%' THEN 'MotorBay'
		WHEN T2.ItmsGrpCod LIKE '142%'
				OR T2.ItmsGrpCod LIKE '144%' 
				OR T2.ItmsGrpCod LIKE '159%'
				OR T2.ItmsGrpCod LIKE '195%' THEN 'Protector'
		WHEN	T2.ItmsGrpCod LIKE '205%'
				OR T2.ItmsGrpCod LIKE '135%'
				OR T2.ItmsGrpCod LIKE '248%' THEN 'CableBay/FSB'
		END AS "Area"				

FROM OWOR T0  								--Main table to display must be declared* before the joins, note it appears on the left side of the ON statement*/

	LEFT JOIN ORDR T1 ON T0.OriginNum = T1.DocNum
	INNER JOIN OITM T2 ON T0.ItemCode = T2.ItemCode					
	
WHERE 
	-- T0.[PostDate] between [%1] AND [%2] AND 			--Decision to omit a date condition as we are only interested in PLANNED and RELEASED #WO
	T0.[Status] IN ('R', 'P')							--RELEASED and PLANNED production orders only.*/

ORDER BY Area  DESC, T0.OriginNum ASC;			--Sort by the group classification first, then by descending SO#
.sn