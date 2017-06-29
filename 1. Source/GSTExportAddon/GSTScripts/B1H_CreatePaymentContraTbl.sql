declare v_tab_exists smallint := 0;
SELECT table_name FROM TABLES WHERE  Schema_Name= CURRENT_SCHEMA AND table_name = 'TBL_PAYMENTCONTRA';
--USING TABLE FOR PAYMENT CONTRA	
v_tab_exists := ::ROWCOUNT;
IF v_tab_exists > 0 THEN

	create column table TBL_PAYMENTCONTRA
	(
		PointAt nvarchar(10), 
		TransID decimal, 
		VatGroup nvarchar(20), 
		Debit decimal, 
		Credit decimal, 
		Balance decimal,
		BaseSum decimal, 
		VatRat decimal, 
		Memo nvarchar(100), 
		transtype nvarchar(10)
	);
END IF;