	create Type TBL_GST03_DETAIL as TABLE
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

	create Type TBL_GST03_SUMMARY  as TABLE
	(
		Point1 nvarchar(100),Point2 nvarchar(100),
		Point3a Date,Point3b Date,
		Point4 Date,
		Point5a decimal ,Point5b decimal ,
		Point6a decimal ,Point6b decimal ,
		Point7 decimal , Point8 decimal ,
		Point9 nvarchar(1),
		Point10 decimal ,Point11 decimal ,
		Point12 decimal ,Point13 decimal,
		Point14 decimal ,Point15 decimal ,
		Point16 decimal ,Point17 decimal ,
		Point18 decimal ,Point19 decimal ,
		Point19_1Code nvarchar(30),Point19_1Value decimal , Point19_1Per decimal,
		Point19_2Code nvarchar(30),Point19_2Value decimal , Point19_2Per decimal,
		Point19_3Code nvarchar(30),Point19_3Value decimal, Point19_3Per decimal,
		Point19_4Code nvarchar(30),Point19_4Value decimal, Point19_4Per decimal,
		Point19_5Code nvarchar(30),Point19_5Value decimal, Point19_5Per decimal,
		Point19_6Code nvarchar(30),Point19_6Value decimal, Point19_6Per decimal
	);