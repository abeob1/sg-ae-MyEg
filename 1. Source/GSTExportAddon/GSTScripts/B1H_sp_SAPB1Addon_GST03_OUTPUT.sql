drop PROCEDURE SP_SAPB1ADDON_PAYMENTCONTRA
create PROCEDURE SP_SAPB1ADDON_GST03_OUTPUT
--**{0}**--
(
	IN pointat nvarchar(10),
	IN fromdate Date,
	IN todate Date,
	IN duedate Date,
	OUT GSTDetail TBL_GST03_DETAIL,
	OUT GSTSummary TBL_GST03_SUMMARY
)
LANGUAGE SQLSCRIPT
AS
BEGIN
	DECLARE i decimal:=1;
	declare maxrow decimal:=0;
	declare minrow decimal:=0;
	
	
	select min("$rowid$") into minrow from "@GST_MSIC";
	select max("$rowid$") into maxrow from "@GST_MSIC";
	
	create local temporary table #GST03DATATBL
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

	insert into #GST03DataTbl(Point1,Point2,Point3a,Point3b, Point4)
	select "TaxIdNum","CompnyName",fromdate,todate,duedate from "OADM";
	
	create local temporary table #TBLTEMP(
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
	
	create local temporary table #TBLMISC (
		ID integer, 
		Code nvarchar(10), 
		Percentage decimal
	);
--------------------------------------POINT 5A----------------------------------	
IF :pointat='5a' or :pointat='A' or :pointat='All' THEN
		insert into #TBLTEMP
		select '5a', T0."SrcObjAbs", T0."Code", 0,0,
		case when T0."CANCELED"='C' then
		(-1)* 
		(
			case when T0."Category"='O' then
				case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
			else
				case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
			End	
		)
		else
		(
			case when T0."Category"='O' then
				case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
			else
				case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
			End	
		) END "VatSum",
		
		case when T0."CANCELED"='C' then
		(-1)* 
		(
			case when T0."Category"='O' then
				case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
			else
				case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
			End	
		)
		else
		(
			case when T0."Category"='O' then
				case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
			else
				case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
			End	
		) END "BaseSum", T0."VatPercent",T0."JrnlMemo",T0."SrcObjType"
		from "B1_VatView" T0 
		Join "OVTG" T1 on T0."Code"=T1."Code" 
		where "DocDate" between :fromdate and :todate
		and T1."ReportCode" in ('SR','DS');
	END IF;
	update #GST03DATATBL set Point5a=(Select SUM(BaseSum) from #TBLTEMP where PointAt='5a');
----------------END POINT 5a-----------------

----------------BEGIN POINT 5b-----------------
if :pointat='5b' or :pointat='All' or :pointat='A' then
	insert into #TBLTEMP
	select '5b', T0."SrcObjAbs", T0."Code", 0,0,
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	) END "VatSum",
	
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	) END "BaseSum", T0."VatPercent",T0."JrnlMemo",T0."SrcObjType"
	from "B1_VatView" T0 
	Join "OVTG" T1 on T0."Code"=T1."Code"
	where "DocDate" between :fromdate and :todate
	and T1."ReportCode" in ('SR','DS','AJS');
end if;
update #GST03DATATBL set Point5b=(select SUM(Balance) from #TBLTEMP where PointAt='5b');
----------------END POINT 5b-----------------

 ----------------BEGIN POINT 6a-----------------
if :pointat='6a' or :pointat='All' or :pointat='A' then
	insert into #TBLTEMP
	select '6a',T0."SrcObjAbs", T0."Code", 0,0,
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	) END "VatSum",
	
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	) END "BaseSum", T0."VatPercent",T0."JrnlMemo",T0."SrcObjType"
	from "B1_VatView" T0 
	Join "OVTG" T1 on T0."Code"=T1."Code"
	where "DocDate" between :fromdate and :todate
		and T1."ReportCode" in ('TX', 'IM','TX-E43', 'TX-RE','RC');
end if;
update #GST03DATATBL set Point6a=(select SUM(BaseSum) from #TBLTEMP where PointAt='6a');
----------------END POINT 6a-----------------
 
----------------BEGIN POINT 6b-----------------
if :pointat='6b' or :pointat='All' or :pointat='A' then
	insert into #TBLTEMP
	select '6b',T0."SrcObjAbs", T0."Code", 0,0,
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	) END "VatSum",
	
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	) END "BaseSum", T0."VatPercent",T0."JrnlMemo",T0."SrcObjType"
	from "B1_VatView" T0 
	Join "OVTG" T1 on T0."Code"=T1."Code" 
	Where "DocDate" between :fromdate and :todate
	and T1."ReportCode" in ('TX', 'IM','TX-E43', 'TX-RE','AJP','RC') ;
end if;
update #GST03DATATBL set Point6b=(select SUM(Balance) from #TBLTEMP where PointAt='6b');
----------------END POINT 6b-----------------

----------------BEGIN POINT 7,8--------------
update #GST03DATATBL set Point7=Point5b-Point6b, Point8=Point6b-Point5b;
---------------END POINT 7,8-------------

----------------BEGIN POINT 10-----------------
if :pointat='10' or :pointat='All' or :pointat='A'  then
	insert into #TBLTEMP
	select '10',T0."SrcObjAbs", T0."Code", 0,0,
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	) END "VatSum",
	
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	) END "BaseSum", T0."VatPercent",T0."JrnlMemo",T0."SrcObjType"
	from "B1_VatView" T0 
	Join "OVTG" T1 on T0."Code"=T1."Code" 
	where "DocDate" between :fromdate and :todate
	and T1."ReportCode" in ('ZRL');
end if;
update #GST03DATATBL set Point10=(select SUM(BaseSum) from #TBLTEMP where PointAt='10');
----------------END POINT 10-----------------

----------------BEGIN POINT 11-----------------
if :pointat='11' or :pointat='All' or :pointat='A' then
	insert into #TBLTEMP
	select '11',T0."SrcObjAbs", T0."Code", 0,0,
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	) END "VatSum",
	
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	) END BaseSum, T0."VatPercent",T0."JrnlMemo",T0."SrcObjType"
	from "B1_VatView" T0 
	Join "OVTG" T1 on T0."Code"=T1."Code" 
	where "DocDate" between :fromdate and :todate
	and T1."ReportCode" in ('ZRE');
end if;
update #GST03DATATBL set Point11=(select SUM(BaseSum) from #TBLTEMP where PointAt='11');
----------------END POINT 11-----------------

----------------BEGIN POINT 12-----------------
if :pointat='12' or :pointat='All' or :pointat='A' then
	insert into #TBLTEMP
	select '12', T0."SrcObjAbs", T0."Code", 0,0,
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	) END "VatSum",
	
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	) END "BaseSum", T0."VatPercent",T0."JrnlMemo",T0."SrcObjType"
	from "B1_VatView" T0 
	Join "OVTG" T1 on T0."Code"=T1."Code" where T0."Category"='O' 
	AND "DocDate" between :fromdate and :todate
	and T1."ReportCode" in ('ES43','ES');
end if;
update #GST03DATATBL set Point12=(select SUM(BaseSum) from #TBLTEMP where PointAt='12');
----------------END POINT 12-----------------

----------------BEGIN POINT 13-----------------
if :pointat='13' or :pointat='All' or :pointat='A' then
	insert into #TBLTEMP
	select '13',T0."SrcObjAbs", T0."Code", 0,0,
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	) END "VatSum",
	
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	) END "BaseSum", T0."VatPercent",T0."JrnlMemo",T0."SrcObjType"
	from "B1_VatView" T0 
	Join "OVTG" T1 on T0."Code"=T1."Code" 
	where "DocDate" between :fromdate and :todate
	and T1."ReportCode" in ('RS');
end if;
update #GST03DATATBL set Point13=(select SUM(BaseSum) from #TBLTEMP where PointAt='13');
----------------END POINT 13-----------------

------------------BEGIN POINT 14-----------------
if :pointat='14' or :pointat='All' or :pointat='A'  then
	insert into #TBLTEMP
	select '14',T0."SrcObjAbs", T0."Code", 0,0,
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	) END "VatSum",
	
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	) END "BaseSum", T0."VatPercent",T0."JrnlMemo",T0."SrcObjType"
	from "B1_VatView" T0 
	Join "OVTG" T1 on T0."Code"=T1."Code" 
	where "DocDate" between :fromdate and :todate
	and T1."ReportCode" in ('IS');
end if;
update #GST03DATATBL set Point14=(select SUM(Balance) from #TBLTEMP where PointAt='14');
------------------END POINT 14-----------------

------------------BEGIN POINT 15-----------------
if :pointat='15' or :pointat='All' or :pointat='A' then
	insert into #TBLTEMP
	select '15',T0."SrcObjAbs", T0."Code", 0,0,
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	) END "VatSum",
	
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	) END "BaseSum", T0."VatPercent",T0."JrnlMemo",T0."SrcObjType"
	from "B1_VatView" T0 
	Join "OVTG" T1 on T0."Code"=T1."Code" 
	where "DocDate" between :fromdate and :todate
	and T1."ReportCode" in ('IS');
end if;
update #GST03DATATBL set Point15=(select SUM(BaseSum) from #TBLTEMP where PointAt='15');
------------------END POINT 15-----------------

------------------BEGIN POINT 16-----------------
if :pointat='16' or :pointat='All' or :pointat='A'  then
	insert into #TBLTEMP
	select '16',T0."SrcObjAbs", T0."Code", 0,0,
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	) END "VatSum",
	
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	) END "BaseSum", T0."VatPercent",T0."JrnlMemo",T0."SrcObjType"
	from "B1_VatView" T0 
	Join "OVTG" T1 on T0."Code"=T1."Code" 
	where "DocDate" between :fromdate and :todate
	and T1."ReportCode" in ('TX-CG');

	---FA ONLY--
end if;
update #GST03DATATBL set Point16=(select SUM(Balance) from #TBLTEMP where PointAt='16');
------------------END POINT 16-----------------

------------------BEGIN POINT 17-----------------
if :pointat='17' or :pointat='All' or :pointat='A' then
	insert into #TBLTEMP
	select '17',T0."SrcObjAbs", T0."Code", 0,0,
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	) END "VatSum",
	
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	) END "BaseSum", T0."VatPercent",T0."JrnlMemo",T0."SrcObjType"
	from "B1_VatView" T0 
	Join "OVTG" T1 on T0."Code"=T1."Code"
	where "DocDate" between :fromdate and :todate
	and T1."ReportCode" in ('AJP');
end if;
update #GST03DATATBL set Point17=(select SUM(Balance+BaseSum) from #TBLTEMP where PointAt='17');
------------END POINT 17------------

------------------BEGIN POINT 18-----------------
if :pointat='18' or :pointat='All' or :pointat='A' then
	insert into #TBLTEMP
	select '18',T0."SrcObjAbs", T0."Code", 0,0,
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."VatSum" else T0."VatSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."VatSum" else T0."VatSum" end 
		End	
	) END VatSum,
	
	case when T0."CANCELED"='C' then
	(-1)* 
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	)
	else
	(
		case when T0."Category"='O' then
			case when "CrditDebit"='D' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		else
			case when "CrditDebit"='C' then (-1)*T0."BaseSum" else T0."BaseSum" end 
		End	
	) END "BaseSum", T0."VatPercent",T0."JrnlMemo",T0."SrcObjType"
	from "B1_VatView" T0 
	Join "OVTG" T1 on T0."Code"=T1."Code" 
	where "DocDate" between :fromdate and :todate
	and T1."ReportCode" in ('AJS');
end if;
update #GST03DATATBL set Point18=(select SUM(Balance+BaseSum) from #TBLTEMP where PointAt='18');
------------END POINT 18------------

------------------BEGIN POINT 19-----------------
insert into #TBLMISC
select "$rowid$", T0."U_MSICCode",T0."U_PERCENTAGE" from "@GST_MSIC" T0;

update #TBLMISC set ID=ID-:minrow-1;

while :i<=5--(select max(ID) from #TBLMISC) 
DO
	if :i=1 then
		update #GST03DATATBL set 
			Point19_1Value=point5b *(select Percentage from #TBLMISC where id=:i),
			Point19_1Per=(select Percentage from #TBLMISC where id=:i),
			Point19_1Code=(select Code from #TBLMISC where id=:i);
	end if;
	
	if :i=2 then
		update #GST03DATATBL set 
			Point19_2Value=point5b *(select Percentage from #TBLMISC where id=:i),
			Point19_2Per=(select Percentage from #TBLMISC where id=:i),
			Point19_2Code=(select Code from #TBLMISC where id=:i);
	end if;
	
	if :i=3 then
		update #GST03DATATBL set 
			Point19_3Value=point5b *(select Percentage from #TBLMISC where id=:i),
			Point19_3Per=(select Percentage from #TBLMISC where id=:i),
			Point19_3Code=(select Code from #TBLMISC where id=:i);
	end if;
	
	if :i=4 then
		update #GST03DATATBL set 
			Point19_4Value=point5b *(select Percentage from #TBLMISC where id=:i),
			Point19_4Per=(select Percentage from #TBLMISC where id=:i),
			Point19_4Code=(select Code from #TBLMISC where id=:i);
	end if;
	
	if :i=5 then
		update #GST03DATATBL set 
			Point19_5Value=point5b *(select Percentage from #TBLMISC where id=:i),
			Point19_5Per=(select Percentage from #TBLMISC where id=:i),
			Point19_5Code=(select Code from #TBLMISC where id=:i);
	end if;
	
	if :i=6 then
		update #GST03DATATBL set
			Point19_6Value =point5b *(select Percentage from #TBLMISC where id=:i), 
			Point19_6Per =(select Percentage from #TBLMISC where id=:i),
			Point19_6Code=(select Code from #TBLMISC where id=:i);
	end if;
	i:=:i+1;
end while;

------------------BEGIN POINT 19-----------------

--USING FOR TBL_PAYMENTCONTR

	--delete from "TBL_PAYMENTCONTRA";
	--Insert into "TBL_PAYMENTCONTRA" select * from #TBLTEMP where PointAt in ('5b','6b') ;

--USING FOR TBL_PAYMENTCONTR
GSTDetail=select * from #TBLTEMP;
GSTSummary=select * from #GST03DATATBL;

IF :pointat='All' then
	select * from #GST03DATATBL;
ELSE
	if:pointat='A' then
		select * from #TBLTEMP; 
	else
		select * from #TBLTEMP where PointAt = :pointat ;
	end if;
END IF;

drop table "#GST03DATATBL";
drop table "#TBLTEMP";
drop table "#TBLMISC";
 

END;

--Call SP_SAPB1ADDON_GST03('A','20151010','20151010','20151010');



