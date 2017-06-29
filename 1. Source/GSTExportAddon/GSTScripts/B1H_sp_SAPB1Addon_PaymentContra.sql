drop PROCEDURE SP_SAPB1ADDON_PAYMENTCONTRA;
create PROCEDURE SP_SAPB1ADDON_PAYMENTCONTRA
--**{0}**--
(
	IN fromdate Date,
	IN todate Date
)
LANGUAGE SQLSCRIPT
AS
BEGIN
	
	create local temporary table #TBLTEMP1(
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
	

----------------BEGIN POINT 5b-----------------
	insert into #TBLTEMP1
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
	left join
	(
		select T0."TransId",T1."U_InvoiceEntry" , T1."Account"
		from "OJDT" T0 
		join "JDT1" T1 on T0."TransId"=T1."TransId"
		Where ifnull(T0."U_ContraPayment",'')='Y'
	) T2 on T2."U_InvoiceEntry"=T0."SrcObjAbs" and T1."Account"=T2."Account"
	where "DocDate" between :fromdate and :todate
	and T1."ReportCode" in ('SR','DS','AJS')
	and T2."TransId" is null;
----------------END POINT 5b-----------------

 
----------------BEGIN POINT 6b-----------------
	insert into #TBLTEMP1
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
	left join
	(
		select T0."TransId",T1."U_InvoiceEntry" , T1."Account"
		from "OJDT" T0 
		join "JDT1" T1 on T0."TransId"=T1."TransId"
		Where ifnull(T0."U_ContraPayment",'')='Y'
	) T2 on T2."U_InvoiceEntry"=T0."SrcObjAbs" and T1."Account"=T2."Account"
	Where "DocDate" between :fromdate and :todate
	and T1."ReportCode" in ('TX', 'IM','TX-E43', 'TX-RE','AJP','RC') 
	and T2."TransId" is null;
----------------END POINT 6b-----------------

	Select case when "POINTAT"='6b' then 'Input' else 'Output' end "Category",
		case when "POINTAT"='5b' then 
			(select sum("BALANCE") from #TBLTEMP1 where "POINTAT"='5b')
		else 
			(select sum("BALANCE") from #TBLTEMP1 where "POINTAT"='6b')
			end "TotalBalance","VATGROUP",
		CAST(T0."TRANSID" as nvarchar(10)) "TransId","DEBIT","CREDIT","BALANCE", "BASESUM","VATRAT", "MEMO","TRANSTYPE", T1."Account" "TaxAccount", 
		(select "U_Value" from "@GSTSETUP" where "Code"='ContraAct') "ContraAccount"
	from #TBLTEMP1 T0
	join "OVTG" T1 on T0."VATGROUP"=T1."Code"
	where "BALANCE"<>0 ;


	drop table "#TBLTEMP1";

END;
