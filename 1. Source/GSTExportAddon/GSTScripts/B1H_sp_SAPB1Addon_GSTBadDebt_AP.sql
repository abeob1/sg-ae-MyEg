create PROCEDURE sp_SAPB1Addon_GSTBadDebt_AP
--**{0}**--
(
	IN Date Date,
	IN Debitor nvarchar(100),
	IN Month int,
	IN ClaimAmt nvarchar(1),
	IN Status nvarchar(1)
)
LANGUAGE SQLSCRIPT
AS
BEGIN

create local temporary table #TBL
	( 
		ck nvarchar(1),
		docentry decimal, 
		DocNum nvarchar(100), 
		Docdate Datetime, 
		NumAtCard nvarchar(100),
		CardCode nvarchar(100), 
		CardName nvarchar(100), 
		SubTol decimal, 
		VatSum decimal, 
		DocTotal decimal, 
		paidsum decimal, 
		Balance decimal, 
		Status nvarchar(100),
		U_BadDebtJE nvarchar(100),
		CrAct nvarchar(100),
		DrAct nvarchar(100), 
		TaxCode nvarchar(100), 
		Amount decimal,
		BadDebtJENo nvarchar(100)
	);

	insert into #TBL	
	select 'N', T0."DocEntry",T0."DocNum",T0."DocDate", T0."NumAtCard", T0."CardCode",
	T0."CardName",T0."DocTotal"-T0."VatSum" as "SubTol",  
	T0."VatSum",T0."DocTotal", T0."PaidSum" , T0."DocTotal"-T0."PaidSum" as "Balance",
	case when T0."U_BadDebt"='Y' then 'YES' else 'NO' End "Status","U_BadDebtJE",
	(select "U_Value" from "@GSTSETUP" where "Code"='ACT5') as "CrAct",
	(
	select "Account" from "OVTG" where "Code"=
		(select "U_Value" from "@GSTSETUP" where "Code"='APOutTaxCode')
	) as "DrAct",
	
	(select "U_Value" from "@GSTSETUP" where "Code"='APOutTaxCode') as "TaxCode",
	
	T2."Amount", T1."Number" as "BadDebtJENo"
	
	from "OPCH" T0 
	left join "OJDT" T1  on T0."U_BadDebtJE"=T1."TransId"
	join
	(
		select --T1.doctotal-T1.paidSum balance, (T1.doctotal-T1.paidSum)/T1.doctotal*T0.Gtotal LineBalance, T0.vatsum, T0.GTotal,
		T0."DocEntry",sum(
		case when T0."GTotal"=0 then 0 
		else (T1."DocTotal"-T1."PaidSum")/T1."DocTotal"*T0."GTotal"/T0."GTotal"*T0."VatSum" 
		End) as "Amount"
		from "PCH1" T0  
		join "OPCH" T1 on T0."DocEntry"=T1."DocEntry"
		Group by T0."DocEntry"
	) T2 on T0."DocEntry"=T2."DocEntry"
	where "CANCELED"='N' and "DocStatus"='O'
	and ( 
		(Year(:Date)-Year("DocDate"))*12 + Month(:Date)-Month("DocDate")>:Month 
		OR :month=0
		or :Status='Y'
		)
	and (T0."CardCode"=:Debitor OR :Debitor='')
	and (ifnull(T0."U_BadDebt",'N')=:Status or :Status='A');
	
	select * from #TBL T0 where T0.Amount >0 or :ClaimAmt='N';
	drop table "#TBL";
END;