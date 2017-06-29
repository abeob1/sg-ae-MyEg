create PROCEDURE sp_SAPB1Addon_21DO
--**{0}**--
(
	IN FromDate date,
	IN Status nvarchar(1)
)
LANGUAGE SQLSCRIPT
AS
BEGIN
	select 'N' as "ck",T0."DocEntry",T0."DocNum",T0."DocDate",T0."CardCode", 
	 T0."CardName",T0."NumAtCard",T0."DocTotal"-T0."VatSum" as "SubTol",
	 T0."VatSum",T0."DocTotal", 
	 ifnull((select "U_Value" from "@GSTSETUP" where "Code"='DOAct'),'') as "AccuralAct",
	 case when ifnull(T0."U_21Day",'N')='Y' then 'YES' else 'NO' End as "Status", T0."U_21DayJE",
	 Days_Between(T0."DocDate",:FromDate) Days,T2."Account" as "TaxAct",
	 sum(T1."VatSum"*T1."OpenQty"/T1."Quantity") as "OutstandVatAmt"
	 from "ODLN" T0 
	 join "DLN1" T1 on T0."DocEntry"=T1."DocEntry"
	 join "OVTG" T2  on T2."Code"=T1."VatGroup"
	 where  T0."CANCELED"='N' 
	 and T1."LineStatus"='O' and T0."DocStatus"='O' 
	 and ifnull("U_21Day",'N')= 'N'--(case when @Status='Y' then 'Y' when @Status='N' then 'N' else  isnull(U_21Day,'N') end)
	 and  Days_Between(T0."DocDate",:FromDate)>=21
	 and :Status in ('N','A')
	 group by T0."DocEntry",T0."DocNum",T0."DocDate",T0."CardCode", T0."CardName",T0."NumAtCard",
	 T0."VatSum",T0."DocTotal",T0."U_21DayJE",T2."Account" ,T0."U_21Day"
	 having sum(T1."VatSum"*T1."OpenQty"/T1."Quantity")<>0
	
	 UNION ALL
	
	 select 'N',T0."DocEntry",T0."DocNum",T0."DocDate",T0."CardCode", T0."CardName",T0."NumAtCard",T0."DocTotal"-T0."VatSum",
	 T0."VatSum",T0."DocTotal", 
	 ifnull((select "U_Value" from "@GSTSETUP" where "Code"='DOAct'),'') ,
	 case when ifnull(T0."U_21Day",'N')='Y' then 'YES' else 'NO' End , T0."U_21DayJE",
	 Days_Between(T0."DocDate",:FromDate),'' , 0  
	 from "ODLN" T0 
	 --join DLN1 T1 with(nolock) on T0.DocEntry=T1.DocEntry
	 --join OVTG T2 with(nolock) on T2.Code=T1.VatGroup
	 join 
	 (
		 select distinct B."U_InvoiceEntry", A."TransId",B."Credit" from "OJDT" A join "JDT1" B on A."TransId"=B."TransId"
		 Where ifnull(A."U_21Day",'')='Y'  and ifnull(B."U_InvoiceEntry",'') <>''
		 ) T3 on T0."U_21DayJE"=T3."TransId"
	
	 where  ifnull("U_21Day",'N')= 'Y'
	 and :Status in ('Y','A');
	 --group by T0.DocEntry,T0.DocNum,T0.DocDate,T0.CardCode, T0.CardName,T0.NumAtCard,
	 --T0.VatSum,T0.DocTotal,T0.U_21DayJE,T2.Account ,T0.U_21Day
END;