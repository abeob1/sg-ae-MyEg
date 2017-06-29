create PROCEDURE sp_B1Addon_BadDebtReverse
--**{0}**--
(
	IN InvoiceNum nvarchar(100),
	IN PaidAmount decimal
)
LANGUAGE SQLSCRIPT
AS
BEGIN
	select 
		case when T1."DocStatus"='C' then --Last Payment
		"Debit" -
		ifnull((
			select sum(ifnull("Credit",0)) from "JDT1" where "U_InvoiceEntry"=:InvoiceNum
			and "VatGroup"=(select "U_Value" from "@GSTSETUP" where "Code"='AROutTaxCode')
			and "Credit">0
		),0)
		else :PaidAmount*6/106  --Partial Payment
		end as "Debit" , 
	(select "U_Value" from "@GSTSETUP" where "Code"='ACT3') "DrAct", 
	(select "U_Value" from "@GSTSETUP" where "Code"='AROutTaxCode') as "TaxCode", 
	(Select "Account" from "OVTG" where "Code"= (select "U_Value" from "@GSTSETUP"
											where "Code"='AROutTaxCode')) as "CrAct" 
	from "JDT1" T0 
	join "OINV" T1 on T1."DocNum"=:InvoiceNum
	JOIN "OJDT" T2 on T0."TransId"=T2."TransId"
	where T2."U_BadDebt"='Y' and  "Debit">0 and "U_InvoiceEntry"=:InvoiceNum
	
	and "VatGroup"=(select "U_Value" from "@GSTSETUP" where "Code"='ARInTaxCode');
END;

