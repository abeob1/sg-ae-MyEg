create function uf_GetTaxBalance
--**{0}**--
(
DocEntry decimal
)
returns TaxBalance decimal
LANGUAGE SQLSCRIPT READS SQL DATA AS
begin
	--declare TaxBalance decimal:=0;
	
	select --T1.doctotal-T1.paidSum balance, (T1.doctotal-T1.paidSum)/T1.doctotal*T0.Gtotal LineBalance, T0.vatsum, T0.GTotal,
		sum(
		case when T0."GTotal"=0 then 0 
		else (T1."DocTotal"-T1."PaidSum")/T1."DocTotal"*T0."GTotal"/T0."GTotal"*T0."VatSum" 
		End) INTO TaxBalance
	from "INV1" T0  
	join "OINV" T1 on T0."DocEntry"=T1."DocEntry"
	where T0."DocEntry"=:DocEntry;
End;