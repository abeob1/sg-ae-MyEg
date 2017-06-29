--**{0}**--
{1} function dbo.uf_GetTaxBalance(@DocEntry numeric(18,0))
returns numeric(18,6)
begin
	declare @TaxBalance1 numeric(18,6)
	declare @TaxBalance2 numeric(18,6)
	select --T1.doctotal-T1.paidSum balance, (T1.doctotal-T1.paidSum)/T1.doctotal*T0.Gtotal LineBalance, T0.vatsum, T0.GTotal,
	@TaxBalance1=sum(
		case when T0.GTotal=0 then 0 
		else (T1.doctotal-T1.paidSum)/T1.doctotal*T0.Gtotal/T0.GTotal*T0.VatSum 
		End)
	from INV1 T0 with(nolock) 
	join OINV T1 with(nolock) on T0.DocEntry=T1.DocEntry
	where T0.docentry=@DocEntry

	select --T1.doctotal-T1.paidSum balance, (T1.doctotal-T1.paidSum)/T1.doctotal*T0.Gtotal LineBalance, T0.vatsum, T0.GTotal,
	@TaxBalance2=sum(
		case when T0.LineTotal=0 then 0 
		else (T1.doctotal-T1.paidSum)/T1.doctotal*T0.LineTotal/T0.LineTotal*T0.VatSum 
		End)
	from INV3 T0 with(nolock) 
	join OINV T1 with(nolock) on T0.DocEntry=T1.DocEntry
	where T0.docentry=@DocEntry

	return isnull(@TaxBalance1,0)+isnull(@TaxBalance2,0)
End
