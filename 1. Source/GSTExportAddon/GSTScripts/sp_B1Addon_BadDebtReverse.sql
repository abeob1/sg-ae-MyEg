create proc sp_B1Addon_BadDebtReverse 
--**{0}**--
@InvoiceNum as nvarchar(100),
@PaidAmount as numeric(19,6)
as

select 
case when T1.DocStatus='C' then --Last Payment
	Debit -
	isnull((
		select sum(isnull(Credit,0)) from jdt1 where U_InvoiceEntry=@InvoiceNum
		and VatGroup=(select U_value from [@GSTSETUP] where code='AROutTaxCode')
		and Credit>0
	),0)
else @PaidAmount*6/106  --Partial Payment
end Debit , 
T0.BaseSum,
(select U_value from [@GSTSETUP] where code='ACT3') DrAct, 
(select U_value from [@GSTSETUP] where code='AROutTaxCode') TaxCode, 
(Select Account from OVTG where Code= (select U_value from [@GSTSETUP] 
										where code='AROutTaxCode')) CrAct 
from JDT1 T0 with(nolock)
join OINV T1 with(nolock) on T1.DocNum=@InvoiceNum
where Debit>0 and U_InvoiceEntry=@InvoiceNum

and VatGroup=(select U_value from [@GSTSETUP] where code='ARInTaxCode')

