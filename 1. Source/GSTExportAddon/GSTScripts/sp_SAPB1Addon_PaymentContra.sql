create proc [dbo].[sp_SAPB1Addon_PaymentContra]
--**{0}**--
@fromdate datetime,
@todate datetime
as
--[sp_SAPB1Addon_PaymentContra] '20010101','20200101'

declare @tbl table(PointAt nvarchar(10), TransID numeric(18,0), 
VatGroup nvarchar(20), Debit numeric(19,6), Credit numeric(19,6), Balance numeric(19,6),
BaseSum numeric(19,6), VatRat numeric(19,6), Memo nvarchar(100), transtype nvarchar(10))

insert into @tbl
exec [sp_SAPB1Addon_GST03] 'A',@fromdate,@Todate,@todate


select case when PointAt='6b' then 'Input' else 'Output' end Category,

case when PointAt='5b' then 
	(select sum(Balance) from @tbl where PointAt='5b')
else 
	(select sum(Balance) from @tbl where PointAt='6b')
 end TotalBalance,VatGroup,
convert(nvarchar(10),T0.TransID) TransID,Debit,Credit,Balance, BaseSum,VatRat, Memo,transtype, T1.Account TaxAccount, 
(select U_Value from [@GSTSETUP] where code='ContraAct') ContraAccount
from @tbl T0
join OVTG T1 on T0.VatGroup=T1.Code
left join
(
	select distinct T0.TransID,T1.U_InvoiceEntry from OJDT T0 with(nolock) 
	join JDT1 T1 with(nolock) on T0.TransID=T1.TransID
	Where isnull(T0.U_ContraPayment,'')='Y'
) T2 on T2.U_InvoiceEntry=T0.TransID
where PointAt in ('5b','6b') and Balance<>0 and T2.TransId is null