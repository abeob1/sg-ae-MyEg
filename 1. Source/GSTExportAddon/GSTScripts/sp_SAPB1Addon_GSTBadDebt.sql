create proc sp_SAPB1Addon_GSTBadDebt
--**{0}**--
@Date datetime,
@Debitor nvarchar(100),
@Month int,
@ClaimAmt nvarchar(1),
@Status nvarchar(1)
as
declare @tbl table (ck nvarchar(1),docentry numeric(19,0), DocNum nvarchar(100), Docdate datetime, NumAtCard nvarchar(100),
CardCode nvarchar(100), CardName nvarchar(100), 
SubTol numeric(19,6), 
VatSum numeric(19,6), 
DocTotal numeric(19,6), 
paidsum numeric(19,6), 
Balance numeric(19,6), 
Status nvarchar(100),
U_BadDebtJE nvarchar(100),
CrAct nvarchar(100),
DrAct nvarchar(100), 
 
TaxCode nvarchar(100), 
BaseAmount  numeric(19,6),
Amount numeric(19,6),
BadDebtJENo nvarchar(100)
)

insert into @tbl
select 'N' ck, docentry,DocNum,docdate, NumAtCard, cardcode,CardName,DocTotal-VatSum SubTol,  
VatSum,DocTotal, paidsum , DocTotal-PaidSum Balance,
case when isnull(T0.U_BadDebt,'N')='Y' then 'YES' else 'NO' End Status,T0.U_BadDebtJE,
(select U_Value from [@GSTSETUP] where code='ACT2') CrAct,
(
select Account from ovtg where code=
	(select U_Value from [@GSTSETUP] where code='ARInTaxCode')
) DrAct,

(select U_Value from [@GSTSETUP] where code='ARInTaxCode') TaxCode,

case when T1.TransId is null then
	case when PaidSum=0 then 
		(
			select sum(isnull(A.BaseSum,0)) from jdt1 A with(nolock) 
			where a.TransId=T0.TransId
			and a.TransType=T0.ObjType
		)
	else 0 end 
else
	(select sum(isnull(A.BaseSum,0))  from JDT1 A where A.TransId=T1.TransId)
end BaseAmount,

dbo.uf_GetTaxBalance(T0.DocEntry) Amount, T1.Number BadDebtJENo

from oinv T0 with(nolock) 
left join OJDT T1 with(nolock) on T0.U_BadDebtJE=T1.TransId

where canceled='N' and DocStatus='O'
and ((datediff(month,DocDate, getdate())>@Month) OR @month=0)
and ((cardCode=@Debitor) OR @Debitor='')
and (isnull(T0.U_BadDebt,'N')=@Status or @Status='A')

select * from @tbl  where (Amount>0 or @ClaimAmt='N')
