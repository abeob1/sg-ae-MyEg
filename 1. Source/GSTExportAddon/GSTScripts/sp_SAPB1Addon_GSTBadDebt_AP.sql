create proc sp_SAPB1Addon_GSTBadDebt_AP
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
Amount numeric(19,6),
BadDebtJENo nvarchar(100)
)

insert into @tbl
select 'N' ck, docentry,DocNum,docdate, NumAtCard, cardcode,CardName,DocTotal-VatSum SubTol,  
VatSum,DocTotal, paidsum , DocTotal-PaidSum Balance,
case when isnull(T0.U_BadDebt,'N')='Y' then 'YES' else 'NO' End Status,T0.U_BadDebtJE,
(select U_Value from [@GSTSETUP] where code='ACT5') CrAct,
(
select Account from ovtg where code=
	(select U_Value from [@GSTSETUP] where code='APOutTaxCode')
) DrAct,

(select U_Value from [@GSTSETUP] where code='APOutTaxCode') TaxCode,

dbo.uf_GetTaxBalance_AP(T0.DocEntry) Amount, T1.Number BadDebtJENo

from OPCH T0 with(nolock) 
left join OJDT T1 with(nolock) on T0.U_BadDebtJE=T1.TransId

where canceled='N' and DocStatus='O'
and ((datediff(month,DocDate, getdate())>@Month) OR @month=0)
and ((cardCode=@Debitor) OR @Debitor='')
and (isnull(T0.U_BadDebt,'N')=@Status or @Status='A')

select * from @tbl  where (Amount>0 or @ClaimAmt='N')
