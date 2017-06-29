create  proc sp_SAPB1Addon_21DO
--**{0}**--
@FromDate datetime,
@Status nvarchar(1)
as

select 'N' ck,T0.DocEntry,T0.DocNum,T0.DocDate,T0.CardCode, T0.CardName,T0.NumAtCard,T0.DocTotal-T0.vatsum SubTol,
 T0.VatSum,T0.DocTotal, 
 isnull((select U_Value from [@GSTSETUP] where code='DOAct'),'') AccuralAct,
 case when isnull(T0.U_21Day,'N')='Y' then 'YES' else 'NO' End Status, T0.U_21DayJE,
 datediff(dd,T0.DocDate,@FromDate) Days,T2.Account TaxAct,
 sum(T1.VatSum*T1.OpenQty/T1.Quantity)  OutstandVatAmt
 from ODLN T0 with(nolock) 
 join DLN1 T1 with(nolock) on T0.DocEntry=T1.DocEntry
 join OVTG T2 with(nolock) on T2.Code=T1.VatGroup
 where  T0.CANCELED='N' 
 and T1.LineStatus='O' and T0.DocStatus='O' 
 and isnull(U_21Day,'N')= 'N'--(case when @Status='Y' then 'Y' when @Status='N' then 'N' else  isnull(U_21Day,'N') end)
 and datediff(dd,T0.DocDate,@FromDate)>=21
 and @Status in ('N','A')
 group by T0.DocEntry,T0.DocNum,T0.DocDate,T0.CardCode, T0.CardName,T0.NumAtCard,
 T0.VatSum,T0.DocTotal,T0.U_21DayJE,T2.Account ,T0.U_21Day
 having sum(T1.VatSum*T1.OpenQty/T1.Quantity)<>0

 UNION ALL

 select 'N' ck,T0.DocEntry,T0.DocNum,T0.DocDate,T0.CardCode, T0.CardName,T0.NumAtCard,T0.DocTotal-T0.vatsum SubTol,
 T0.VatSum,T0.DocTotal, 

 isnull((select U_Value from [@GSTSETUP] where code='DOAct'),'') AccuralAct,
 case when isnull(T0.U_21Day,'N')='Y' then 'YES' else 'NO' End Status, T0.U_21DayJE,
 datediff(dd,T0.DocDate,@FromDate) Days,'' TaxAct,
 0  OutstandVatAmt
 from ODLN T0 with(nolock) 
 --join DLN1 T1 with(nolock) on T0.DocEntry=T1.DocEntry
 --join OVTG T2 with(nolock) on T2.Code=T1.VatGroup
 join 
 (
 select distinct B.U_InvoiceEntry, A.TransId,B.Credit from OJDT A join JDT1 B on A.TransId=B.TransId
 Where isnull(A.U_21Day,'')='Y'  and isnull(B.U_InvoiceEntry,'') <>''
 ) T3 on T0.U_21DayJE=T3.TransId

 where  isnull(U_21Day,'N')= 'Y'
 and @Status in ('Y','A')
 --group by T0.DocEntry,T0.DocNum,T0.DocDate,T0.CardCode, T0.CardName,T0.NumAtCard,
 --T0.VatSum,T0.DocTotal,T0.U_21DayJE,T2.Account ,T0.U_21Day
