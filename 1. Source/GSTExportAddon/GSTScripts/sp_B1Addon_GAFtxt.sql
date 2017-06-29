create proc sp_B1Addon_GAFtxt
--**{0}**--
@FromDate datetime,
@ToDate datetime
as
--set @FromDate='2012-08-01'
--set @ToDate='2012-08-30'
--sp_B1Addon_GAFtxt '2015-08-01','2015-09-01'
---------------------------COMPANY---------------------
declare @company table(
	BusinessName nvarchar(500),
	BusinessRN nvarchar(500),
	GSTNumber nvarchar(500),
	PeriodStart date,
	PeriodEnd date,
	GAFCreationDate date,
	ProductVersion nvarchar(500),
	GAFVersion nvarchar(500)
)

Insert into @company
SELECT top(1) CompnyName BusinessName,isnull(T0.FreeZoneNo,'') BusinessRN, TaxIdNum GSTNumber, T1.F_RefDate PeriodStart,T1.T_RefDate PeriodEnd,
GETDATE() GAFCreationDate,'SAP_B1_9_PL4' ProductVersion, '20150101' GAFVersion
FROM OADM T0
join OACP T1 on T1.Year=Year(@FromDate)

declare @companyxml nvarchar(max)
select @companyxml='C|'+ BusinessName+'|'+ BusinessRN+'|'+GSTNumber+
	'|'+convert(nvarchar(10),PeriodStart,103)+
	'|'+convert(nvarchar(10),PeriodEnd,103)+
	'|'+convert(nvarchar(10),GAFCreationDate,103)+
	'|'+ProductVersion+'|'+GAFVersion+CHAR(13)+CHAR(10)
from @company

---------------------------Purchase---------------------
declare @purchase table(
	SupplierName  nvarchar(500),
	SupplierBRN  nvarchar(500),
	InvoiceDate date,
	InvoiceNumber  nvarchar(500),
	ImportDeclarationNo  nvarchar(500),
	LineNumber int,
	ProductDescription  nvarchar(500),
	PurchaseValueMYR numeric(18,6),
	GSTValueMYR numeric(18,6),
	TaxCode  nvarchar(500),
	FCYCode  nvarchar(500),
	PurchaseFCY numeric(18,6),
	GSTFCY numeric(18,6)
)

insert into @purchase
select T1.CardName SupplierName, T2.LicTradNum SupplierBRN,
T0.DocDate InvoiceDate,T1.NumAtCard InvoiceNumber,isnull(T1.ImportEnt,'') ImportDeclarationNo, T0.LineNum LineNumber,
T0.Dscription ProductDescription, T0.LineTotal PurchaseValueMYR,
T0.VatSum GSTValueMYR, isnull(t0.TaxCode,'') ,  T1.DocCur FCYCode, T0.TotalFrgn PurchaseFCY, T0.VatSumFrgn GSTFCY
from PCH1 T0 with(nolock)
join OPCH T1 with(nolock) on T0.DocEntry=T1.DocEntry
join OCRD T2 with(nolock) on T2.CardCode=T1.CardCode
Where T1.DocDate between @FromDate and @ToDate

declare @purchasexml nvarchar(max)
set @purchasexml=''
select @purchasexml=@purchasexml+
'P|'+SupplierName+'|'+ SupplierBRN+'|'+convert(nvarchar(10),InvoiceDate,103)+'|'+InvoiceNumber+'|'+ImportDeclarationNo+'|'+
convert(nvarchar(100),LineNumber)+'|'+ProductDescription+'|'+convert(nvarchar(100),PurchaseValueMYR)+'|'+
convert(nvarchar(100),GSTValueMYR)+'|'+TaxCode+'|'+
FCYCode+'|'+convert(nvarchar(100),PurchaseFCY)+'|'+convert(nvarchar(100),GSTFCY)+CHAR(13)+CHAR(10)
from @purchase

---------------------------Supply---------------------
declare @Supply table (
	CustomerName  nvarchar(500),
	CustomerBRN  nvarchar(500),
	InvoiceDate date,
	InvoiceNumber  nvarchar(500),
	LineNumber int,
	ProductDescription  nvarchar(500),
	SupplyValueMYR numeric(18,6),
	GSTValueMYR numeric(18,6),
	TaxCode  nvarchar(500),
	Country  nvarchar(500),
	FCYCode  nvarchar(500),
	SupplyFCY numeric(18,6),
	GSTFCY numeric(18,6)
)

insert into @Supply
select T1.CardName CustomerName, T2.LicTradNum CustomerBRN,
T0.DocDate InvoiceDate,T1.NumAtCard InvoiceNumber, T0.LineNum LineNumber,
isnull(T0.Dscription,'') ProductDescription, T0.LineTotal SupplyValueMYR,
T0.VatSum GSTValueMYR, isnull(t0.TaxCode,'') , T2.Country, T1.DocCur FCYCode, T0.TotalFrgn SupplyFCY, T0.VatSumFrgn GSTFCY
from INV1 T0 with(nolock)
join OINV T1 with(nolock) on T0.DocEntry=T1.DocEntry
join OCRD T2 with(nolock) on T2.CardCode=T1.CardCode
Where T1.DocDate between @FromDate and @ToDate

declare @supplyxml nvarchar(max)
set @supplyxml=''
select @supplyxml=@supplyxml+'S|'+CustomerName+'|'+CustomerBRN+'|'+convert(nvarchar(10),InvoiceDate,103)+'|'+InvoiceNumber+
'|'+convert(nvarchar(100),LineNumber)+'|'+ProductDescription+'|'+convert(nvarchar(100),SupplyValueMYR)+'|'+
convert(nvarchar(100),GSTValueMYR)+'|'+TaxCode+'|'+Country+'|'+
FCYCode+'|'+convert(nvarchar(100),SupplyFCY)+'|'+convert(nvarchar(100),GSTFCY)+CHAR(13)+CHAR(10)
from @Supply



---------------------------Ledger---------------------
declare @ledger table (
	TransactionDate date,
	AccountID nvarchar(500),
	AccountName nvarchar(500),
	TransactionDescription nvarchar(500),
	Name nvarchar(500),
	TransactionID numeric(18,0),
	SourceType nvarchar(500),
	SourceDocumentID nvarchar(500),
	Debit numeric(18,6),
	Credit numeric(18,6),
	Balance numeric(18,6)
	
)
insert into @ledger
select T0.RefDate TransactionDate,T1.Account AccountID,T2.AcctName AccountName,
T1.LineMemo TransactionDescription, isnull(T3.CardName,'') Name,
T0.transid TransactionID, T0.TransType SourceType, T0.BaseRef SourceDocumentID,
T1.Debit,T1.Credit, 
case when DebCred ='D' then T1.BalDueDeb else T1.BalDueCred end Balance
from OJDT T0 with(nolock)
join JDT1 T1 with(nolock) on T0.TransId=T1.transid
join OACT T2 with(nolock) on T2.AcctCode=T1.Account
left join OCRD T3 with(nolock) on T3.CardCode=T1.ShortName
where T0.RefDate between @FromDate and @ToDate

declare @ledgerxml nvarchar(max)
set @ledgerxml=''
select @ledgerxml=@ledgerxml+'L|'+convert(nvarchar(10),TransactionDate,103)+'|'+AccountID+'|'+AccountName+'|'+TransactionDescription+'|'+
Name+'|'+convert(nvarchar(100),TransactionID)+'|'+SourceDocumentID+'|'+SourceType+'|'+
convert(nvarchar(100),Debit)+'|'+convert(nvarchar(100),Credit)+'|'+convert(nvarchar(100),Balance)+CHAR(13)+CHAR(10)
from @ledger



---------------------------footer---------------------
--------------optimize this by using temp table---------
declare @footer table(
TotalPurchaseCount numeric(18,6),
TotalPurchaseAmount numeric(18,6),
TotalPurchaseAmountGST numeric(18,6),
TotalSupplyCount numeric(18,6),
TotalSupplyAmount numeric(18,6),
TotalSupplyAmountGST numeric(18,6),
TotalLedgerCount numeric(18,6),
TotalLedgerDebit numeric(18,6),
TotalLedgerCredit numeric(18,6),
TotalLedgerBalance numeric(18,6)
)
insert into @footer
select 
	isnull((select COUNT(*) from @purchase),0) TotalPurchaseCount,
	isnull((select sum(PurchaseValueMYR) from @purchase),0) TotalPurchaseAmount,
	isnull((select sum(GSTValueMYR) from @purchase),0) TotalPurchaseAmountGST,
	isnull((select COUNT(*) from @Supply),0) TotalSupplyCount,
	isnull((select sum(SupplyValueMYR) from @Supply),0) TotalSupplyAmount,
	isnull((select sum(GSTValueMYR) from @Supply),0) TotalSupplyAmountGST,
	isnull((select COUNT(*) from @ledger),0)	TotalLedgerCount,
	isnull((select SUM(Debit) from @ledger),0) TotalLedgerDebit,
	isnull((select SUM(credit) from @ledger),0) TotalLedgerCredit,
	isnull((select SUM(Balance) from @ledger),0) TotalLedgerBalance

declare @footerxml nvarchar(max)
set @footerxml=''--isnull((select * from @footer FOR XML RAW ('Footer'), ELEMENTS),'')
select @footerxml=@footerxml+'F|'
+convert(nvarchar(100),TotalPurchaseCount)+'|'+convert(nvarchar(100),TotalPurchaseAmount)+'|'+convert(nvarchar(100),TotalPurchaseAmountGST)+'|'
+convert(nvarchar(100),TotalSupplyCount)+'|'+convert(nvarchar(100),TotalSupplyAmount)+'|'+convert(nvarchar(100),TotalSupplyAmountGST)+'|'
+convert(nvarchar(100),TotalLedgerCount)+'|'+convert(nvarchar(100),TotalLedgerDebit)+'|'+convert(nvarchar(100),TotalLedgerCredit)+'|'+convert(nvarchar(100),TotalLedgerBalance)+CHAR(13)+CHAR(10)
from @footer

declare @GAFXml nvarchar(max)
set @GAFXml=isnull(@companyxml,'')+isnull(@purchasexml,'')+isnull(@supplyxml,'')+isnull(@ledgerxml,'')+isnull(@footerxml,'')

--select * from @company
--select * from @purchase
--select * from @Supply
--select * from @ledger
--select * from @footer

select @GAFXml xmlGAF,LEN(@GAFXml) xmllen