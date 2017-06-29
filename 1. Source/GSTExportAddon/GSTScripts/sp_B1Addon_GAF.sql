create proc sp_B1Addon_GAF
--**{0}**--
@FromDate datetime,
@ToDate datetime
as
--set @FromDate='2012-08-01'
--set @ToDate='2012-08-30'
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
SELECT CompnyName BusinessName,T0.FreeZoneNo BusinessRN, TaxIdNum GSTNumber, T1.F_RefDate PeriodStart,T1.T_RefDate PeriodEnd,
GETDATE() GAFCreationDate,'SAP_B1_9_PL4' ProductVersion, '20150101' GAFVersion
FROM OADM T0
join OACP T1 on T1.Year=2014

declare @companyxml nvarchar(max)
set @companyxml=isnull(( select * from @company FOR XML RAW ('Company'), ROOT ('Companies'), ELEMENTS),'')

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
T0.DocDate InvoiceDate,T1.NumAtCard InvoiceNumber,T1.ImportEnt ImportDeclarationNo, T0.LineNum LineNumber,
T0.Dscription ProductDescription, T0.LineTotal PurchaseValueMYR,
T0.VatSum GSTValueMYR, t0.TaxCode ,  T1.DocCur FCYCode, T0.TotalFrgn PurchaseFCY, T0.VatSumFrgn GSTFCY
from PCH1 T0 with(nolock)
join OPCH T1 with(nolock) on T0.DocEntry=T1.DocEntry
join OCRD T2 with(nolock) on T2.CardCode=T1.CardCode
Where T1.DocDate between @FromDate and @ToDate

declare @purchasexml nvarchar(max)
set @purchasexml=isnull((select * from @purchase FOR XML RAW ('Purchase'), ROOT ('Purchases'), ELEMENTS),'')

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
T0.Dscription ProductDescription, T0.LineTotal SupplyValueMYR,
T0.VatSum GSTValueMYR, t0.TaxCode , T2.Country, T1.DocCur FCYCode, T0.TotalFrgn SupplyFCY, T0.VatSumFrgn GSTFCY
from INV1 T0 with(nolock)
join OINV T1 with(nolock) on T0.DocEntry=T1.DocEntry
join OCRD T2 with(nolock) on T2.CardCode=T1.CardCode
Where T1.DocDate between @FromDate and @ToDate

declare @supplyxml nvarchar(max)
set @supplyxml=isnull((select * from @supply FOR XML RAW ('Supply'), ROOT ('Supplies'), ELEMENTS),'')

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
T1.LineMemo TransactionDescription, T3.CardName Name,
T0.transid TransactionID, T0.TransType SourceType, T0.BaseRef SourceDocumentID,
T1.Debit,T1.Credit, 
case when DebCred ='D' then T1.BalDueDeb else T1.BalDueCred end Balance
from OJDT T0 with(nolock)
join JDT1 T1 with(nolock) on T0.TransId=T1.transid
join OACT T2 with(nolock) on T2.AcctCode=T1.Account
left join OCRD T3 with(nolock) on T3.CardCode=T1.ShortName
where T0.RefDate between @FromDate and @ToDate

declare @ledgerxml nvarchar(max)
set @ledgerxml=isnull((select * from @ledger FOR XML RAW ('LedgerEntry'), ROOT ('Ledger'), ELEMENTS),'')


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
	(select COUNT(*) from OPCH where DocDate between @FromDate and @ToDate) TotalPurchaseCount,
	(select sum(LineTotal) from PCH1 T0 join OPCH T1 on T0.docEntry=T1.DocEntry where T1.DocDate between @FromDate	and @ToDate) TotalPurchaseAmount,
	(select sum(T0.VatSum) from PCH1 T0 join OPCH T1 on T0.docEntry=T1.DocEntry where T1.DocDate between @FromDate	and @ToDate) TotalPurchaseAmountGST,
	(select COUNT(*) from OINV where DocDate between @FromDate and @ToDate) TotalSupplyCount,
	(select sum(LineTotal) from INV1 T0 join OINV T1 on T0.docEntry=T1.DocEntry where T1.DocDate between @FromDate	and @ToDate) TotalSupplyAmount,
	(select sum(T0.VatSum) from INV1 T0 join OINV T1 on T0.docEntry=T1.DocEntry where T1.DocDate between @FromDate	and @ToDate) TotalSupplyAmountGST,
	(select COUNT(*) from OJDT where RefDate between @FromDate and @ToDate)	TotalLedgerCount,
	(select SUM(debit) from JDT1 T0 with(Nolock) join OJDT T1 with(nolock) on T0.TransID=T1.TransID where T1.RefDate between @FromDate and @ToDate) TotalLedgerDebit,
	(select SUM(credit) from JDT1 T0 with(Nolock) join OJDT T1 with(nolock) on T0.TransID=T1.TransID where T1.RefDate between @FromDate and @ToDate) TotalLedgerCredit,
	0 TotalLedgerBalance

declare @footerxml nvarchar(max)
set @footerxml=isnull((select * from @footer FOR XML RAW ('Footer'), ELEMENTS),'')

declare @GAFXml nvarchar(max)
set @GAFXml='<?xml version="1.0" encoding="utf-8"?> <GSTAuditFile>'
set @GAFXml=isnull(@GAFXml,'')+isnull(@companyxml,'')+isnull(@purchasexml,'')+isnull(@supplyxml,'')+isnull(@ledgerxml,'')+isnull(@footerxml,'')
set @GAFXml=@GAFXml+'</GSTAuditFile>'

--select * from @company
--select * from @purchase
--select * from @Supply
--select * from @ledger
--select * from @footer

select @GAFXml xmlGAF,LEN(@GAFXml) xmllen