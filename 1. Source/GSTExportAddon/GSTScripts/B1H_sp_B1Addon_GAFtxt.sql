create PROCEDURE  sp_B1Addon_GAFtxt
--**{0}**--
(
	IN FromDate date,
	IN ToDate date
)
LANGUAGE SQLSCRIPT
AS
BEGIN
---------------------------COMPANY---------------------
	declare companyxml nvarchar(5000);
	declare purchasexml nvarchar(5000):='';
	declare ledgerxml nvarchar(5000):='';
	declare supplyxml nvarchar(5000):='';
	declare GAFXml nvarchar(5000):='';
	declare footerxml nvarchar(5000):='';
	
	create local temporary table #COMPANY
	( 
		BusinessName nvarchar(500),
		BusinessRN nvarchar(500),
		GSTNumber nvarchar(500),
		PeriodStart date,
		PeriodEnd date,
		GAFCreationDate date,
		ProductVersion nvarchar(500),
		GAFVersion nvarchar(500)
	);
	create local temporary table #SUPPLY 
	(
		CustomerName  nvarchar(500),
		CustomerBRN  nvarchar(500),
		InvoiceDate date,
		InvoiceNumber  nvarchar(500),
		LineNumber int,
		ProductDescription  nvarchar(500),
		SupplyValueMYR decimal,
		GSTValueMYR decimal,
		TaxCode  nvarchar(500),
		Country  nvarchar(500),
		FCYCode  nvarchar(500),
		SupplyFCY decimal,
		GSTFCY decimal
	);
	create local temporary table #PURCHASE 
	(
		SupplierName  nvarchar(500),
		SupplierBRN  nvarchar(500),
		InvoiceDate date,
		InvoiceNumber  nvarchar(500),
		ImportDeclarationNo  nvarchar(500),
		LineNumber int,
		ProductDescription  nvarchar(500),
		PurchaseValueMYR decimal,
		GSTValueMYR decimal,
		TaxCode  nvarchar(500),
		FCYCode  nvarchar(500),
		PurchaseFCY decimal,
		GSTFCY decimal
	);
	create local temporary table #LEDGER 
	(
		TransactionDate date,
		AccountID nvarchar(500),
		AccountName nvarchar(500),
		TransactionDescription nvarchar(500),
		Name nvarchar(500),
		TransactionID decimal,
		SourceType nvarchar(500),
		SourceDocumentID nvarchar(500),
		Debit decimal,
		Credit decimal,
		Balance decimal	
	);
	create local temporary table #FOOTER
	(
		TotalPurchaseCount decimal,
		TotalPurchaseAmount decimal,
		TotalPurchaseAmountGST decimal,
		TotalSupplyCount decimal,
		TotalSupplyAmount decimal,
		TotalSupplyAmountGST decimal,
		TotalLedgerCount decimal,
		TotalLedgerDebit decimal,
		TotalLedgerCredit decimal,
		TotalLedgerBalance decimal
	);
		
		
	Insert into #COMPANY
	SELECT top 1 "CompnyName" as "BusinessName",T0."FreeZoneNo" as "BusinessRN", 
	"TaxIdNum" as "GSTNumber", T1."F_RefDate" as "PeriodStart",T1."T_RefDate" as "PeriodEnd",
	CURRENT_DATE as "GAFCreationDate",'SAP_B1_9_PL4' as "ProductVersion", '20150101' as "GAFVersion"
	FROM "OADM" T0
	join "OACP" T1 on T1."Year"=Year(:FromDate);

	
	select 'C' ||
		char(124) || CAST(BusinessName as nvarchar(10)) ||
		char(124) || CAST(BusinessRN as nvarchar(10)) ||
		char(124) || CAST(GSTNumber as nvarchar(10)) ||
		char(124) || CAST(PeriodStart as nvarchar(10)) ||
		char(124) || CAST(PeriodEnd as nvarchar(10))||
		char(124) || CAST(GAFCreationDate as nvarchar(10)) ||
		char(124) || CAST(ProductVersion as nvarchar(10)) ||
		char(124) || CAST(GAFVersion as nvarchar(10)) || CHAR(13) || CHAR(10)
		into companyxml
	from #COMPANY;
	select companyxml from DUMMY;
	
-----------------------------Purchase---------------------
	

	--insert into #PURCHASE
	--select T1."CardName", T2."LicTradNum",
	--T0."DocDate",T1."NumAtCard",T1."ImportEnt", T0."LineNum",
	--T0."Dscription", T0."LineTotal",
	--T0."VatSum", T0."TaxCode", T1."DocCur", T0."TotalFrgn", T0."VatSumFrgn"
	--from "PCH1" T0  
	--join "OPCH" T1   on T0."DocEntry"=T1."DocEntry"
	--join "OCRD" T2   on T2."CardCode"=T1."CardCode"
	--Where T1."DocDate" between :FromDate and :ToDate;

	
	
--	select 'P' ||
--	char(124) || SupplierName ||
--	char(124) || SupplierBRN ||
--	char(124) || cast(InvoiceDate as nvarchar(10)) ||
--	char(124) || InvoiceNumber ||
--	char(124) || ImportDeclarationNo ||
--	char(124) || cast(LineNumber as nvarchar(100)) || 
--	char(124) || ProductDescription || 
--	char(124) || cast(PurchaseValueMYR as nvarchar(100)) ||
--	char(124) || cast(GSTValueMYR as nvarchar(100)) ||
--	char(124) || TaxCode ||
--	char(124) || FCYCode ||
--	char(124) || cast(PurchaseFCY as nvarchar(100)) ||
--	char(124) || cast(GSTFCY as nvarchar(100))  || CHAR(13)|| CHAR(10)
--	into purchasexml
--	from #PURCHASE;

-----------------------------Supply---------------------
	
--	insert into #SUPPLY 
--	select T1."CardName", T2."LicTradNum",
--	T0."DocDate",T1."NumAtCard", T0."LineNum",
--	T0."Dscription", T0."LineTotal",
--	T0."VatSum", T0."TaxCode" , T2."Country", T1."DocCur", T0."TotalFrgn", T0."VatSumFrgn"
--	from "INV1" T0  
--	join "OINV" T1   on T0."DocEntry"=T1."DocEntry"
--	join "OCRD" T2   on T2."CardCode"=T1."CardCode"
--	Where T1."DocDate" between :FromDate and :ToDate;
	
	
--	select 'S|'+CustomerName+'|'+CustomerBRN+'|'+cast(InvoiceDate as nvarchar(10))+'|'+InvoiceNumber+
--	'|'+cast(LineNumber as nvarchar(100))+'|'+ProductDescription+'|'+cast(SupplyValueMYR as nvarchar(100))+'|'+
--	cast(GSTValueMYR as nvarchar(100))+'|'+TaxCode+'|'+Country+'|'+
--	FCYCode+'|'+cast(SupplyFCY as nvarchar(100))+'|'+cast(GSTFCY as nvarchar(100))+CHAR(13)+CHAR(10)
--	into supplyxml
--	from #SUPPLY;



-----------------------------Ledger---------------------

--	insert into #LEDGER
--	select T0."RefDate",T1."Account",T2."AcctName",
--	T1."LineMemo", T3."CardName",
--	T0."TransId", T0."TransType", T0."BaseRef",
--	T1."Debit",T1."Credit", 
--	case when "DebCred" ='D' then T1."BalDueDeb" else T1."BalDueCred" end Balance
--	from "OJDT" T0  
--	join "JDT1" T1   on T0."TransId"=T1."TransId"
--	join "OACT" T2   on T2."AcctCode"=T1."Account"
--	left join "OCRD" T3   on T3."CardCode"=T1."ShortName"
--	where T0."RefDate" between :FromDate and :ToDate;
	
	

--	select 'L|'+cast(TransactionDate as nvarchar(10))+'|'+AccountID+'|'+AccountName+'|'+TransactionDescription+'|'+
--	Name+'|'+cast(TransactionID as nvarchar(100))+'|'+SourceDocumentID+'|'+SourceType+'|'+
--	cast(Debit as nvarchar(100))+'|'+cast(Credit as nvarchar(100))+'|'+cast(Balance as nvarchar(100))+CHAR(13)+CHAR(10)
--	into ledgerxml
--	from #LEDGER;
	
	
	
---------------------------footer---------------------
		
--	insert into #FOOTER
--	select 
--		ifnull((select COUNT(*) from #PURCHASE),0) 				as "TotalPurchaseCount",
--		ifnull((select sum(PurchaseValueMYR) from #PURCHASE),0) as "TotalPurchaseAmount",
--		ifnull((select sum(GSTValueMYR) from #PURCHASE),0) 		as "TotalPurchaseAmountGST",
--		ifnull((select COUNT(*) from #SUPPLY),0)  				as "TotalSupplyCount",
--		ifnull((select sum(SupplyValueMYR) from #SUPPLY),0) 	as "TotalSupplyAmount",
--		ifnull((select sum(GSTValueMYR) from #SUPPLY),0)  		as "TotalSupplyAmountGST",
--		ifnull((select COUNT(*) from #LEDGER),0)	 			as "TotalLedgerCount",
--		ifnull((select SUM(Debit) from #LEDGER),0)  			as "TotalLedgerDebit",
--		ifnull((select SUM(credit) from #LEDGER),0)  			as "TotalLedgerCredit",
--		ifnull((select SUM(Balance) from #LEDGER),0)  			as "TotalLedgerBalance" from DUMMY;	
	
	
--	--set @footerxml=''--isnull((select * from @footer FOR XML RAW ('Footer'), ELEMENTS),'')
--	select 'F|'
--	+cast(TotalPurchaseCount as nvarchar(100))+'|'
--	+cast(TotalPurchaseAmount as nvarchar(100))+'|'
--	+cast(TotalPurchaseAmountGST as nvarchar(100))+'|'
--	+cast(TotalSupplyCount as nvarchar(100))+'|'
--	+cast(TotalSupplyAmount as nvarchar(100))+'|'
--	+cast(TotalSupplyAmountGST as nvarchar(100))+'|'
--	+cast(TotalLedgerCount as nvarchar(100))+'|'
--	+cast(TotalLedgerDebit as nvarchar(100))+'|'
--	+cast(TotalLedgerCredit as nvarchar(100))+'|'
--	+cast(TotalLedgerBalance as nvarchar(100))+CHAR(13)+CHAR(10)
--	into footerxml
--	from #FOOTER;
	
	
	select ifnull(companyxml,'') || ifnull(purchasexml,'') || ifnull(supplyxml,'')
		|| ifnull(ledgerxml,'') || ifnull(footerxml,'')
		into GAFXml from DUMMY;
		
	select :GAFXml as "xmlGAF",LENGTH(:GAFXml) as "xmllen" from DUMMY;
	
	drop table "#COMPANY";
	drop table "#FOOTER";
	drop table "#PURCHASE";
	drop table "#SUPPLY";
	drop table "#LEDGER";
	
ENd;