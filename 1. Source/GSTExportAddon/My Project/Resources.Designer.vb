﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.36392
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'This class was auto-generated by the StronglyTypedResourceBuilder
    'class via a tool like ResGen or Visual Studio.
    'To add or remove a member, edit your .ResX file then rerun ResGen
    'with the /str option, or rebuild your VS project.
    '''<summary>
    '''  A strongly-typed resource class, for looking up localized strings, etc.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Friend Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  Returns the cached ResourceManager instance used by this class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("GSTAddon.Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Overrides the current thread's CurrentUICulture property for all
        '''  resource lookups using this strongly typed resource class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create PROCEDURE sp_B1Addon_BadDebtReverse
        '''--**{0}**--
        '''(
        '''	IN InvoiceNum nvarchar(100),
        '''	IN PaidAmount decimal
        ''')
        '''LANGUAGE SQLSCRIPT
        '''AS
        '''BEGIN
        '''	select 
        '''		case when T1.&quot;DocStatus&quot;=&apos;C&apos; then --Last Payment
        '''		&quot;Debit&quot; -
        '''		ifnull((
        '''			select sum(ifnull(&quot;Credit&quot;,0)) from &quot;JDT1&quot; where &quot;U_InvoiceEntry&quot;=:InvoiceNum
        '''			and &quot;VatGroup&quot;=(select &quot;U_Value&quot; from &quot;@GSTSETUP&quot; where &quot;Code&quot;=&apos;AROutTaxCode&apos;)
        '''			and &quot;Credit&quot;&gt;0
        '''		),0)
        '''		else :PaidAmount*6/106  --Partial Payment
        '''		end as &quot;Debit&quot; , 
        '''	(select &quot;U_Value&quot; [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property B1H_sp_B1Addon_BadDebtReverse() As String
            Get
                Return ResourceManager.GetString("B1H_sp_B1Addon_BadDebtReverse", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create PROCEDURE sp_B1Addon_BadDebtReverse_AP
        '''--**{0}**--
        '''(
        '''	IN InvoiceNum nvarchar(100),
        '''	IN PaidAmount decimal
        ''')
        '''LANGUAGE SQLSCRIPT
        '''AS
        '''BEGIN
        '''
        '''	select 
        '''		case when T1.&quot;DocStatus&quot;=&apos;C&apos; then --Last Payment
        '''		&quot;Debit&quot; -
        '''		ifnull((
        '''			select sum(ifnull(&quot;Credit&quot;,0)) from &quot;JDT1&quot; where &quot;U_InvoiceEntry&quot;=:InvoiceNum
        '''			and &quot;VatGroup&quot;=(select &quot;U_Value&quot; from &quot;@GSTSETUP&quot; where &quot;Code&quot;=&apos;APInTaxCode&apos;)
        '''			and &quot;Credit&quot;&gt;0
        '''		),0)
        '''		else :PaidAmount*6/106  --Partial Payment
        '''		end as &quot;Debit&quot; , 
        '''	(select &quot;U_Va [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property B1H_sp_B1Addon_BadDebtReverse_AP() As String
            Get
                Return ResourceManager.GetString("B1H_sp_B1Addon_BadDebtReverse_AP", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create PROCEDURE  sp_B1Addon_GAF
        '''--**{0}**--
        '''(
        '''	IN FromDate date,
        '''	IN ToDate date
        ''')
        '''LANGUAGE SQLSCRIPT
        '''AS
        '''BEGIN
        '''---------------------------COMPANY---------------------
        '''--declare @company table(
        '''--	BusinessName nvarchar(500),
        '''--	BusinessRN nvarchar(500),
        '''--	GSTNumber nvarchar(500),
        '''--	PeriodStart date,
        '''--	PeriodEnd date,
        '''--	GAFCreationDate date,
        '''--	ProductVersion nvarchar(500),
        '''--	GAFVersion nvarchar(500)
        '''--)
        '''select * FROM &quot;OADM&quot;;
        '''--Insert into @company
        '''--SELECT CompnyName BusinessName, [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property B1H_sp_B1Addon_GAF() As String
            Get
                Return ResourceManager.GetString("B1H_sp_B1Addon_GAF", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create PROCEDURE  sp_B1Addon_GAFtxt
        '''--**{0}**--
        '''(
        '''	IN FromDate date,
        '''	IN ToDate date
        ''')
        '''LANGUAGE SQLSCRIPT
        '''AS
        '''BEGIN
        '''---------------------------COMPANY---------------------
        '''	declare companyxml nvarchar(5000);
        '''	declare purchasexml nvarchar(5000):=&apos;&apos;;
        '''	declare ledgerxml nvarchar(5000):=&apos;&apos;;
        '''	declare supplyxml nvarchar(5000):=&apos;&apos;;
        '''	declare GAFXml nvarchar(5000):=&apos;&apos;;
        '''	declare footerxml nvarchar(5000):=&apos;&apos;;
        '''	
        '''	create local temporary table #COMPANY
        '''	( 
        '''		BusinessName nvarchar(500),
        '''		BusinessRN nva [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property B1H_sp_B1Addon_GAFtxt() As String
            Get
                Return ResourceManager.GetString("B1H_sp_B1Addon_GAFtxt", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to drop  PROCEDURE sp_B1Addon_GSTReturn;
        '''create PROCEDURE sp_B1Addon_GSTReturn
        '''--**{0}**--
        '''(
        '''	IN FromDate datetime,
        '''	IN ToDate datetime
        ''')
        '''LANGUAGE SQLSCRIPT
        '''AS
        '''BEGIN
        '''	DECLARE i decimal:=1;
        '''	declare maxrow decimal:=0;
        '''	declare minrow decimal:=0;
        '''	declare pointat nvarchar(100):=&apos;A&apos;;
        '''	
        '''	select min(&quot;$rowid$&quot;) into minrow from &quot;@GST_MSIC&quot;;
        '''	select max(&quot;$rowid$&quot;) into maxrow from &quot;@GST_MSIC&quot;;
        '''	
        '''	create local temporary table #GST03DATATBL
        '''	(
        '''		Point1 nvarchar(100),Point2 nvarchar(100),
        '''		Point3a  [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property B1H_sp_B1Addon_GSTReturn() As String
            Get
                Return ResourceManager.GetString("B1H_sp_B1Addon_GSTReturn", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create PROCEDURE sp_SAPB1Addon_21DO
        '''--**{0}**--
        '''(
        '''	IN FromDate date,
        '''	IN Status nvarchar(1)
        ''')
        '''LANGUAGE SQLSCRIPT
        '''AS
        '''BEGIN
        '''	select &apos;N&apos; as &quot;ck&quot;,T0.&quot;DocEntry&quot;,T0.&quot;DocNum&quot;,T0.&quot;DocDate&quot;,T0.&quot;CardCode&quot;, 
        '''	 T0.&quot;CardName&quot;,T0.&quot;NumAtCard&quot;,T0.&quot;DocTotal&quot;-T0.&quot;VatSum&quot; as &quot;SubTol&quot;,
        '''	 T0.&quot;VatSum&quot;,T0.&quot;DocTotal&quot;, 
        '''	 ifnull((select &quot;U_Value&quot; from &quot;@GSTSETUP&quot; where &quot;Code&quot;=&apos;DOAct&apos;),&apos;&apos;) as &quot;AccuralAct&quot;,
        '''	 case when ifnull(T0.&quot;U_21Day&quot;,&apos;N&apos;)=&apos;Y&apos; then &apos;YES&apos; else &apos;NO&apos; End as &quot;Status&quot;, T0.&quot;U_21DayJE&quot;,
        '''	 Days_Between(T0.&quot; [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property B1H_sp_SAPB1Addon_21DO() As String
            Get
                Return ResourceManager.GetString("B1H_sp_SAPB1Addon_21DO", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to drop PROCEDURE SP_SAPB1ADDON_GST03;
        '''create PROCEDURE SP_SAPB1ADDON_GST03
        '''--**{0}**--
        '''(
        '''	IN pointat nvarchar(10),
        '''	IN fromdate Date,
        '''	IN todate Date,
        '''	IN duedate Date
        ''')
        '''LANGUAGE SQLSCRIPT
        '''AS
        '''BEGIN
        '''	--DECLARE i decimal:=1;
        '''	--declare maxrow decimal:=0;
        '''	--declare minrow decimal:=0;
        '''	
        '''	
        '''	--select min(&quot;$rowid$&quot;) into minrow from &quot;@GST_MSIC&quot;;
        '''	--select max(&quot;$rowid$&quot;) into maxrow from &quot;@GST_MSIC&quot;;
        '''	
        '''	create local temporary table #GST03DATATBL
        '''	(
        '''		Point1 nvarchar(100),Point2 nvarchar(100),
        ''' [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property B1H_sp_SAPB1Addon_GST03() As String
            Get
                Return ResourceManager.GetString("B1H_sp_SAPB1Addon_GST03", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create PROCEDURE sp_SAPB1Addon_GSTBadDebt
        '''--**{0}**--
        '''(
        '''	IN Date Date,
        '''	IN Debitor nvarchar(100),
        '''	IN Month int,
        '''	IN ClaimAmt nvarchar(1),
        '''	IN Status nvarchar(1)
        ''')
        '''LANGUAGE SQLSCRIPT
        '''AS
        '''BEGIN
        '''	create local temporary table #TBL
        '''	( 
        '''		ck nvarchar(1),
        '''		docentry decimal, 
        '''		DocNum nvarchar(100), 
        '''		Docdate datetime, 
        '''		NumAtCard nvarchar(100),
        '''		CardCode nvarchar(100), 
        '''		CardName nvarchar(100), 
        '''		SubTol decimal, 
        '''		VatSum decimal, 
        '''		DocTotal decimal, 
        '''		paidsum decimal, 
        '''		Balance d [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property B1H_sp_SAPB1Addon_GSTBadDebt() As String
            Get
                Return ResourceManager.GetString("B1H_sp_SAPB1Addon_GSTBadDebt", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create PROCEDURE sp_SAPB1Addon_GSTBadDebt_AP
        '''--**{0}**--
        '''(
        '''	IN Date Date,
        '''	IN Debitor nvarchar(100),
        '''	IN Month int,
        '''	IN ClaimAmt nvarchar(1),
        '''	IN Status nvarchar(1)
        ''')
        '''LANGUAGE SQLSCRIPT
        '''AS
        '''BEGIN
        '''
        '''create local temporary table #TBL
        '''	( 
        '''		ck nvarchar(1),
        '''		docentry decimal, 
        '''		DocNum nvarchar(100), 
        '''		Docdate Datetime, 
        '''		NumAtCard nvarchar(100),
        '''		CardCode nvarchar(100), 
        '''		CardName nvarchar(100), 
        '''		SubTol decimal, 
        '''		VatSum decimal, 
        '''		DocTotal decimal, 
        '''		paidsum decimal, 
        '''		Balan [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property B1H_sp_SAPB1Addon_GSTBadDebt_AP() As String
            Get
                Return ResourceManager.GetString("B1H_sp_SAPB1Addon_GSTBadDebt_AP", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to drop PROCEDURE SP_SAPB1ADDON_PAYMENTCONTRA;
        '''create PROCEDURE SP_SAPB1ADDON_PAYMENTCONTRA
        '''--**{0}**--
        '''(
        '''	IN fromdate Date,
        '''	IN todate Date
        ''')
        '''LANGUAGE SQLSCRIPT
        '''AS
        '''BEGIN
        '''	
        '''	create local temporary table #TBLTEMP1(
        '''		PointAt nvarchar(10), 
        '''		TransID decimal, 
        '''		VatGroup nvarchar(20), 
        '''		Debit decimal, 
        '''		Credit decimal, 
        '''		Balance decimal,
        '''		BaseSum decimal, 
        '''		VatRat decimal, 
        '''		Memo nvarchar(100), 
        '''		transtype nvarchar(10)
        '''	);	
        '''	
        '''
        '''----------------BEGIN POINT 5b-----------------
        '''	inse [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property B1H_sp_SAPB1Addon_PaymentContra() As String
            Get
                Return ResourceManager.GetString("B1H_sp_SAPB1Addon_PaymentContra", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create function uf_GetTaxBalance
        '''--**{0}**--
        '''(
        '''DocEntry decimal
        ''')
        '''returns TaxBalance decimal
        '''LANGUAGE SQLSCRIPT READS SQL DATA AS
        '''begin
        '''	--declare TaxBalance decimal:=0;
        '''	
        '''	select --T1.doctotal-T1.paidSum balance, (T1.doctotal-T1.paidSum)/T1.doctotal*T0.Gtotal LineBalance, T0.vatsum, T0.GTotal,
        '''		sum(
        '''		case when T0.&quot;GTotal&quot;=0 then 0 
        '''		else (T1.&quot;DocTotal&quot;-T1.&quot;PaidSum&quot;)/T1.&quot;DocTotal&quot;*T0.&quot;GTotal&quot;/T0.&quot;GTotal&quot;*T0.&quot;VatSum&quot; 
        '''		End) INTO TaxBalance
        '''	from &quot;INV1&quot; T0  
        '''	join &quot;OINV&quot; T1 on T0.&quot;DocEntry [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property B1H_uf_GetTaxBalance() As String
            Get
                Return ResourceManager.GetString("B1H_uf_GetTaxBalance", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create function uf_GetTaxBalance_AP
        '''--**{0}**--
        '''(
        '''DocEntry decimal
        ''')
        '''returns TaxBalance decimal
        '''LANGUAGE SQLSCRIPT READS SQL DATA AS
        '''begin
        '''	--declare TaxBalance decimal:=0;
        '''	
        '''	select --T1.doctotal-T1.paidSum balance, (T1.doctotal-T1.paidSum)/T1.doctotal*T0.Gtotal LineBalance, T0.vatsum, T0.GTotal,
        '''		sum(
        '''		case when T0.&quot;GTotal&quot;=0 then 0 
        '''		else (T1.&quot;DocTotal&quot;-T1.&quot;PaidSum&quot;)/T1.&quot;DocTotal&quot;*T0.&quot;GTotal&quot;/T0.&quot;GTotal&quot;*T0.&quot;VatSum&quot; 
        '''		End) INTO TaxBalance
        '''	from &quot;PCH1&quot; T0  
        '''	join &quot;OPCH&quot; T1 on T0.&quot;DocEn [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property B1H_uf_GetTaxBalance_AP() As String
            Get
                Return ResourceManager.GetString("B1H_uf_GetTaxBalance_AP", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create proc sp_B1Addon_BadDebtReverse 
        '''--**{0}**--
        '''@InvoiceNum as nvarchar(100),
        '''@PaidAmount as numeric(19,6)
        '''as
        '''
        '''select 
        '''case when T1.DocStatus=&apos;C&apos; then --Last Payment
        '''	Debit -
        '''	isnull((
        '''		select sum(isnull(Credit,0)) from jdt1 where U_InvoiceEntry=@InvoiceNum
        '''		and VatGroup=(select U_value from [@GSTSETUP] where code=&apos;AROutTaxCode&apos;)
        '''		and Credit&gt;0
        '''	),0)
        '''else @PaidAmount*6/106  --Partial Payment
        '''end Debit , 
        '''T0.BaseSum,
        '''(select U_value from [@GSTSETUP] where code=&apos;ACT3&apos;) DrAct, 
        '''(select U [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property sp_B1Addon_BadDebtReverse() As String
            Get
                Return ResourceManager.GetString("sp_B1Addon_BadDebtReverse", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create proc sp_B1Addon_BadDebtReverse_AP 
        '''--**{0}**--
        '''@InvoiceNum as nvarchar(100),
        '''@PaidAmount as numeric(19,6)
        '''as
        '''
        '''select 
        '''	case when T1.DocStatus=&apos;C&apos; then --Last Payment
        '''	Debit -
        '''	isnull((
        '''		select sum(isnull(Credit,0)) from jdt1 where U_InvoiceEntry=@InvoiceNum
        '''		and VatGroup=(select U_value from [@GSTSETUP] where code=&apos;APInTaxCode&apos;)
        '''		and Credit&gt;0
        '''	),0)
        '''	else @PaidAmount*6/106  --Partial Payment
        '''	end Debit , 
        '''(select U_value from [@GSTSETUP] where code=&apos;ACT3&apos;) DrAct, 
        '''(select U_value f [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property sp_B1Addon_BadDebtReverse_AP() As String
            Get
                Return ResourceManager.GetString("sp_B1Addon_BadDebtReverse_AP", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create proc sp_B1Addon_GAF
        '''--**{0}**--
        '''@FromDate datetime,
        '''@ToDate datetime
        '''as
        '''--set @FromDate=&apos;2012-08-01&apos;
        '''--set @ToDate=&apos;2012-08-30&apos;
        '''---------------------------COMPANY---------------------
        '''declare @company table(
        '''	BusinessName nvarchar(500),
        '''	BusinessRN nvarchar(500),
        '''	GSTNumber nvarchar(500),
        '''	PeriodStart date,
        '''	PeriodEnd date,
        '''	GAFCreationDate date,
        '''	ProductVersion nvarchar(500),
        '''	GAFVersion nvarchar(500)
        ''')
        '''
        '''Insert into @company
        '''SELECT CompnyName BusinessName,T0.FreeZoneNo BusinessRN [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property sp_B1Addon_GAF() As String
            Get
                Return ResourceManager.GetString("sp_B1Addon_GAF", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create proc sp_B1Addon_GAFtxt
        '''--**{0}**--
        '''@FromDate datetime,
        '''@ToDate datetime
        '''as
        '''--set @FromDate=&apos;2012-08-01&apos;
        '''--set @ToDate=&apos;2012-08-30&apos;
        '''--sp_B1Addon_GAFtxt &apos;2015-08-01&apos;,&apos;2015-09-01&apos;
        '''---------------------------COMPANY---------------------
        '''declare @company table(
        '''	BusinessName nvarchar(500),
        '''	BusinessRN nvarchar(500),
        '''	GSTNumber nvarchar(500),
        '''	PeriodStart date,
        '''	PeriodEnd date,
        '''	GAFCreationDate date,
        '''	ProductVersion nvarchar(500),
        '''	GAFVersion nvarchar(500)
        ''')
        '''
        '''Insert into @company
        '''SELEC [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property sp_B1Addon_GAFtxt() As String
            Get
                Return ResourceManager.GetString("sp_B1Addon_GAFtxt", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create proc sp_B1Addon_GSTReturn
        '''--**{0}**--
        '''@FromDate datetime,
        '''@ToDate datetime
        '''as
        '''declare @GST03DataTbl table
        '''(
        '''Point1 nvarchar(100),Point2 nvarchar(100),
        '''Point3a datetime,Point3b datetime,
        '''Point4 datetime,
        '''Point5a numeric(19,6),Point5b numeric(19,6),
        '''Point6a numeric(19,6),Point6b numeric(19,6),
        '''Point7 numeric(19,6), Point8 numeric(19,6),
        '''Point9 nvarchar(1),
        '''Point10 numeric(19,6),Point11 numeric(19,6),
        '''Point12 numeric(19,6),Point13 numeric(19,6),
        '''Point14 numeric(19,6),Point15 numeric(19,6 [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property sp_B1Addon_GSTReturn() As String
            Get
                Return ResourceManager.GetString("sp_B1Addon_GSTReturn", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create  proc sp_SAPB1Addon_21DO
        '''--**{0}**--
        '''@FromDate datetime,
        '''@Status nvarchar(1)
        '''as
        '''
        '''select &apos;N&apos; ck,T0.DocEntry,T0.DocNum,T0.DocDate,T0.CardCode, T0.CardName,T0.NumAtCard,T0.DocTotal-T0.vatsum SubTol,
        ''' T0.VatSum,T0.DocTotal, 
        ''' isnull((select U_Value from [@GSTSETUP] where code=&apos;DOAct&apos;),&apos;&apos;) AccuralAct,
        ''' case when isnull(T0.U_21Day,&apos;N&apos;)=&apos;Y&apos; then &apos;YES&apos; else &apos;NO&apos; End Status, T0.U_21DayJE,
        ''' datediff(dd,T0.DocDate,@FromDate) Days,T2.Account TaxAct,
        ''' sum(T1.VatSum*T1.OpenQty/T1.Quantity)  OutstandVatA [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property sp_SAPB1Addon_21DO() As String
            Get
                Return ResourceManager.GetString("sp_SAPB1Addon_21DO", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create proc [dbo].[sp_SAPB1Addon_GST03]
        '''--**{0}**--
        '''@pointat nvarchar(10),
        '''@fromdate datetime,
        '''@todate datetime,
        '''@duedate datetime
        '''as
        '''---------------------VERSION NOTE---------------------
        '''/*
        '''2015-05-11: fix vat amount with credit note
        '''2015-05-11: fix Date range comparation
        '''2015-05-12: fix divide zero
        '''2015-05-13: fix 5a,5b negative value
        '''2015-05-13: add due date
        '''2015-05-14: update point 6a
        '''2015-06-08: update point 19
        '''2015-07-31: update point
        '''2015-09-09: Update as Chun Kit&apos;s excel file
        '''2015 [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property sp_SAPB1Addon_GST03() As String
            Get
                Return ResourceManager.GetString("sp_SAPB1Addon_GST03", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create proc sp_SAPB1Addon_GSTBadDebt
        '''--**{0}**--
        '''@Date datetime,
        '''@Debitor nvarchar(100),
        '''@Month int,
        '''@ClaimAmt nvarchar(1),
        '''@Status nvarchar(1)
        '''as
        '''declare @tbl table (ck nvarchar(1),docentry numeric(19,0), DocNum nvarchar(100), Docdate datetime, NumAtCard nvarchar(100),
        '''CardCode nvarchar(100), CardName nvarchar(100), 
        '''SubTol numeric(19,6), 
        '''VatSum numeric(19,6), 
        '''DocTotal numeric(19,6), 
        '''paidsum numeric(19,6), 
        '''Balance numeric(19,6), 
        '''Status nvarchar(100),
        '''U_BadDebtJE nvarchar(100),
        '''CrAct n [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property sp_SAPB1Addon_GSTBadDebt() As String
            Get
                Return ResourceManager.GetString("sp_SAPB1Addon_GSTBadDebt", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create proc sp_SAPB1Addon_GSTBadDebt_AP
        '''--**{0}**--
        '''@Date datetime,
        '''@Debitor nvarchar(100),
        '''@Month int,
        '''@ClaimAmt nvarchar(1),
        '''@Status nvarchar(1)
        '''as
        '''declare @tbl table (ck nvarchar(1),docentry numeric(19,0), DocNum nvarchar(100), Docdate datetime, NumAtCard nvarchar(100),
        '''CardCode nvarchar(100), CardName nvarchar(100), 
        '''SubTol numeric(19,6), 
        '''VatSum numeric(19,6), 
        '''DocTotal numeric(19,6), 
        '''paidsum numeric(19,6), 
        '''Balance numeric(19,6), 
        '''Status nvarchar(100),
        '''U_BadDebtJE nvarchar(100),
        '''CrAc [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property sp_SAPB1Addon_GSTBadDebt_AP() As String
            Get
                Return ResourceManager.GetString("sp_SAPB1Addon_GSTBadDebt_AP", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to create proc [dbo].[sp_SAPB1Addon_PaymentContra]
        '''--**{0}**--
        '''@fromdate datetime,
        '''@todate datetime
        '''as
        '''--[sp_SAPB1Addon_PaymentContra] &apos;20010101&apos;,&apos;20200101&apos;
        '''
        '''declare @tbl table(PointAt nvarchar(10), TransID numeric(18,0), 
        '''VatGroup nvarchar(20), Debit numeric(19,6), Credit numeric(19,6), Balance numeric(19,6),
        '''BaseSum numeric(19,6), VatRat numeric(19,6), Memo nvarchar(100), transtype nvarchar(10))
        '''
        '''insert into @tbl
        '''exec [sp_SAPB1Addon_GST03] &apos;A&apos;,@fromdate,@Todate,@todate
        '''
        '''
        '''select case when Point [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property sp_SAPB1Addon_PaymentContra() As String
            Get
                Return ResourceManager.GetString("sp_SAPB1Addon_PaymentContra", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to --**{0}**--
        '''{1} function dbo.uf_GetTaxBalance(@DocEntry numeric(18,0))
        '''returns numeric(18,6)
        '''begin
        '''	declare @TaxBalance1 numeric(18,6)
        '''	declare @TaxBalance2 numeric(18,6)
        '''	select --T1.doctotal-T1.paidSum balance, (T1.doctotal-T1.paidSum)/T1.doctotal*T0.Gtotal LineBalance, T0.vatsum, T0.GTotal,
        '''	@TaxBalance1=sum(
        '''		case when T0.GTotal=0 then 0 
        '''		else (T1.doctotal-T1.paidSum)/T1.doctotal*T0.Gtotal/T0.GTotal*T0.VatSum 
        '''		End)
        '''	from INV1 T0 with(nolock) 
        '''	join OINV T1 with(nolock) on T0.DocEntry=T1 [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property uf_GetTaxBalance() As String
            Get
                Return ResourceManager.GetString("uf_GetTaxBalance", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to --**{0}**--
        '''{1} function dbo.uf_GetTaxBalance_AP(@DocEntry numeric(18,0))
        '''returns numeric(18,6)
        '''begin
        '''	declare @TaxBalance1 numeric(18,6)
        '''	declare @TaxBalance2 numeric(18,6)
        '''	select --T1.doctotal-T1.paidSum balance, (T1.doctotal-T1.paidSum)/T1.doctotal*T0.Gtotal LineBalance, T0.vatsum, T0.GTotal,
        '''	@TaxBalance1=sum(
        '''		case when T0.GTotal=0 then 0 
        '''		else (T1.doctotal-T1.paidSum)/T1.doctotal*T0.Gtotal/T0.GTotal*T0.VatSum 
        '''		End)
        '''	from PCH1 T0 with(nolock) 
        '''	join OPCH T1 with(nolock) on T0.DocEntry [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property uf_GetTaxBalance_AP() As String
            Get
                Return ResourceManager.GetString("uf_GetTaxBalance_AP", resourceCulture)
            End Get
        End Property
    End Module
End Namespace