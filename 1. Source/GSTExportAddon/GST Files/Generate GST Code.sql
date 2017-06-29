
CREATE TABLE [dbo].[_GSTCODE](
	[Code] [nvarchar](30) NOT NULL,
	[Percent] [numeric](19, 6) NULL,
	[Description] [ntext] NULL,
 CONSTRAINT [KGST03CODE_PR] PRIMARY KEY CLUSTERED 
(
	[Code] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

 if (select COUNT(*) from [_GSTCode] where Code='TX')=0
	Insert into [_GSTCODE] (Code,[Percent],[Description]) 
	values ('TX',6,'Purchase with GST incurred at 6% and Directly attributable to taxable supplies')
 if (select COUNT(*) from [_GSTCode] where Code='IM')=0
	Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('IM',6,'Import of goods with GST incurred')
 if (select COUNT(*) from [_GSTCode] where Code='IS')=0
	Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('IS',0,'Imports under special scheme with no GST incurred (e.g. Approved Trader Scheme, ATMS Scheme)')
 if (select COUNT(*) from [_GSTCode] where Code='BL')=0
	Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('BL',6,'Purchases with GST incurred but not claimable (Disallowance of Input Tax) (e.g. medical expenses for staff)')
 if (select COUNT(*) from [_GSTCode] where Code='NR')=0
	Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('NR',0,'Purchase from non GST-registered supplier with no GST incurred')
 if (select COUNT(*) from [_GSTCode] where Code='ZP')=0
	Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('ZP',0,'Purchase from GST-registered supplier with no GST incurred. (e.g. supplier provides transportation of goods that qualify as international services)')
 if (select COUNT(*) from [_GSTCode] where Code='EP')=0 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('EP',0,'Purchases exempted from GST. E.g. purchase of residential property or financial services')
 if (select COUNT(*) from [_GSTCode] where Code='OP')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('OP',0,'Purcahse transactions which is out of the scope of GST legislation (e.g. purchase of goods overseas).')
 if (select COUNT(*) from [_GSTCode] where Code='TX-E43')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('TX-E43',6,'Purchase with GST incurred directly attributable to incidental exempt supplies.')
 if (select COUNT(*) from [_GSTCode] where Code='TX-N43')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('TX-N43',6,'Purchase with GST incurred directly attributable to non-incidental exempt supplies.')
 if (select COUNT(*) from [_GSTCode] where Code='TX-RE')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('TX-RE',6,'Purchase with GST incurred that is not directly attributable to taxable or exempt supplies.')
 if (select COUNT(*) from [_GSTCode] where Code='GP')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('GP',0,'Purchase transactions which disregarded under GST legislation (e.g. purchase within GST group registration)')
 if (select COUNT(*) from [_GSTCode] where Code='AJP')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('AJP',6,'Any adjustment made to Input Tax e.g.: Bad Debt Relief, Debit Note & other input tax adjustment.')
 if (select COUNT(*) from [_GSTCode] where Code='SR')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('SR',6,'Standard-rated supplies with GST Charged.')
 if (select COUNT(*) from [_GSTCode] where Code='ZRL')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('ZRL',0,'Local supply of zero rated supplies.')
 if (select COUNT(*) from [_GSTCode] where Code='ZRE')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('ZRE',0,'Exportation of goods or services which are subject to zero rated supplies.')
 if (select COUNT(*) from [_GSTCode] where Code='ES43')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('ES43',0,'Incidental Exempt supplies.')
 if (select COUNT(*) from [_GSTCode] where Code='DS')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('DS',6,'Deemed supplies (e.g. transfer or disposal of business assets without condieration).')
 if (select COUNT(*) from [_GSTCode] where Code='OS')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('OS',0,'Out-of-scope supplies.')
 if (select COUNT(*) from [_GSTCode] where Code='ES')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('ES',0,'Exempt supplies under GST.')
 if (select COUNT(*) from [_GSTCode] where Code='RS')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('RS',0,'Relief supply under GST.')
 if (select COUNT(*) from [_GSTCode] where Code='GS')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('GS',0,'Disregarded supplies.')
 if (select COUNT(*) from [_GSTCode] where Code='AJS')=0	 
	 Insert into [_GSTCODE] (Code,[Percent],[Description])  values ('AJS',6,'Any adjustment made to Output Tax e.g: Longer period adjustment, Bad debt recover, outstanding invoice > 6 months & other output tax adjustments.')
 