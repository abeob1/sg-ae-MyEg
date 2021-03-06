create proc [dbo].[sp_SAPB1Addon_GST03]
--**{0}**--
@pointat nvarchar(10),
@fromdate datetime,
@todate datetime,
@duedate datetime
as
---------------------VERSION NOTE---------------------
/*
2015-05-11: fix vat amount with credit note
2015-05-11: fix Date range comparation
2015-05-12: fix divide zero
2015-05-13: fix 5a,5b negative value
2015-05-13: add due date
2015-05-14: update point 6a
2015-06-08: update point 19
2015-07-31: update point
2015-09-09: Update as Chun Kit's excel file
2015-09-14: Update the way getting Tax Information
*/
-------------------------------------------------------
declare @GST03DataTbl table
(
Point1 nvarchar(100),Point2 nvarchar(100),
Point3a datetime,Point3b datetime,
Point4 datetime,
Point5a numeric(19,6),Point5b numeric(19,6),
Point6a numeric(19,6),Point6b numeric(19,6),
Point7 numeric(19,6), Point8 numeric(19,6),
Point9 nvarchar(1),
Point10 numeric(19,6),Point11 numeric(19,6),
Point12 numeric(19,6),Point13 numeric(19,6),
Point14 numeric(19,6),Point15 numeric(19,6),
Point16 numeric(19,6),Point17 numeric(19,6),
Point18 numeric(19,6),Point19 numeric(19,6),
Point19_1Code nvarchar(30),Point19_1Value numeric(19,6), Point19_1Per numeric(19,6),
Point19_2Code nvarchar(30),Point19_2Value numeric(19,6), Point19_2Per numeric(19,6),
Point19_3Code nvarchar(30),Point19_3Value numeric(19,6), Point19_3Per numeric(19,6),
Point19_4Code nvarchar(30),Point19_4Value numeric(19,6), Point19_4Per numeric(19,6),
Point19_5Code nvarchar(30),Point19_5Value numeric(19,6), Point19_5Per numeric(19,6),
Point19_6Code nvarchar(30),Point19_6Value numeric(19,6), Point19_6Per numeric (19,6)
)



declare @tbl table(PointAt nvarchar(10), TransID numeric(18,0), 
VatGroup nvarchar(20), Debit numeric(19,6), Credit numeric(19,6), Balance numeric(19,6),
BaseSum numeric(19,6), VatRat numeric(19,6), Memo nvarchar(100), transtype nvarchar(10))

----------------point 1/2/3a/3b/4--------
insert into @GST03DataTbl(Point1,Point2,Point3a,Point3b, Point4)
select TaxIdNum,CompnyName,@fromdate,@todate,@duedate from OADM

----------------BEGIN POINT 5a-----------------
if @pointat='5a' or @pointat='All' or @pointat='A'
begin
	insert into @tbl 
	select '5a', T0.SrcObjAbs, T0.Code, 0,0,
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	) END VatSum,
	
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	) END BaseSum, T0.vatpercent,T0.JrnlMemo,T0.Srcobjtype
	from dbo.B1_VatView T0 
	Join OVTG T1 on T0.Code=T1.Code 
	where DocDate between 
	Convert(DateTime, Convert(VarChar(30), @fromdate, 102) + ' 12:00:00.000 AM', 121)
	and 
	Convert(DateTime, Convert(VarChar(30), @todate, 102) + ' 12:00:00.000 AM', 121)
	and T1.ReportCode in ('SR','DS')
end
update @GST03DataTbl set Point5a=isnull((select SUM(BaseSum) from @tbl where PointAt='5a'),0)
----------------END POINT 5a-----------------

----------------BEGIN POINT 5b-----------------
if @pointat='5b' or @pointat='All' or @pointat='A'
begin
	insert into @tbl
	select '5b', T0.SrcObjAbs, T0.Code, 0,0,
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	) END VatSum,
	
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	) END BaseSum, T0.vatpercent,T0.JrnlMemo,T0.Srcobjtype
	from dbo.B1_VatView T0 
	Join OVTG T1 on T0.Code=T1.Code 
	where DocDate between 
	Convert(DateTime, Convert(VarChar(30), @fromdate, 102) + ' 12:00:00.000 AM', 121)
	and 
	Convert(DateTime, Convert(VarChar(30), @todate , 102) + ' 12:00:00.000 AM', 121)
	and T1.ReportCode in ('SR','DS','AJS')
end
update @GST03DataTbl set Point5b=isnull((select SUM(Balance) from @tbl where PointAt='5b'),0)
----------------END POINT 5b-----------------

 ----------------BEGIN POINT 6a-----------------
if @pointat='6a' or @pointat='All' or @pointat='A'
begin
	insert into @tbl
	select '6a',T0.SrcObjAbs, T0.Code, 0,0,
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	) END VatSum,
	
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	) END BaseSum, T0.vatpercent,T0.JrnlMemo,T0.Srcobjtype
	from dbo.B1_VatView T0 
	Join OVTG T1 on T0.Code=T1.Code
	where DocDate between 
	Convert(DateTime, Convert(VarChar(30), @fromdate, 102) + ' 12:00:00.000 AM', 121)
	and 
	Convert(DateTime, Convert(VarChar(30), @todate, 102) + ' 12:00:00.000 AM', 121)
	and T1.ReportCode in ('TX', 'IM','TX-E43', 'TX-RE','RC')

end
update @GST03DataTbl set Point6a=isnull((select SUM(BaseSum) from @tbl where PointAt='6a'),0)
----------------END POINT 6a-----------------
 
----------------BEGIN POINT 6b-----------------
if @pointat='6b' or @pointat='All' or @pointat='A'
begin
	insert into @tbl
	select '6b',T0.SrcObjAbs, T0.Code, 0,0,
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	) END VatSum,
	
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	) END BaseSum, T0.vatpercent,T0.JrnlMemo,T0.Srcobjtype
	from dbo.B1_VatView T0 
	Join OVTG T1 on T0.Code=T1.Code 
	Where DocDate between 
	Convert(DateTime, Convert(VarChar(30), @fromdate, 102) + ' 12:00:00.000 AM', 121)
	and 
	Convert(DateTime, Convert(VarChar(30), @todate, 102) + ' 12:00:00.000 AM', 121)
	and T1.ReportCode in ('TX', 'IM','TX-E43', 'TX-RE','AJP','RC') 

end
update @GST03DataTbl set Point6b=isnull((select SUM(Balance) from @tbl where PointAt='6b'),0)
----------------END POINT 6b-----------------

----------------BEGIN POINT 7,8--------------
update @GST03DataTbl set Point7=Point5b-Point6b, Point8=Point6b-Point5b
---------------END POINT 7,8-------------

----------------BEGIN POINT 10-----------------
if @pointat='10' or @pointat='All' or @pointat='A'
begin
	insert into @tbl
	select '10',T0.SrcObjAbs, T0.Code, 0,0,
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	) END VatSum,
	
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	) END BaseSum, T0.vatpercent,T0.JrnlMemo,T0.Srcobjtype
	from dbo.B1_VatView T0 
	Join OVTG T1 on T0.Code=T1.Code 
	where
	DocDate between 
	Convert(DateTime, Convert(VarChar(30), @fromdate, 102) + ' 12:00:00.000 AM', 121)
	and 
	Convert(DateTime, Convert(VarChar(30), @todate, 102) + ' 12:00:00.000 AM', 121)
	and T1.ReportCode in ('ZRL')
end
update @GST03DataTbl set Point10=isnull((select SUM(BaseSum) from @tbl where PointAt='10'),0)
----------------END POINT 10-----------------

----------------BEGIN POINT 11-----------------
if @pointat='11' or @pointat='All' or @pointat='A'
begin
	insert into @tbl
	select '11',T0.SrcObjAbs, T0.Code, 0,0,
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	) END VatSum,
	
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	) END BaseSum, T0.vatpercent,T0.JrnlMemo,T0.Srcobjtype
	from dbo.B1_VatView T0 
	Join OVTG T1 on T0.Code=T1.Code 
	where DocDate between 
	Convert(DateTime, Convert(VarChar(30), @fromdate, 102) + ' 12:00:00.000 AM', 121)
	and 
	Convert(DateTime, Convert(VarChar(30), @todate, 102) + ' 12:00:00.000 AM', 121)
	and T1.ReportCode in ('ZRE')
end
update @GST03DataTbl set Point11=isnull((select SUM(BaseSum) from @tbl where PointAt='11'),0)
----------------END POINT 11-----------------

----------------BEGIN POINT 12-----------------
if @pointat='12' or @pointat='All' or @pointat='A'
begin
	insert into @tbl
	select '12', T0.SrcObjAbs, T0.Code, 0,0,
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	) END VatSum,
	
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	) END BaseSum, T0.vatpercent,T0.JrnlMemo,T0.Srcobjtype
	from dbo.B1_VatView T0 
	Join OVTG T1 on T0.Code=T1.Code where T0.Category='O' 
	and DocDate between 
	Convert(DateTime, Convert(VarChar(30), @fromdate, 102) + ' 12:00:00.000 AM', 121)
	and 
	Convert(DateTime, Convert(VarChar(30), @todate, 102) + ' 12:00:00.000 AM', 121)
	and T1.ReportCode in ('ES43','ES')
end
update @GST03DataTbl set Point12=isnull((select SUM(isnull(BaseSum,0)) from @tbl where PointAt='12'),0)
----------------END POINT 12-----------------

----------------BEGIN POINT 13-----------------
if @pointat='13' or @pointat='All' or @pointat='A'
begin
	insert into @tbl
	select '13',T0.SrcObjAbs, T0.Code, 0,0,
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	) END VatSum,
	
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	) END BaseSum, T0.vatpercent,T0.JrnlMemo,T0.Srcobjtype
	from dbo.B1_VatView T0 
	Join OVTG T1 on T0.Code=T1.Code 
	where DocDate between 
	Convert(DateTime, Convert(VarChar(30), @fromdate, 102) + ' 12:00:00.000 AM', 121)
	and 
	Convert(DateTime, Convert(VarChar(30), @todate, 102) + ' 12:00:00.000 AM', 121)
	and T1.ReportCode in ('RS')
end
update @GST03DataTbl set Point13=isnull((select SUM(BaseSum) from @tbl where PointAt='13'),0)
----------------END POINT 13-----------------

------------------BEGIN POINT 14-----------------
if @pointat='14' or @pointat='All' or @pointat='A'
begin
	insert into @tbl
	select '14',T0.SrcObjAbs, T0.Code, 0,0,
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	) END VatSum,
	
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	) END BaseSum, T0.vatpercent,T0.JrnlMemo,T0.Srcobjtype
	from dbo.B1_VatView T0 
	Join OVTG T1 on T0.Code=T1.Code 
	where DocDate between 
	Convert(DateTime, Convert(VarChar(30), @fromdate, 102) + ' 12:00:00.000 AM', 121)
	and 
	Convert(DateTime, Convert(VarChar(30), @todate, 102) + ' 12:00:00.000 AM', 121)
	and T1.ReportCode in ('IS')
end
update @GST03DataTbl set Point14=isnull((select SUM(Balance) from @tbl where PointAt='14'),0)
------------------END POINT 14-----------------

------------------BEGIN POINT 15-----------------
if @pointat='15' or @pointat='All' or @pointat='A'
begin
	insert into @tbl
	select '15',T0.SrcObjAbs, T0.Code, 0,0,
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	) END VatSum,
	
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	) END BaseSum, T0.vatpercent,T0.JrnlMemo,T0.Srcobjtype
	from dbo.B1_VatView T0 
	Join OVTG T1 on T0.Code=T1.Code 
	where DocDate between 
	Convert(DateTime, Convert(VarChar(30), @fromdate, 102) + ' 12:00:00.000 AM', 121)
	and 
	Convert(DateTime, Convert(VarChar(30), @todate, 102) + ' 12:00:00.000 AM', 121)
	and T1.ReportCode in ('IS')
end
update @GST03DataTbl set Point15=isnull((select SUM(BaseSum) from @tbl where PointAt='15'),0)
------------------END POINT 15-----------------

------------------BEGIN POINT 16-----------------
if @pointat='16' or @pointat='All' or @pointat='A'
begin
	insert into @tbl
	select '16',T0.SrcObjAbs, T0.Code, 0,0,
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	) END VatSum,
	
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	) END BaseSum, T0.vatpercent,T0.JrnlMemo,T0.Srcobjtype
	from dbo.B1_VatView T0 
	Join OVTG T1 on T0.Code=T1.Code 
	where DocDate between  
	Convert(DateTime, Convert(VarChar(30), @fromdate, 102) + ' 12:00:00.000 AM', 121)
	and 
	Convert(DateTime, Convert(VarChar(30), @todate, 102) + ' 12:00:00.000 AM', 121)
	and T1.ReportCode in ('TX-CG')

	---FA ONLY--
end
update @GST03DataTbl set Point16=isnull((select SUM(Balance) from @tbl where PointAt='16'),0)
------------------END POINT 16-----------------

------------------BEGIN POINT 17-----------------
if @pointat='17' or @pointat='All' or @pointat='A'
begin
	insert into @tbl
	select '17',T0.SrcObjAbs, T0.Code, 0,0,
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	) END VatSum,
	
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	) END BaseSum, T0.vatpercent,T0.JrnlMemo,T0.Srcobjtype
	from dbo.B1_VatView T0 
	Join OVTG T1 on T0.Code=T1.Code 
	where DocDate between 
	Convert(DateTime, Convert(VarChar(30), @fromdate, 102) + ' 12:00:00.000 AM', 121)
	and 
	Convert(DateTime, Convert(VarChar(30), @todate, 102) + ' 12:00:00.000 AM', 121)
	and T1.ReportCode in ('AJP')
end
update @GST03DataTbl set Point17=isnull((select SUM(isnull(Balance,0)+isnull(BaseSum,0)) from @tbl where PointAt='17'),0)
------------END POINT 17------------

------------------BEGIN POINT 18-----------------
if @pointat='18' or @pointat='All' or @pointat='A'
begin
	insert into @tbl
	select '18',T0.SrcObjAbs, T0.Code, 0,0,
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.VatSum else T0.VatSum end 
		else
			case when CrditDebit='C' then (-1)*T0.VatSum else T0.VatSum end 
		End	
	) END VatSum,
	
	case when T0.Canceled='C' then
	(-1)* 
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	)
	else
	(
		case when T0.Category='O' then
			case when CrditDebit='D' then (-1)*T0.BaseSum else T0.BaseSum end 
		else
			case when CrditDebit='C' then (-1)*T0.BaseSum else T0.BaseSum end 
		End	
	) END BaseSum, T0.vatpercent,T0.JrnlMemo,T0.Srcobjtype
	from dbo.B1_VatView T0 
	Join OVTG T1 on T0.Code=T1.Code 
	where DocDate between  
	Convert(DateTime, Convert(VarChar(30), @fromdate, 102) + ' 12:00:00.000 AM', 121)
	and 
	Convert(DateTime, Convert(VarChar(30), @todate, 102) + ' 12:00:00.000 AM', 121)
	and T1.ReportCode in ('AJS')
end
update @GST03DataTbl set Point18=isnull((select SUM(isnull(Balance,0)+isnull(BaseSum,0)) from @tbl where PointAt='18'),0)
------------END POINT 18------------

------------------BEGIN POINT 19-----------------
declare @tblMisc table (ID int identity(1,1), Code nvarchar(10), Percentage numeric(18,6))

insert into @tblMisc
select top(6) U_MSICCode,U_PERCENTAGE from [@GST_MSIC]
declare @i int
set @i=0
while @i<=(select max(ID) from @tblMisc)
begin
	if @i=1
	begin
		update @GST03DataTbl set 
			Point19_1Value=point5b *(select Percentage from @tblMisc where id=@i),
			Point19_1Per=(select Percentage from @tblMisc where id=@i),
			Point19_1Code=(select Code from @tblMisc where id=@i)
	end
	if @i=2
	begin
		update @GST03DataTbl set 
			Point19_2Value=point5b *(select Percentage from @tblMisc where id=@i),
			Point19_2Per=(select Percentage from @tblMisc where id=@i),
			Point19_2Code=(select Code from @tblMisc where id=@i)
	end
	if @i=3
	begin
		update @GST03DataTbl set 
			Point19_3Value=point5b *(select Percentage from @tblMisc where id=@i),
			Point19_3Per=(select Percentage from @tblMisc where id=@i),
			Point19_3Code=(select Code from @tblMisc where id=@i)
	end
	if @i=4
	begin
		update @GST03DataTbl set 
			Point19_4Value=point5b *(select Percentage from @tblMisc where id=@i),
			Point19_4Per=(select Percentage from @tblMisc where id=@i),
			Point19_4Code=(select Code from @tblMisc where id=@i)
	end
	if @i=5
	begin
		update @GST03DataTbl set 
			Point19_5Value=point5b *(select Percentage from @tblMisc where id=@i),
			Point19_5Per=(select Percentage from @tblMisc where id=@i),
			Point19_5Code=(select Code from @tblMisc where id=@i)
	end
	if @i=6
	begin
		update @GST03DataTbl set
			Point19_6Value =point5b *(select Percentage from @tblMisc where id=@i), 
			Point19_6Per =(select Percentage from @tblMisc where id=@i),
			Point19_6Code=(select Code from @tblMisc where id=@i)
	end
	set @i=@i+1
end

------------------BEGIN POINT 19-----------------
if isnull(@pointat,'')='All'
	select * from @GST03DataTbl
else
begin
	if isnull(@pointat,'')='A'
		select * from @tbl 
	else
		select * from @tbl where
		Charindex (',' + PointAt + ',' ,',' +@pointat+ ',') > 0 
end