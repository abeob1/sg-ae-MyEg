create proc sp_B1Addon_GSTReturn
--**{0}**--
@FromDate datetime,
@ToDate datetime
as
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
insert into @GST03DataTbl
exec sp_SAPB1Addon_GST03 'All',@FromDate,@ToDate,@ToDate

select 
isnull(CONVERT(nvarchar(100),point5a),'0.00') +'|'+
isnull(CONVERT(nvarchar(100),point5b),'0.00') +'|'+
isnull(CONVERT(nvarchar(100),point6a),'0.00') +'|'+
isnull(CONVERT(nvarchar(100),point6b),'0.00') +'|'+
'1' +'|'+

isnull(CONVERT(nvarchar(100),point10),'0.00') +'|'+
isnull(CONVERT(nvarchar(100),point11),'0.00') +'|'+
isnull(CONVERT(nvarchar(100),point12),'0.00') +'|'+
isnull(CONVERT(nvarchar(100),point13),'0.00') +'|'+
isnull(CONVERT(nvarchar(100),point14),'0.00') +'|'+
isnull(CONVERT(nvarchar(100),point15),'0.00') +'|'+
isnull(CONVERT(nvarchar(100),point16),'0.00') +'|'+
isnull(CONVERT(nvarchar(100),point17),'0.00') +'|'+
--isnull(CONVERT(nvarchar(100),point18),'0.00') +'|'+

isnull(CONVERT(nvarchar(100),Point19_1Code),'')+'|'+
isnull(CONVERT(nvarchar(100),Point19_1Value/100),'0.00')+'|'+
isnull(CONVERT(nvarchar(100),Point19_2Code),'')+'|'+
isnull(CONVERT(nvarchar(100),Point19_2Value/100),'0.00')+'|'+
isnull(CONVERT(nvarchar(100),Point19_3Code),'')+'|'+
isnull(CONVERT(nvarchar(100),Point19_3Value/100),'0.00')+'|'+
isnull(CONVERT(nvarchar(100),Point19_4Code),'')+'|'+
isnull(CONVERT(nvarchar(100),Point19_4Value/100),'0.00')+'|'+
isnull(CONVERT(nvarchar(100),Point19_5Code),'')+'|'+
isnull(CONVERT(nvarchar(100),Point19_5Value/100),'0.00')+'|0'

txtFilestr
 from @GST03DataTbl
