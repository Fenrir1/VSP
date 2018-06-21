<%@ language="VBScript"%><!-- #include file="const.asp" --><!-- #include file="common.asp" -->
<%
' основной экран мониторинга VIP банкоматов БПТ

' подключение к БД
set Conn=Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeout=180
Conn.CommandTimeout=10
Conn.Open(ConnectionString)

Auth_Name=Request.ServerVariables("AUTH_USER")
set Cmd=Server.CreateObject("ADODB.Command")
Cmd.ActiveConnection=Conn
Cmd.CommandType=adCmdStoredProc
'Фиксирование входа
Cmd.CommandText="sp_LogEnters"
Cmd.Parameters.Refresh
Cmd.Parameters("@Login") = Auth_Name
Cmd.Parameters("@UserName") = ""
Cmd.Execute
FUserName=Cmd.Parameters("@UserName")

if Len(FUserName)=0 then ' если юзер не зареген
  Conn.Close
  set Conn = Nothing
  set Cmd = Nothing
  Response.Write("<html><body><div style='text-align: center;'><span style='font-size: 14pt; font-weight: 600; color: #800000}'>Для пользовате main ля "&Auth_Name&" доступ не определен.</span></div></body></html>")
else ' юзер зареген, продолжаем:

' далее (до начала html кода) в рабочие переменные считываем данные из БД для отображения элементов на странице
set Rs=Server.CreateObject("ADODB.Recordset")

function parseInterval(minInterval)
	res = ""
	if (minInterval>=24*60) then
		res = res&cStr(minInterval\(24*60))&" д. "
		minInterval = minInterval mod (24*60)
	end if
	if (minInterval>=60) then
		res = res&cStr(minInterval\60)&" ч. "
		minInterval = minInterval mod 60
	end if
	if (minInterval>=0) then
		res = res&cStr(minInterval)&" м."
	end if
	parseInterval = res
end function

'--------------------------------------------------------------------------------
'------START: Auto update series-------------------------------------------------
'--------------------------------------------------------------------------------
tempTime = Now
sqlstr = "set dateformat dmy; exec sp_Update_Interwal_Series @Date='"&DateTimeFormat(tempTime, "dd.mm.yy hh:nn")&"'"
'Rs.Open sqlstr, Conn

sqlstr = "set dateformat dmy; exec sp_Update_Warning_Series @Date='"&DateTimeFormat(tempTime, "dd.mm.yy hh:nn")&"'"
'Rs.Open sqlstr, Conn
'--------------------------------------------------------------------------------
'------END: Auto update series---------------------------------------------------
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
'------START: get main vars---------------------------------------------
'--------------------------------------------------------------------------------
Color2=0
dim Colors2(4,2)
dim Text2(4,2)
for i=1 to 4
  for j=1 to 2
    Colors2(i, j)=0
	Text2(i,j)=""
  next
next  

dim series(10), ATMID(10,4), TOP5TABLES(5)
for i=0 to 10 
  series(i)=""
  ATMID(i,0)="" ' NumEmv
  ATMID(i,1)=""	' DeviceId
  ATMID(i,2)=""	' DeviceType
  ATMID(i,3)=0	' DaysCount
next

for i=0 to 5
	TOP5TABLES(i)="<table width=""97%"" cellspacing=""0"" ><tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >&nbsp;</td></tr>"&_
				  "<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >&nbsp;</td></tr>"&_
				  "<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >&nbsp;</td></tr>"&_
				  "<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >&nbsp;</td></tr>"&_
				  "<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >&nbsp;</td></tr>"&_
				  "<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >"&_
				  "<div style='width= 100%; height: 62px; overflow: hidden;' >&nbsp;</div></td></tr></table>"
next

    ATMIDList = ""
	sqlstr = "select top 10 vd.DeviceID,vd.NumEmv,vd.DeviceType,ISNULL(vd.DaysCount,0) DaysCount,vd.City,vd.Address,vd.Location,vd.SERIAL_NUMBER,vd.Type "
	sqlstr = sqlstr&" from VIP_Day_Order  vdo join VIP_Device vd on vd.DeviceID=vdo.DeviceID and vd.DeviceType=vdo.DeviceType "
	'sqlstr = sqlstr&" where exists(select * from VIP_Intervals_Series where "
	'sqlstr = sqlstr&" DeviceType=vd.DeviceType and DeviceId=vd.DeviceId and [Time]>=convert(datetime,FLOOR(convert(float,GETDATE())))) "
	sqlstr = sqlstr&" order by OrderNum "

    count = 0
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        do while (not Rs.EOF)and(count<=10)
            found = 0
            for i=0 to 10 
              if ((ATMID(i,0)=Rs.Fields("NumEmv"))and(ATMID(i,1)=Rs.Fields("DeviceType"))) then
                found = 1
                exit for
              end if
            next
            if (found=0) then
                ATMID(count,0)=Rs.Fields("NumEmv")
				ATMID(count,1)=Rs.Fields("DeviceType")
				ATMID(count,2)=Rs.Fields("DeviceID")
				ATMID(count,3)=Rs.Fields("DaysCount")
                if (ATMIDList<>"") then ATMIDList = ATMIDList&"," end if
                ATMIDList = ATMIDList&Rs.Fields("DeviceID")
                count = count+1
            end if
            Rs.MoveNext
        loop
    end if
    Rs.Close

	sqlstr = "select NumEmv, DeviceType, DeviceID, ISNULL(DaysCount,0) DaysCount from VIP_Device vd where exists(select * from VIP_Intervals_Series where "
	sqlstr = sqlstr&" DeviceType=vd.DeviceType and DeviceId=vd.DeviceId and [Time]>=convert(datetime,FLOOR(convert(float,GETDATE())))) "
	sqlstr = sqlstr&" order by ISNULL(DaysCount,0), DeviceType, DeviceID "
	

    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        do while ((not Rs.EOF)and(count<=10))
            found = 0
            for i=0 to 10 
              if ((ATMID(i,0)=Rs.Fields("NumEmv"))and(ATMID(i,1)=Rs.Fields("DeviceType"))) then
                found = 1
                exit for
              end if
            next
            if (found=0) then
                ATMID(count,0)=Rs.Fields("NumEmv")
				ATMID(count,1)=Rs.Fields("DeviceType")
				ATMID(count,2)=Rs.Fields("DeviceID")
				ATMID(count,3)=Rs.Fields("DaysCount")
                if (ATMIDList<>"") then ATMIDList = ATMIDList&"," end if
                ATMIDList = ATMIDList&Rs.Fields("DeviceID")
                count = count+1
            end if
            Rs.MoveNext
        loop
    end if
    Rs.Close	
	
	'------------------------------------
	'--START: Series for Chart1----------
	'------------------------------------
	for i=0 to 10 
	  if (ATMID(i,0)="") then
		exit for
	  end if
	  
	  serieID = i
	  v=9.5-serieID
	  v=replace(v, ",", ".")
	  
	  sqlstr = "select DeviceID,DeviceType,dateadd(month,-1,[TIME]) [TIME] from VIP_Intervals_Series where DeviceID="&ATMID(i,2)&" and DeviceType='"&ATMID(i,1)&"' "
	  sqlstr = sqlstr&" and [TIME]>=convert(datetime,FLOOR(convert(float,Getdate()))) order by [TIME]"
	  Rs.Open sqlstr, Conn
	  If not Rs.EOF then
		do while (not Rs.EOF)
			'series(serieID) = series(serieID)&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("Start_time"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"},"
			series(serieID)=series(serieID)&vbCrLf&"{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"},"
			Rs.MoveNext
        loop
	  else

		series(serieID)=series(serieID)&vbCrLf&"{color: null, marker: {fillColor: '#363636', lineColor: '#363636', radius: 5}, x: Date.UTC("&DateTimeFormat(dateadd("m",-1,Now), "yyyy, mm, dd, 10, 15")&"), y: "&v&"},"

	  end if
	  Rs.Close
  
	next
	'------------------------------------
	'--END: Series for Chart1------------
	'------------------------------------

AllSeries=""
for i=0 to 10
  if series(i)<>"" then
    series(i)=left(series(i), len(series(i))-1)
    'scatter
	if (i>0) then
		series(i)=", { name: '"&ATMID(i,0)&"', type: 'scatter', data: ["&series(i)&"]}"  
	else 
		series(i)=" { name: '"&ATMID(i,0)&"', type: 'scatter', data: ["&series(i)&"]}" 
	end if 
	'CID(i)=left(CID(i), 12)
  end if
  AllSeries=AllSeries+series(i)
next

if (AllSeries="") then
	AllSeries = AllSeries & "{ name: '00', type: 'scatter', visible: false, data: [ "
	AllSeries = AllSeries & " {color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 2}, "
 	AllSeries = AllSeries & " x: Date.UTC("&DateTimeFormat(Now, "yyyy, mm, dd, hh, nn")&"), y: 0 } ]} "
end if
'------------------------------------
'--END: Series for Chart1------------
'------------------------------------



'------------------------------------
'--START: Series for Chart2----------
'------------------------------------
AllSeries2=""


if (DatePart("h",Now)>8) then
	'tempTime =  cDate(DateTimeFormat(DateAdd("h", -8, Now), "dd.mm.yy hh:nn"))
	tempTime =  DateAdd("h", -8, Now)
	if (DatePart("n",tempTime)>32) then
		tempTime =  cDate(DateTimeFormat(DateAdd("h", -7, Now), "dd.mm.yy hh:00"))
	elseif (DatePart("n",tempTime)>5) then
		tempTime =  cDate(DateTimeFormat(DateAdd("h", -8, Now), "dd.mm.yy hh:30"))
	end if
else
	tempTime =  cDate(cStr(Date)&" 00:30:00")
end if

'tempTime = cDate(cStr(Date)&" 00:30:00")
tempStopTime = Now

do while (datediff("n",tempTime,tempStopTime)>=1)
	FailDeviceCount = 0
	sqlstr = "set dateformat dmy; select count(DeviceId) DeviceCount from VIP_Intervals_Series where [TIME]='"&DateTimeFormat(tempTime, "dd.mm.yy hh:nn")&"' "
	Rs.Open sqlstr, Conn
	If not Rs.EOF then
		FailDeviceCount = Rs.Fields("DeviceCount")
	end if
	Rs.Close
	
	if (AllSeries2="") then
		AllSeries2=AllSeries2&vbCrLf&"{color: '#99ccff', x: Date.UTC("&DateTimeFormat(DateAdd("m", -1, tempTime), "yyyy, mm, dd, hh, nn")&"), y: "&FailDeviceCount&"}"
	else
		AllSeries2=AllSeries2&vbCrLf&",{color: '#99ccff', x: Date.UTC("&DateTimeFormat(DateAdd("m", -1, tempTime), "yyyy, mm, dd, hh, nn")&"), y: "&FailDeviceCount&"}"
	end if 
			
	tempTime = DateAdd("n", 30, tempTime)
loop

AllSeries2=" { name: 'CountFailDevices', type: 'column', data: ["&AllSeries2&"]}"

'------------------------------------
'--END: Series for Chart2------------
'------------------------------------

'------------------------------------
'--START: Data CircleIndicator1------
'------------------------------------
sqlstr = " select max(LastUpdate) lt  from VIP_files"
Rs.Open sqlstr, Conn
If not Rs.EOF then
	LastFileTimeFull = DateTimeFormat(Rs.Fields("lt"),"dd.mm.yyyy hh:nn")
	LastFileTime = DateTimeFormat(Rs.Fields("lt"),"hh:nn")
end if
Rs.Close


circleIndicatorColor = clNormal

sqlstr = " if exists(SELECT * from VIP_Intervals_Union where Closed<>1 and isNull(onControl,0)=0)  select '#FF0000' cl "
sqlstr = sqlstr&" else if exists(SELECT * from VIP_Intervals_Union where Closed<>1 and isNull(onControl,0)=1)  select '#FFFF00' cl "
sqlstr = sqlstr&" else select '#00FF00' cl "
Rs.Open sqlstr, Conn
If not Rs.EOF then
	circleIndicatorColor = Rs.Fields("cl")
end if
Rs.Close

TimeEV = 0
TimeVA = 0
TimeVC = 0
TimeFV = 0
sqlstr = "select FileType,LastUpdate, DATEDIFF(MINUTE,LastUpdate,GETDATE()) IntervalLength from  VIP_files"
Rs.Open sqlstr, Conn
If not Rs.EOF then
do while (not Rs.EOF)
	if (Rs.Fields("FileType")="EV") then
		TimeEV = Rs.Fields("IntervalLength")
	elseif (Rs.Fields("FileType")="VA") then
		TimeVA = Rs.Fields("IntervalLength")
	elseif (Rs.Fields("FileType")="VC") then
		TimeVC = Rs.Fields("IntervalLength")
	elseif (Rs.Fields("FileType")="FV") then
		TimeFV = Rs.Fields("IntervalLength")
	end if
	Rs.MoveNext
loop
end if
Rs.Close



	PeriodVC = 30
	PeriodVA = 30
	PeriodFV = 10
	PeriodEV = 10
sqlstr = "select PeriodVC,PeriodVA,PeriodFV,PeriodEV from VIP_Config"
Rs.Open sqlstr, Conn
If not Rs.EOF then
	PeriodVC = cInt(Rs.Fields("PeriodVC"))
	PeriodVA = cInt(Rs.Fields("PeriodVA"))
	PeriodFV = cInt(Rs.Fields("PeriodFV"))
	PeriodEV = cInt(Rs.Fields("PeriodEV"))

end if
Rs.Close


circleIndicatorMarker = ""
if (((TimeEV>PeriodEV) and (PeriodEV>0))or((TimeFV>PeriodFV)and (PeriodFV>0))or((TimeVA>PeriodVA)and (PeriodVA>0))or((TimeVC>PeriodVC)and (PeriodEV>0))) then
	circleIndicatorMarker = "circleIndicator.renderer.image('q.gif', 75, 75, 150, 150).add();"
end if 
'------------------------------------
'--END: Data CircleIndicator1--------
'------------------------------------

'------------------------------------
'--START: Data Table1----------------
'------------------------------------
ValueAll = 0
ValueATM = 0
ValueBPT = 0
ValueAll_discr = "0 ATM; 0 BPT"
ValueFail = 0
ValueFailATM = 0
ValueFailBPT = 0
ValueFail_discr = "0 ATM; 0 BPT"
ValueControl = 0
ValueControlATM = 0
ValueControlBPT = 0
ValueControl_discr = "0 ATM; 0 BPT"

ValueFail_noUnion = 0
ValueFail_link = 0
ValueFail_money = 0
ValueFail_hardware = 0
ValueFail_service = 0


sqlstr = "select DeviceType,count(distinct DeviceId) DeviceCount from VIP_Device group by DeviceType"
Rs.Open sqlstr, Conn
If not Rs.EOF then
do while (not Rs.EOF)
	if (Rs.Fields("DeviceType")="ATM") then
		ValueATM = Rs.Fields("DeviceCount")
	end if
	
	if (Rs.Fields("DeviceType")="BPT") then
		ValueBPT = Rs.Fields("DeviceCount")
	end if
	Rs.MoveNext
loop
end if
Rs.Close

ValueAll = ValueATM+ValueBPT
ValueAll_discr = ValueATM&" ATM; "&ValueBPT&" BPT"

sqlstr = "select DeviceType,count(distinct DeviceId) DeviceCount from VIP_Intervals_Union where Closed=0 group by DeviceType"
Rs.Open sqlstr, Conn
If not Rs.EOF then
do while (not Rs.EOF)
	if (Rs.Fields("DeviceType")="ATM") then
		ValueFailATM = Rs.Fields("DeviceCount")
	end if
	
	if (Rs.Fields("DeviceType")="BPT") then
		ValueFailBPT = Rs.Fields("DeviceCount")
	end if
	Rs.MoveNext
loop
end if
Rs.Close

ValueFail = ValueFailATM+ValueFailBPT
ValueFail_discr = ValueFailATM&" ATM; "&ValueFailBPT&" BPT"

sqlstr = "select DeviceType,count(distinct DeviceId) DeviceCount from VIP_Intervals_Union where Closed=0 and ISNULL(onControl,0)=1 group by DeviceType"
Rs.Open sqlstr, Conn
If not Rs.EOF then
do while (not Rs.EOF)
	if (Rs.Fields("DeviceType")="ATM") then
		ValueControlATM = Rs.Fields("DeviceCount")
	end if
	
	if (Rs.Fields("DeviceType")="BPT") then
		ValueControlBPT = Rs.Fields("DeviceCount")
	end if
	Rs.MoveNext
loop
end if
Rs.Close

ValueControl = ValueControlATM+ValueControlBPT
ValueControl_discr = ValueControlATM&" ATM; "&ValueControlBPT&" BPT"

sqlstr = "select [Class],count(DeviceId) DeviceCount from VIP_Intervals where Closed=0 group by [Class]"
Rs.Open sqlstr, Conn
If not Rs.EOF then
do while (not Rs.EOF)
	if (Rs.Fields("Class")="Link") then
		ValueFail_link = Rs.Fields("DeviceCount")
	elseif (Rs.Fields("Class")="Money") then
		ValueFail_money = Rs.Fields("DeviceCount")
	elseif (Rs.Fields("Class")="Hardware") then
		ValueFail_hardware = Rs.Fields("DeviceCount")
	elseif (Rs.Fields("Class")="Service") then
		ValueFail_service = Rs.Fields("DeviceCount")
	end if
	Rs.MoveNext
loop
end if
Rs.Close

ValueFail_noUnion = ValueFail_link + ValueFail_money + ValueFail_hardware + ValueFail_service

'------------------------------------
'--END: Data Table1------------------
'------------------------------------	


'------------------------------------
'--START: Data TOP 5-----------------
'------------------------------------
IntervalsCount = 0
TopCount = Request("topcount")
if ISNULL(TopCount) then 
	TopCount = 0
end if


TimeToReloadTableFull = 5*60
sqlstr = "select IsNUll(Top5Interval,300) Top5Interval from VIP_Config"
Rs.Open sqlstr, Conn
If not Rs.EOF then
	TimeToReloadTableFull = Rs.Fields("Top5Interval")
end if
Rs.Close

TimeToReloadTable = Request("timetoreload")
if ISNULL(TimeToReloadTable)or(TimeToReloadTable=0) then 
	TimeToReloadTable = TimeToReloadTableFull
end if

TimeToReloadPage = 5*60
if (cInt(TimeToReloadTable)<=cInt(TimeToReloadPage)) then
	TimeToReloadPage = TimeToReloadTable
	TimeToReloadTable = TimeToReloadTableFull
else
	TimeToReloadTable = TimeToReloadTable - TimeToReloadPage
end if


Select Case TopCount
	Case 0 
		'TOP 5 - order by StartTime
		sqlstr = "select  top 5 vd.NumEmv,vd.DeviceType, vd.[City], vd.[Address], datediff(MINUTE,viu.[StartTime],Getdate()) IntervalLength, "
		sqlstr = sqlstr&" ISNULL(viu.[Comment1],'&nbsp;') Comment1, ISNULL(viu.[Comment2],'&nbsp;') Comment2 "
		sqlstr = sqlstr&" from VIP_Intervals_Union viu join VIP_Device vd on vd.DeviceId=viu.DeviceId and vd.DeviceType=viu.DeviceType "
		sqlstr = sqlstr&" where viu.Closed=0 ORDER BY viu.[StartTime]"

	Case 1
		'TOP 5 - order by StartTime desc
		sqlstr = "select  top 5 vd.NumEmv,vd.DeviceType, vd.[City], vd.[Address], datediff(MINUTE,viu.[StartTime],Getdate()) IntervalLength, "
		sqlstr = sqlstr&" ISNULL(viu.[Comment1],'&nbsp;') Comment1, ISNULL(viu.[Comment2],'&nbsp;') Comment2 "
		sqlstr = sqlstr&" from VIP_Intervals_Union viu join VIP_Device vd on vd.DeviceId=viu.DeviceId and vd.DeviceType=viu.DeviceType "
		sqlstr = sqlstr&" where viu.Closed=0 ORDER BY viu.[StartTime] desc"	
		
End Select

TableNum = 0 
Rs.Open sqlstr, Conn
If not Rs.EOF then
do while (not Rs.EOF)
	TOP5TABLES(TableNum)="<table width=""97%"" cellspacing=""0"" >"&_
					"<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >"&Left(cStr(Rs.Fields("DeviceType")&" "&Rs.Fields("NumEmv")),22)&"</td></tr>"&_
					"<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >"&Left(cStr(Rs.Fields("City")),22)&"</td></tr>"&_
					"<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >"
	addr=cStr(Rs.Fields("Address"))
	if (Len(cStr(addr))>25) then
		TOP5TABLES(TableNum)= TOP5TABLES(TableNum)&"<marquee loop='infinite' width='365' >"&addr&"</marquee></td></tr>"
	else
		TOP5TABLES(TableNum)= TOP5TABLES(TableNum)&"<div style='width: 368px; height: 31px;  word-wrap: break-word; overflow: hidden;' >"&addr&"</div></td></tr>"
	end if 
	
	TOP5TABLES(TableNum)= TOP5TABLES(TableNum)&"<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >"&Left(cStr(parseInterval(Rs.Fields("IntervalLength"))),22)&"</td></tr>"&_
					"<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >"&Left(cStr(Rs.Fields("Comment2")),22)&"</td></tr>"&_
					"<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' ><div style='width: 360px; height: 62px;  word-wrap: break-word; overflow: hidden;' >"&Rs.Fields("Comment1")&"</div></td></tr></table>"

	TableNum = TableNum + 1
	Rs.MoveNext
loop
end if
Rs.Close

TopCount = 1 - 1*TopCount



'TableNum = TopCount+1
'sqlstr = "WITH C1 AS (select ROW_NUMBER() OVER (ORDER BY viu.[StartTime]) num, vd.NumEmv,vd.DeviceType, vd.[City], vd.[Address], datediff(MINUTE,viu.[StartTime],Getdate()) IntervalLength, "
'sqlstr = sqlstr&" ISNULL(viu.[Comment1],'&nbsp;') Comment1, ISNULL(viu.[Comment2],'&nbsp;') Comment2 "
'sqlstr = sqlstr&" from VIP_Intervals_Union viu join VIP_Device vd on vd.DeviceId=viu.DeviceId and vd.DeviceType=viu.DeviceType "
'sqlstr = sqlstr&" where viu.Closed=0) SELECT *, (SELECT MAX(num) FROM C1) IntervalsCount FROM C1 order by num "
'Rs.Open sqlstr, Conn
'If not Rs.EOF then
'IntervalsCount = Rs.Fields("IntervalsCount")
'do while ((not Rs.EOF)and(TableNum<=(TopCount+5)))



'	if (CInt(Rs.Fields("num"))>=(TopCount+1)) then
'		TOP5TABLES(TableNum-TopCount-1)="<table width=""97%"" cellspacing=""0"" >"&_
'					  "<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >"&Rs.Fields("DeviceType")&" "&Rs.Fields("NumEmv")&"</td></tr>"&_
'					  "<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >"&Rs.Fields("City")&"</td></tr>"&_
'					  "<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >"
'		if (Len(cStr(Rs.Fields("Address")))>25) then
'			TOP5TABLES(TableNum-TopCount-1)= TOP5TABLES(TableNum-TopCount-1)&"<marquee loop='infinite' width='365' >"&Rs.Fields("Address")&"</marquee></td></tr>"
'		else
'			TOP5TABLES(TableNum-TopCount-1)= TOP5TABLES(TableNum-TopCount-1)&Rs.Fields("Address")&"</td></tr>"
'		end if 
		
'		TOP5TABLES(TableNum-TopCount-1)= TOP5TABLES(TableNum-TopCount-1)&"<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >"&parseInterval(Rs.Fields("IntervalLength"))&"</td></tr>"&_
'					  "<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' >"&Rs.Fields("Comment2")&"</td></tr>"&_
'					  "<tr><td style='border: solid 1px #C0C0C0; font-size: 18pt; font-weight: 400;' ><div style='width= 100%; height: 62px; overflow: hidden;' >"&Rs.Fields("Comment1")&"</div></td></tr></table>"

'		TableNum = TableNum + 1
'	end if
'	Rs.MoveNext
'loop
'end if
'Rs.Close

'if (cInt(IntervalsCount)>(TopCount+5)) then 
'	TopCount = TopCount+5
'else 
'	TopCount = 0
'end if
'------------------------------------
'--END: Data TOP 5-------------------
'------------------------------------	

VIP_Title = ""
sqlstr = "select VIP_Title from VIP_Config"
Rs.Open sqlstr, Conn
If not Rs.EOF then
	VIP_Title = Rs.Fields("VIP_Title")
end if
Rs.Close


'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------


if (DatePart("h",Now)>8) then
	CurrentStratTime = DateTimeFormat(DateAdd("m", -1, DateAdd("h", -8, Now)), "yyyy, mm, dd, hh, nn")
else
	CurrentStratTime = DateTimeFormat(DateAdd("m", -1, Now), "yyyy, mm, dd, 0, 0")
end if

if (DatePart("h",Now)>16) then
	CurrentEndTime = DateTimeFormat(DateAdd("m", -1, Now), "yyyy, mm, dd, hh, nn")
else
	CurrentEndTime = DateTimeFormat(DateAdd("m", -1, Now), "yyyy, mm, dd, hh, nn")
end if

'CurrentStratTime = DateTimeFormat(DateAdd("m", -1, Now), "yyyy, mm, dd, 0, 0")
CurrentTime = DateTimeFormat(DateAdd("m", -1, Now), "yyyy, mm, dd, hh, nn")
'CurrentTimeLabel = DateTimeFormat(DateAdd("m", -1, Now), "hh:nn")
CurrentTimeLabel = DateTimeFormat(DateAdd("m", -1, cDate(LastFileTimeFull)), "hh:nn")
'CurrentEndTime = DateTimeFormat(DateAdd("m", -1, Now), "yyyy, mm, dd, 23, 59")
%>
<!DOCTYPE HTML>
<html>
<head>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">

		<meta http-equiv='refresh' content='<%=TimeToReloadPage %>; url=main3.asp?topcount=<%=TopCount %>&timetoreload=<%=TimeToReloadTable %>'>

		<!-- 1. Add these JavaScript inclusions in the head of your page -->
		<script type="text/javascript" src="js/jquery.min.js"></script>
		<script type="text/javascript" src="js/highcharts.js"></script>
		<script type="text/javascript" src="js/themes/gray.js"></script>
		<!-- 2. Add the JavaScript to initialize the chart on document ready -->
		<script type="text/javascript">
		
			var chart1;
			var chart2;
			var circleIndicator;

			var FlagOut=1;
			
			var now = new Date(); 
			var now_utc =  Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(),  now.getHours(), now.getMinutes(), now.getSeconds());			
			

			$(document).ready(function() {

			    // Третий график
			    circleIndicator = new Highcharts.Chart({
			        chart:   {renderTo: 'circleIndicator', type: 'line', margin: [0, 0, 0, 0] },
			        credits: {enabled: false},	legend:  {enabled: false},	tooltip: {enabled: false},	title:   {text: ''}
			    });
			    circleIndicator.renderer.circle(150, 150, 90).attr({
			        fill: '<% =circleIndicatorColor %>',
			        stroke: '<% =circleIndicatorColor %>'
			    }).add();
				
				<%=circleIndicatorMarker %>

				
				// Первый график
			    chart1 = new Highcharts.Chart({
			        chart: {
			            renderTo: 'container1',
						ignoreHiddenSeries : false,
						marginLeft: 178,
			            marginRight: 20
			        },
			        credits: {enabled: false},
			        legend:  {enabled: false},
			        tooltip: {enabled: false},
			        title:   {align: 'right', text: 'Интервалы неработоспособности устройств'},
			        subtitle: {align: 'left', text: 'ID'},
			        xAxis: {
			            min: Date.UTC(<%=CurrentStratTime %>),
			            max: Date.UTC(<%=CurrentEndTime %>),
			            type: 'datetime',
			            tickInterval: 1800*1000,
						gridLineWidth: 1,
						gridLineColor: 'rgba(255, 255, 255, 0.1)',
			            dateTimeLabelFormats: { // don't display the dummy year
			                hour: '%H:%M'
			            },
						plotLines: [{
										color: '#C0C0C0',
										width: 1,
										value: Date.UTC(<%=CurrentTime %>),
										label: {
											text: '<%=CurrentTimeLabel %>',
											style: {
												color: '#ccc',
												fontWeight: 'bold'
											},
											verticalAlign: 'bottom',
											rotation: 0,
											x: -40,
											y: -5
										}
									}]
			        },
			        yAxis: {
			            min: 0,
			            max: 10,
			            tickInterval: 1,
						plotLines: [{
			                value: 0,
			                width: 1,
			                color: '#808080'
			            }],
			            labels: { 
							useHTML: true,
							formatter: function() 
			            {
			                var t;
			                if (this.value == 0) {t='<% 
							if (ATMID(9,1)<>"") then
								Response.write "("&ATMID(9,1)&") "
							end if
							Response.write ATMID(9,0)
							if (ATMID(9,3)>0) then
								Response.write "<span class=""hc-label"" >("&ATMID(9,3)&")</span>"
							end if
							%>'};
			                if (this.value == 1) {t='<% 
							if (ATMID(8,1)<>"") then
								Response.write "("&ATMID(8,1)&") "
							end if
							Response.write ATMID(8,0)
							if (ATMID(8,3)>0) then
								Response.write "<span class=""hc-label"" >("&ATMID(8,3)&")</span>"
							end if
							%>'};
			                if (this.value == 2) {t='<%
							if (ATMID(7,1)<>"") then
								Response.write "("&ATMID(7,1)&") "
							end if 
							Response.write ATMID(7,0)
							if (ATMID(7,3)>0) then
								Response.write "<span class=""hc-label"" >("&ATMID(7,3)&")</span>"
							end if
							%>'};
			                if (this.value == 3) {t='<% 
							if (ATMID(6,1)<>"") then
								Response.write "("&ATMID(6,1)&") "
							end if 
							Response.write ATMID(6,0)
							if (ATMID(6,3)>0) then
								Response.write "<span class=""hc-label"" >("&ATMID(6,3)&")</span>"
							end if
							%>'};
			                if (this.value == 4) {t='<% 
							if (ATMID(5,1)<>"") then
								Response.write "("&ATMID(5,1)&") "
							end if 
							Response.write ATMID(5,0)
							if (ATMID(5,3)>0) then
								Response.write "<span class=""hc-label"" >("&ATMID(5,3)&")</span>"
							end if
							%>'};
			                if (this.value == 5) {t='<% 
							if (ATMID(4,1)<>"") then
								Response.write "("&ATMID(4,1)&") "
							end if
							Response.write ATMID(4,0)
							if (ATMID(4,3)>0) then
								Response.write "<span class=""hc-label"" >("&ATMID(4,3)&")</span>"
							end if
							%>'};
			                if (this.value == 6) {t='<% 
							if (ATMID(3,1)<>"") then
								Response.write "("&ATMID(3,1)&") "
							end if
							Response.write ATMID(3,0)
							if (ATMID(3,3)>0) then
								Response.write "<span class=""hc-label"" >("&ATMID(3,3)&")</span>"
							end if
							%>'};
			                if (this.value == 7) {t='<% 
							if (ATMID(2,1)<>"") then
								Response.write "("&ATMID(2,1)&") "
							end if
							Response.write ATMID(2,0)
							if (ATMID(2,3)>0) then
								Response.write "<span class=""hc-label"" >("&ATMID(2,3)&")</span>"
							end if
							%>'};
			                if (this.value == 8) {t='<% 
							if (ATMID(1,1)<>"") then
								Response.write "("&ATMID(1,1)&") "
							end if
							Response.write ATMID(1,0)
							if (ATMID(1,3)>0) then
								Response.write "<span class=""hc-label"" >("&ATMID(1,3)&")</span>"
							end if
							%>'};
			                if (this.value == 9) {t='<% 
							if (ATMID(0,1)<>"") then
								Response.write "("&ATMID(0,1)&") "
							end if
							Response.write ATMID(0,0)
							if (ATMID(0,3)>0) then
								Response.write "<span class=""hc-label"" >("&ATMID(0,3)&")</span>"
							end if
							%>'};
			                return t;
			            },
			                style: {color: '#FFFFFF', font: 'bold 20px Arial'},
			                align: 'left',
			                x: -173,
			                y: -15
			            },
						title: { margin: 150, text: ' '},
			            plotLines: [{
			                value: 0,
			                width: 1,
			                color: '#808080'
			            }]
			        },
			        plotOptions: {
			            scatter: {
			                dataLabels: {
			                    enabled: false,
			                    align: 'right',
			                    style: { font: 'bold 24px Arial' },
			                    formatter: function() {	return this.point.name; }
			                },
			                marker: {
			                    enabled: true, 
			                    symbol: 'circle'
			                },
			                enableMouseTracking: false
			            }
						/*series: {
							dataLabels: {
								enabled: true,
								align: 'center',
								style: { font: 'bold 24px Arial' },
								formatter: function() { return this.point.name; }
							}
						}, 
						line: {
							dataLabels: {enabled: false, color: '#FFFF00', y: -8},
							lineWidth: 0,
							marker: { enabled: false },
							enableMouseTracking: false
						},
						scatter: {
							dataLabels: {
								enabled: false,
								align: 'right',
								style: { font: 'bold 14px Arial' },
								formatter: function() {	return this.point.name; }
							},
							marker: {
								enabled: true, 
								symbol: 'circle'
							},
							enableMouseTracking: false
						}*/						
			        },
			        series: [
					 <% =AllSeries %> 
                    ]
            });
			
				// Второй график
			    chart2 = new Highcharts.Chart({
			        chart: {
			            renderTo: 'container2',
						ignoreHiddenSeries : false,
						marginLeft: 178,
			            marginRight: 20
			        },
			        credits: {enabled: false},
			        legend:  {enabled: false},
			        tooltip: {enabled: false},
			        title:   {align: 'right', text: 'Количество сбойных устройств'},
			        xAxis: {
			            min: Date.UTC(<%=CurrentStratTime %>),
			            max: Date.UTC(<%=CurrentEndTime %>),
			            type: 'datetime',
			            tickInterval: 1800*1000,
			            dateTimeLabelFormats: { // don't display the dummy year
			                hour: '%H:%M'
			            },
						plotLines: [{
										color: '#C0C0C0',
										width: 1,
										value: Date.UTC(<%=CurrentTime %>),
										label: {
											text: '<%=CurrentTimeLabel %>',
											style: {
												color: '#ccc',
												fontWeight: 'bold'
											},
											verticalAlign: 'top',
											rotation: 0,
											x: -40,
											y: 15
										}
									}]
			        },
			        yAxis: {
			            min: 0,
			            tickInterval: 1,
			            labels: { 
							enabled: false,
			                verticalAlign: 'middle',
							style: {color: '#FFFFFF', font: 'normal 20px Arial' },
			                align: 'right'
			            },
			            title: { margin: 150, text: ' '},
			            plotLines: [{
			                value: 0,
			                width: 1,
			                color: '#808080'
			            }]
			        },
			        plotOptions: {

						series: {
							dataLabels: {
								enabled: true,
								color: '#C0C0C0', 
								style: { font: 'normal 20px Arial' }
							}
						},

			            column: {
							pointPadding: 0.2,
							borderWidth: 0
						}
					
			        },
			        series: [
					 <% =AllSeries2 %> 
                    ]
				});


				//------------------------------------------------------------
				//--------------Update Series---------------------------------
				$.get('detail3.asp',{todo:'update_series'},	function(data){});
				//------------------------------------------------------------
				//--------------Update Series---------------------------------				

			});

		</script>
		
	<style type="text/css">

		html { overflow: hidden; }
	
	.hc-label {
		  color: #00FFFF;
		}
	<!--
	BODY {
		margin: 0px;
		background-color: #242424;
	}
TABLE {
	margin: 0px;
	padding: 0px;
}
TD {
	padding-top: 1px;
	padding-bottom: 1px;
	text-align: center;
	color: #FFFFFF;
	font-family: Verdana, Arial, helvetica, sans-serif, Geneva;
}
TD.Head {
	color: #000000;
	font-size: 24pt;
}
TD.Txt {
	color: #FFFFFF;
	font-size: 36pt;
	font-weight: 700;
}
	
	-->
	</style>
</head>
<body>
<div align="center">
<table border="0" padding="0"  width="1920px" height="1080px" style="border: none;">
	<tr>
		<td style="border: none;" colspan="4" align="center" ><div id="containerTitle"  style="width: 100%; height: 24px;font-size: 20pt; font-weight: 400;"><%=VIP_Title %></div></td>
	</tr>
	<tr>
		<td style="border: none; vertical-align: top;"><div id="circleIndicator"  style="width: 300px; height: 300px; margin: 0 auto;"></div>
		<p  style="font-size: 20pt;" ><%=LastFileTimeFull %></p>
		</td>
		<td style="border: none;">
		  <div id="containerAA" style="width: 210px; height: 500px; margin: 0 auto; font-size: 22pt;">
		    <table border="0" height="100%" width="100%" cellspacing="0" >
			  <tr><td style="border-top: solid 1px #C0C0C0; border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center">Всего:</td></tr>
			  <tr><td style="border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center; font-size: 40pt; font-weight: 600;"><% =ValueAll %></td></tr>
			  <tr><td style="border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0;  text-align: center; font-size: 20pt; font-weight: 400;"><% =ValueAll_discr %></td></tr>
              <tr><td style="border-top: solid 1px #C0C0C0; border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center" nowrap>Сбойные:</td></tr>
			  <tr><td style="border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center; font-size: 40pt; font-weight: 600;"><% =ValueFail %></td></tr>
			  <tr><td style="border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center; font-size: 20pt; font-weight: 400;"><% =ValueFail_discr %></td></tr>
              <tr><td style="border-top: solid 1px #C0C0C0; border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center" nowrap>На контроле:</td></tr>
			  <tr><td style="border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0;  text-align: center; font-size: 40pt; font-weight: 600;"><% =ValueControl %></td></tr>
			  <tr><td style="border-bottom: solid 1px #C0C0C0; border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center; font-size: 20pt; font-weight: 400;"><% =ValueControl_discr %></td></tr>
			</table>
		  </div></td>
		<td style="border: none;" colspan="2"><div id="container1"  style="width: 1396px; height: 500px; margin: 0 auto"></div></td>
	</tr>
	<tr>
		<td style="border: none;"  colspan="2" >
			<div id="containerBB"  style="width: 510px; height: 294px; margin: 0 auto; font-size: 22pt;">
				<table border="0" height="100%" width="100%" cellspacing="0" >
				  <tr><td style="border-top: solid 1px #C0C0C0; border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center; font-size: 20pt; font-weight: 400;">НЕИСПРАВНОСТЬ</td>
					<td style="border-top: solid 1px #C0C0C0;" ><% =LastFileTime %></td>
					<td style="border-top: solid 1px #C0C0C0; border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0;" >ВСЕГО</td></tr>
				  <tr><td style="border-left: solid 1px #C0C0C0; border-top: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center; font-size: 20pt; font-weight: 400; color: #66FF66;">ПО СВЯЗИ</td>
					<td style="border-top: solid 1px #C0C0C0;" ><% =ValueFail_link %></td>
					<td rowspan=4 style="border-top: solid 1px #C0C0C0; border-bottom: solid 1px #C0C0C0; border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0;" ><% =ValueFail_noUnion %></td></tr>
				  <tr><td style="border-left: solid 1px #C0C0C0; border-top: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center; font-size: 20pt; font-weight: 400; color: #FFFF66;">ПО ДЕНЬГАМ</td>
					<td style="border-top: solid 1px #C0C0C0;" ><% =ValueFail_money %></td></tr>
				  <tr><td style="border-left: solid 1px #C0C0C0; border-top: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center; font-size: 20pt; font-weight: 400; color: #66FFFF;">ПО ОБОРУДОВАНИЮ</td>
					<td style="border-top: solid 1px #C0C0C0;" ><% =ValueFail_hardware %></td></tr>
				  <tr><td style="border-left: solid 1px #C0C0C0; border-top: solid 1px #C0C0C0; border-bottom: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center; font-size: 20pt; font-weight: 400; color: #FF66FF;">НЕ ОБСЛУЖИВАЕТ</td>
					<td  style="border-top: solid 1px #C0C0C0; border-bottom: solid 1px #C0C0C0;" ><% =ValueFail_service %></td></tr>
				<tr><td colspan=3 style="border: 0; text-align: left; font-size: 20pt; font-weight: 400; color: #FFFFFF;">TOP5</td></tr>
				</table>
			</div>
		</td>
		<td style="border: none;" colspan="2"><div id="container2"  style="width: 1396px; height: 294px; margin: 0 auto"></div></td>
	</tr>
	<tr>
		<td style="border: none;" colspan="4">
			<div id="containerC1"  style="float: left; width: 382px; height: 252px; margin: 0 auto">
			<%= TOP5TABLES(0) %>
			</div>
			<div id="containerC2"  style="float: left; width: 382px; height: 252px; margin: 0 auto">
			<%= TOP5TABLES(1) %>
			</div>
			<div id="containerC3"  style="float: left; width: 384px; height: 252px; margin: 0 auto">
			<%= TOP5TABLES(2) %>
			</div>
			<div id="containerC4"  style="float: left; width: 382px; height: 252px; margin: 0 auto">
			<%= TOP5TABLES(3) %>
			</div>
			<div id="containerC5"  style="float: left; width: 382px; height: 252px; margin: 0 auto">
			<%= TOP5TABLES(4) %>
			</div>
		</td>

	</tr>
</table>
</div>
</body>
</html>
<%

  Conn.Close
  set Conn = Nothing
  set Rs = Nothing

end if
%>
<!-- Разработка: Берников И.П. -->
<!-- Для вывода графиков используется библиотека Highcharts JS - http://highsoft.com/ -->
