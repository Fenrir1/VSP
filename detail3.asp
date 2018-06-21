<%@ language="VBScript"%><!-- #include file="const.asp" --><!-- #include file="common.asp" -->
<%
' экран детализация мониторинга VIP банкоматов БПТ

Response.Charset = "windows-1251"

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


set Rs=Server.CreateObject("ADODB.Recordset")

Function URLDecode(sConvert)
    Dim aSplit
    Dim sOutput
 	if (sConvert<>"") then
     aSplit = Split(sConvert, ";")

     If IsArray(aSplit) Then
      for I = 0 to UBound(aSplit) - 1
        'добавил проверку на символ №
        if (aSplit(i)="2116") then
            sOutput = sOutput & "№"
        else
	        sOutput = sOutput & Chr("&H"&aSplit(i))
        end if
	  Next
     End If
    end if
    URLDecode = sOutput
End Function

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


dim series(20), ATMID(20,4), series_warn(20), ATMID_warn(20,4)
VDCount=20
VDCount_warn=20
for i=0 to 20 
  series(i)=""
  ATMID(i,0)="" ' NumEmv
  ATMID(i,1)=""	' DeviceId
  ATMID(i,2)=""	' DeviceType
  ATMID(i,3)=""	' DaysCount
  
  'Wirnings Chart
  series_warn(i)=""
  ATMID_warn(i,0)="" ' NumEmv
  ATMID_warn(i,1)=""	' DeviceId
  ATMID_warn(i,2)=""	' DeviceType
  ATMID_warn(i,3)=""	' DaysCount
next


'-----------------------------------------------------------------------
'--START:  Chart1-------------------------------------------------------
'-----------------------------------------------------------------------
    ATMIDList = ""

	sqlstr = "select vn.NumEmv, vd.DeviceType, vd.DeviceID, ISNULL(DaysCount,0) DaysCount from VIP_Day_Order vd join VIP_Device vn on vn.DeviceType=vd.DeviceType and vn.DeviceId=vd.DeviceId "
	'sqlstr = sqlstr&" where exists(select * from VIP_Intervals_Series where "
	'sqlstr = sqlstr&" DeviceType=vd.DeviceType and DeviceId=vd.DeviceId and [Time]>=convert(datetime,FLOOR(convert(float,GETDATE())))) "
	sqlstr = sqlstr&" order by OrderNum "
	
	'sqlstr = sqlstr&" from VIP_Day_Order  vdo join VIP_Device vd on vd.DeviceID=vdo.DeviceID and vd.DeviceType=vdo.DeviceType order by OrderNum "
    count = 0
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        do while (not Rs.EOF)
            found = 0
            for i=0 to 20 
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
	
	'sqlstr = sqlstr&" from VIP_Day_Order  vdo join VIP_Device vd on vd.DeviceID=vdo.DeviceID and vd.DeviceType=vdo.DeviceType order by OrderNum "
    'count = 0
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        do while (not Rs.EOF)
            found = 0
            for i=0 to 20 
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
	
	VDCount = count

	'------------------------------------
	'--START: Series for Chart1----------
	'------------------------------------
	for i=0 to VDCount 
	  if (ATMID(i,0)="") then
		exit for
	  end if
	  
	  serieID = i
	  'v=VDCount-0.5-serieID
	  v=VDCount-1-serieID
	  v=replace(v, ",", ".")
	  
	  sqlstr = "select DeviceID,DeviceType,dateadd(month,-1,[TIME]) [TIME] from VIP_Intervals_Series where DeviceID="&ATMID(i,2)&" and DeviceType='"&ATMID(i,1)&"' "
	  sqlstr = sqlstr&" and [TIME]>=convert(datetime,FLOOR(convert(float,Getdate()))) order by [TIME]"
	  Rs.Open sqlstr, Conn
	  If not Rs.EOF then
		do while (not Rs.EOF)
			'series(serieID) = series(serieID)&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("Start_time"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"},"
			series(serieID)=series(serieID)&vbCrLf&"{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 2}, x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"},"
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
for i=0 to VDCount
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
 	AllSeries = AllSeries & " x: Date.UTC("&DateTimeFormat(Now, "yyyy, mm, dd, hh, nn")&"), y: -1 } ]} "
end if
'-----------------------------------------------------------------------
'--END:  Chart1---------------------------------------------------------
'-----------------------------------------------------------------------

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

LastFileTime = DateTimeFormat(Now,"hh:nn")
ValueFail_noUnion = 0
ValueFail_link = 0
ValueFail_money = 0
ValueFail_hardware = 0
ValueFail_service = 0


sqlstr = "select DeviceType,count(DeviceId) DeviceCount from VIP_Device group by DeviceType"
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

sqlstr = "select DeviceType,count(DeviceId) DeviceCount from VIP_Intervals_Union where Closed=0 group by DeviceType"
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

sqlstr = "select DeviceType,count(DeviceId) DeviceCount from VIP_Intervals_Union where Closed=0 and ISNULL(onControl,0)=1 group by DeviceType"
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

'-----------------------------------------------------------------------
'--START:  Chart2-------------------------------------------------------
'-----------------------------------------------------------------------
    ATMIDList = ""
	sqlstr = "select NumEmv, DeviceType, DeviceID, ISNULL(DaysCount,0) DaysCount from VIP_Device vd where exists(select * from VIP_Warnings_Series where "
	sqlstr = sqlstr&" DeviceType=vd.DeviceType and DeviceId=vd.DeviceId and [Time]>=convert(datetime,FLOOR(convert(float,GETDATE())))) "
	sqlstr = sqlstr&" order by DeviceType, DeviceID "

    count = 0
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        do while (not Rs.EOF)
            found = 0
            for i=0 to 10 
              if ((ATMID_warn(i,0)=Rs.Fields("NumEmv"))and(ATMID_warn(i,1)=Rs.Fields("DeviceType"))) then
                found = 1
                exit for
              end if
            next
            if (found=0) then
                ATMID_warn(count,0)=Rs.Fields("NumEmv")
				ATMID_warn(count,1)=Rs.Fields("DeviceType")
				ATMID_warn(count,2)=Rs.Fields("DeviceID")
				ATMID_warn(count,3)=Rs.Fields("DaysCount")
                if (ATMIDList<>"") then ATMIDList = ATMIDList&"," end if
                ATMIDList = ATMIDList&Rs.Fields("DeviceID")
                count = count+1
            end if
            Rs.MoveNext
        loop
    end if
    Rs.Close
	
	VDCount_warn = count

	'------------------------------------
	'--START: Series for Chart2----------
	'------------------------------------
	for i=0 to VDCount_warn 
	  if (ATMID_warn(i,0)="") then
		exit for
	  end if
	  
	  serieID = i
	  'v=VDCount_warn-0.5-serieID
	  v=VDCount_warn-1-serieID
	  v=replace(v, ",", ".")
	  
	  sqlstr = "select DeviceID,DeviceType,dateadd(month,-1,[TIME]) [TIME] from VIP_Warnings_Series where DeviceID="&ATMID_warn(i,2)&" and DeviceType='"&ATMID_warn(i,1)&"' "
	  sqlstr = sqlstr&" and [TIME]>=convert(datetime,FLOOR(convert(float,Getdate()))) order by [TIME]"
	  Rs.Open sqlstr, Conn
	  If not Rs.EOF then
		do while (not Rs.EOF)
			'series(serieID) = series(serieID)&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("Start_time"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"},"
			series_warn(serieID)=series_warn(serieID)&vbCrLf&"{color: null, marker: {fillColor: '#99ccff', lineColor: '#99ccff', radius: 2}, x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"},"
			Rs.MoveNext
        loop
	  end if
	  Rs.Close
  
	next
	'------------------------------------
	'--END: Series for Chart2------------
	'------------------------------------

AllSeries_warn=""
for i=0 to VDCount_warn
  if series_warn(i)<>"" then
    series_warn(i)=left(series_warn(i), len(series_warn(i))-1)
    'scatter
	if (i>0) then
		series_warn(i)=", { name: '"&ATMID_warn(i,0)&"', type: 'scatter', data: ["&series_warn(i)&"]}"  
	else 
		series_warn(i)=" { name: '"&ATMID_warn(i,0)&"', type: 'scatter', data: ["&series_warn(i)&"]}" 
	end if 
	'CID(i)=left(CID(i), 12)
  end if
  AllSeries_warn=AllSeries_warn+series_warn(i)
next

if (AllSeries_warn="") then
	AllSeries_warn = AllSeries_warn & "{ name: '00', type: 'scatter', visible: false, data: [ "
	AllSeries_warn = AllSeries_warn & " {color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 2}, "
 	AllSeries_warn = AllSeries_warn & " x: Date.UTC("&DateTimeFormat(Now, "yyyy, mm, dd, hh, nn")&"), y: -1 } ]} "
end if
'-----------------------------------------------------------------------
'--END:  Chart2---------------------------------------------------------
'-----------------------------------------------------------------------


'------------------------------------
'--START: Data Table2----------------
'------------------------------------
ValueWarining = 0
ValueWariningATM = 0
ValueWariningBPT = 0
ValueWarining_discr = "0 ATM; 0 BPT"
ValueControlWarning = 0
ValueControlWarningATM = 0
ValueControlWarningBPT = 0
ValueControlWarning_discr = "0 ATM; 0 BPT"

sqlstr = "select DeviceType,count(distinct DeviceId) DeviceCount from VIP_Warnings_Union where Closed=0 group by DeviceType"
Rs.Open sqlstr, Conn
If not Rs.EOF then
do while (not Rs.EOF)
	if (Rs.Fields("DeviceType")="ATM") then
		ValueWariningATM = Rs.Fields("DeviceCount")
	end if
	
	if (Rs.Fields("DeviceType")="BPT") then
		ValueWariningBPT = Rs.Fields("DeviceCount")
	end if
	Rs.MoveNext
loop
end if
Rs.Close

ValueWarining = ValueWariningATM+ValueWariningBPT
ValueWarining_discr = ValueWariningATM&" ATM; "&ValueWariningBPT&" BPT"

sqlstr = "select DeviceType,count(distinct DeviceId) DeviceCount from VIP_Warnings_Union where Closed=0 and ISNULL(onControl,0)=1 group by DeviceType"
Rs.Open sqlstr, Conn
If not Rs.EOF then
do while (not Rs.EOF)
	if (Rs.Fields("DeviceType")="ATM") then
		ValueControlWarningATM = Rs.Fields("DeviceCount")
	end if
	
	if (Rs.Fields("DeviceType")="BPT") then
		ValueControlWarningBPT = Rs.Fields("DeviceCount")
	end if
	Rs.MoveNext
loop
end if
Rs.Close

ValueControlWarning = ValueControlWarningATM+ValueControlWarningBPT
ValueControlWarning_discr = ValueControlWarningATM&" ATM; "&ValueControlWarningBPT&" BPT"


'------------------------------------
'--END: Data Table2------------------
'------------------------------------

'------------------------------------
'--START: Link to QOS----------------
'------------------------------------	
QOSLink = ""
sqlstr = "select IsNUll(QOSLink,'') QOSLink from VIP_Config"
Rs.Open sqlstr, Conn
If not Rs.EOF then
	QOSLink = Rs.Fields("QOSLink")
end if
Rs.Close
'------------------------------------
'--END: Link to QOS----------------
'------------------------------------	

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------

function get_interval( d_id , d_type )

	sqlstr = "set dateformat dmy; select vd.*,datediff(minute,viu.StartTime,ISNULL(EndTime,GETDATE())) IntervalLength, viu.StartTime, viu.Comment1, viu.Comment2 from VIP_Device vd join "
	sqlstr = sqlstr&" (select * from VIP_Intervals_Union where '"&Request("timePoint")&"' between StartTime and ISNULL(EndTime,GETDATE()) ) viu on "
	sqlstr = sqlstr&" vd.DeviceId=viu.DeviceID and vd.DeviceType=viu.DeviceType "
	sqlstr = sqlstr&" where vd.DeviceId="&Request("deviceId")&" and vd.DeviceType='"&Request("deviceType")&"' "
	'response.write sqlstr
	Rs.Open sqlstr, Conn
	If not Rs.EOF then
		response.write "<table width=""85%"" cellspacing=""0"" >"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >ID</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("DeviceType")&" "&Rs.Fields("NumEmv")&"</td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Серийный номер</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("SERIAL_NUMBER")&"</td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Модель</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("Type")&"</td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Город</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("City")&"</td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Адрес</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("Address")&"</td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Точка</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("Location")&"</td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Длительность</td><td style='border: solid 1px #C0C0C0;' >"&parseInterval(Rs.Fields("IntervalLength"))&"</td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Причина</td><td style='border: solid 1px #C0C0C0;' >"
		response.write "<input style=""width: 500px"" id='comm1' value='"&Rs.Fields("Comment1")&"' ></td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Сроки</td><td style='border: solid 1px #C0C0C0;' >"
		response.write "<input style=""width: 500px"" id='comm2' value='"&Rs.Fields("Comment2")&"' ></td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' colspan=2 ><button onClick=""saveComments(1,"&Request("deviceId")&",'"&Request("deviceType")&"','"&Rs.Fields("StartTime")&"')"" >Сохранить комментарии</button></td></tr>"
		IntervalStart = Rs.Fields("StartTime")
		Rs.Close
		
		sqlstr = "set dateformat dmy; select * from VIP_Intervals where DeviceId="&Request("deviceId")&" and DeviceType='"&Request("deviceType")&"' "
		sqlstr = sqlstr&" and StartTime>='"&DateTimeFormat(IntervalStart, "dd.mm.yy hh:nn")&"' "
		Rs.Open sqlstr, Conn
		If not Rs.EOF then
			do while (not Rs.EOF)
				tempStat = Rs.Fields("Class")
				if (tempStat="Money") then
					tempStat="по деньгам"
				elseif (tempStat="Link") then
					tempStat="по связи"
				elseif (tempStat="Hardware") then
					tempStat="по оборудованию"
				elseif (tempStat="Service") then
					tempStat="не обслуживает"
				end if
				response.write "<tr><td style='border: solid 1px #C0C0C0;' >Дата/Время начала</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("StartTime")&"</td></tr>"
				response.write "<tr><td style='border: solid 1px #C0C0C0;' >Статус</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("Status")&"</td></tr>"
				response.write "<tr><td style='border: solid 1px #C0C0C0;' >Неисправность</td><td style='border: solid 1px #C0C0C0;' >"&tempStat&"</td></tr>"
				response.write "<tr><td style='border: solid 1px #C0C0C0;' >Модуль</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("Hardware")&"</td></tr>"
				response.write "<tr><td style='border: solid 1px #C0C0C0;' >Сообщение</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("Name")&"</td></tr>"

				Rs.MoveNext
			loop
		end if
		response.write "</table>"
		
	end if
	Rs.Close
	
end function

function get_warning( d_id , d_type )

	sqlstr = "set dateformat dmy; select vd.*,datediff(minute,viu.StartTime,ISNULL(EndTime,GETDATE())) IntervalLength, viu.StartTime, viu.Comment1, viu.Comment2 from VIP_Device vd join "
	sqlstr = sqlstr&" (select * from VIP_Warnings_Union where  '"&Request("timePoint")&"' between StartTime and ISNULL(EndTime,GETDATE()) ) viu on "
	sqlstr = sqlstr&" vd.DeviceId=viu.DeviceID and vd.DeviceType=viu.DeviceType "
	sqlstr = sqlstr&" where vd.DeviceId="&Request("deviceId")&" and vd.DeviceType='"&Request("deviceType")&"' "
	'response.write sqlstr
	Rs.Open sqlstr, Conn
	If not Rs.EOF then
		response.write "<table width=""100%"" cellspacing=""0"" >"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >ID</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("DeviceType")&" "&Rs.Fields("NumEmv")&"</td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Серийный номер</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("SERIAL_NUMBER")&"</td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Модель</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("Type")&"</td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Город</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("City")&"</td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Адрес</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("Address")&"</td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Точка</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("Location")&"</td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Длительность</td><td style='border: solid 1px #C0C0C0;' >"&parseInterval(Rs.Fields("IntervalLength"))&"</td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Причина</td><td style='border: solid 1px #C0C0C0;' >"
		response.write "<input style=""width: 500px"" id='comm1' value='"&Rs.Fields("Comment1")&"' ></td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0;' >Сроки</td><td style='border: solid 1px #C0C0C0;' >"
		response.write "<input style=""width: 500px"" id='comm2' value='"&Rs.Fields("Comment2")&"' ></td></tr>"
		response.write "<tr><td style='border: solid 1px #C0C0C0; border-right: 0;' ></td><td style='border: solid 1px #C0C0C0; border-left: 0;'  ><button onClick=""saveComments(2,"&Request("deviceId")&",'"&Request("deviceType")&"','"&Rs.Fields("StartTime")&"')"" >Сохранить комментарии</button></td></tr>"
		IntervalStart = Rs.Fields("StartTime")
		Rs.Close
		
		sqlstr = "set dateformat dmy; select * from VIP_Warnings where DeviceId="&Request("deviceId")&" and DeviceType='"&Request("deviceType")&"' "
		sqlstr = sqlstr&" and StartTime>='"&DateTimeFormat(IntervalStart, "dd.mm.yy hh:nn")&"' "
		Rs.Open sqlstr, Conn
		If not Rs.EOF then
			do while (not Rs.EOF)
				response.write "<tr><td style='border: solid 1px #C0C0C0;' >Дата/Время начала</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("StartTime")&"</td></tr>"
				response.write "<tr><td style='border: solid 1px #C0C0C0;' >Статус</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("Status")&"</td></tr>"
				'response.write "<tr><td style='border: solid 1px #C0C0C0;' >Неисправность</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("Class")&"</td></tr>"
				response.write "<tr><td style='border: solid 1px #C0C0C0;' >Модуль</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("Hardware")&"</td></tr>"
				response.write "<tr><td style='border: solid 1px #C0C0C0;' >Сообщение</td><td style='border: solid 1px #C0C0C0;' >"&Rs.Fields("Name")&"</td></tr>"

				Rs.MoveNext
			loop
		end if
		response.write "</table>"
		
	end if
	Rs.Close
	
end function

function save_comments(tableType, d_id, d_type, startTime, comm1, comm2)
	if (cInt(tableType)=1) then
		'URLDecode
		onControl = 0
		if ((comm1<>"")and(comm2<>"")) then
			onControl =1
		end if
		sqlstr = "set dateformat dmy; UPDATE VIP_Intervals_Union set Comment1='"&URLDecode(comm1)&"', Comment2='"&URLDecode(comm2)&"', onControl="&onControl&" "
		sqlstr = sqlstr&" WHERE  DeviceId="&d_id&" and DeviceType='"&d_type&"' and StartTime='"&DateTimeFormat(startTime, "dd.mm.yy hh:nn:ss")&"' "
		'response.write sqlstr
		Rs.Open sqlstr, Conn
	elseif (cInt(tableType)=2) then
		onControl = 0
		if ((comm1<>"")and(comm2<>"")) then
			onControl =1
		end if
		sqlstr = "set dateformat dmy; UPDATE VIP_Warnings_Union set Comment1='"&URLDecode(comm1)&"', Comment2='"&URLDecode(comm2)&"', onControl="&onControl&" "
		sqlstr = sqlstr&" WHERE  DeviceId="&d_id&" and DeviceType='"&d_type&"' and StartTime='"&DateTimeFormat(startTime, "dd.mm.yy hh:nn:ss")&"' "
		Rs.Open sqlstr, Conn
	end if
end function

'--------------------------------------------------------------------------------
'------START: Auto update series-------------------------------------------------
'--------------------------------------------------------------------------------
function update_series
	tempTime = Now
	sqlstr = "set dateformat dmy; exec sp_Update_Interwal_Series @Date='"&DateTimeFormat(tempTime, "dd.mm.yy hh:nn")&"'"
	Rs.Open sqlstr, Conn

	sqlstr = "set dateformat dmy; exec sp_Update_Warning_Series @Date='"&DateTimeFormat(tempTime, "dd.mm.yy hh:nn")&"'"
	Rs.Open sqlstr, Conn
end function
'--------------------------------------------------------------------------------
'------END: Auto update series---------------------------------------------------
'--------------------------------------------------------------------------------


if NOT IsEmpty(Request("todo")) then
	if Request("todo") = "get_interval" then
		get_interval Request("deviceId"), Request("deviceType")
	elseif Request("todo") = "get_warning" then
		get_warning Request("deviceId"), Request("deviceType")
	elseif Request("todo") = "save_comments" then
		save_comments Request("tableType"), Request("deviceId"), Request("deviceType"), Request("startTime"), Request("comm1"), Request("comm2")
	elseif Request("todo") = "update_series" then
		update_series
	end if
	Response.End
end if

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------


CurrentStratTime = DateTimeFormat(DateAdd("m", -1, Now), "yyyy, mm, dd, 0, 0")
CurrentTime = DateTimeFormat(DateAdd("m", -1, Now), "yyyy, mm, dd, hh, nn")
CurrentEndTime = DateTimeFormat(DateAdd("m", -1, Now), "yyyy, mm, dd, 23, 59")
%>
<!DOCTYPE HTML>
<html>
<head>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
		<!-- 1. Add these JavaScript inclusions in the head of your page -->
		<script type="text/javascript" src="js/jquery.min.js"></script>
		<!-- <script type="text/javascript" src="js/highcharts.js"></script> -->
		<script type="text/javascript" src="js/highstock.js"></script>
		
		<script type="text/javascript" src="js/themes/gray.js"></script>
		
		<!-- 2. Add the JavaScript to initialize the chart on document ready -->

	<style type="text/css">
	<!--
BODY {
	margin: 0px;
	background-color: #242424;
}
TABLE {
	margin-bottom: 0px;
	margin-top: 0px;
}
TD {
	padding-top: 1px;
	padding-bottom: 1px;
	padding-left: 2px;
	padding-right: 2px;
	text-align: center;
	color: #FFFFFF;
	font-family: Verdana, Arial, helvetica, sans-serif, Geneva;
}

TH.A {
	color: #FFFFFF;
	background-color: #6F8CBF;
	padding-left: 2px;
	padding-right: 2px;
	text-align: left;
}

-->
	</style>
		<script type="text/javascript">
		
			var chart1;
			var chart2;
			var chart3;
			var chartA;
			var chartB;
			var chartC;
			var FlagOut=1;
			
			var now = new Date(); 
			var now_utc =  Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(),  now.getHours(), now.getMinutes(), now.getSeconds());			
	
			var DeviceIntervals = [<%
			for i=0 to (VDCount-1)
				if (i>0) then
					Response.write ",{ 'NumEmv': '"&ATMID(VDCount-i-1,0)&"', 'DeviceId': "&ATMID(VDCount-i-1,2)&",  'DeviceType': '"&ATMID(VDCount-i-1,1)&"'}"
				else
					Response.write "{ 'NumEmv': '"&ATMID(VDCount-i-1,0)&"', 'DeviceId': "&ATMID(VDCount-i-1,2)&",  'DeviceType': '"&ATMID(VDCount-i-1,1)&"'}"
				end if
			next
%>];
			var DeviceWarnings = [<%
			for i=0 to (VDCount_warn-1)
				if (i>0) then
					Response.write ",{ 'NumEmv': '"&ATMID_warn(VDCount_warn-i-1,0)&"', 'DeviceId': "&ATMID_warn(VDCount_warn-i-1,2)&",  'DeviceType': '"&ATMID_warn(VDCount_warn-i-1,1)&"'}"
				else
					Response.write "{ 'NumEmv': '"&ATMID_warn(VDCount_warn-i-1,0)&"', 'DeviceId': "&ATMID_warn(VDCount_warn-i-1,2)&",  'DeviceType': '"&ATMID_warn(VDCount_warn-i-1,1)&"'}"
				end if
			next
%>];
			

    /*-------------------------------------------------------------------------------------------*/
    /*-------------START: Convrte text to HEX----------------------------------------------------*/
    /*-------------------------------------------------------------------------------------------*/
    function dec2hex(textString) {
        return (textString + 0).toString(16).toUpperCase();
    }

    function converterhex(text) {
        var charmap1 = ["а", "б", "в", "г", "д", "е", "ё", "ж", "з", "и", "й", "к", "л", "м", "н", "о", "п", "р", "с", "т", "у", "ф", "х", "ц", "ч", "ш", "щ", "ъ", "ы", "ь", "э", "ю", "я"];
        var charmap1b = ["А", "Б", "В", "Г", "Д", "Е", "Ё", "Ж", "З", "И", "Й", "К", "Л", "М", "Н", "О", "П", "Р", "С", "Т", "У", "Ф", "Х", "Ц", "Ч", "Ш", "Щ", "Ъ", "Ы", "Ь", "Э", "Ю", "Я"];
        var charmap2 = ["E0", "E1", "E2", "E3", "E4", "E5", "B8", "E6", "E7", "E8", "E9", "EA", "EB", "EC", "ED", "EE", "EF", "F0", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "FA", "FB", "FC", "FD", "FE", "FF"];
        var charmap2b = ["C0", "C1", "C2", "C3", "C4", "C5", "A8", "C6", "C7", "C8", "C9", "CA", "CB", "CC", "CD", "CE", "CF", "D0", "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "DA", "DB", "DC", "DD", "DE", "DF"];
        var Res = "";
        var t = 0;
        for (var i = 0; i < text.length; i++) {
            t = 0;
            for (var j = 0; j <= 32; j++) {
                if (text.charAt(i) == charmap1[j].charAt(0)) { Res = Res + charmap2[j] + ";"; t = 1; break; }
                if (text.charAt(i) == charmap1b[j].charAt(0)) { Res = Res + charmap2b[j] + ";"; t = 1; break; }
            }
            if (t == 0) Res = Res + dec2hex(text.charCodeAt(i)) + ";";  //--если введен символ, а не буква русского языка
        }
        return Res
    }
    /*-------------------------------------------------------------------------------------------*/
    /*-------------END: Convrte text to HEX------------------------------------------------------*/
    /*-------------------------------------------------------------------------------------------*/			
	
	/*-------------------------------------------------------------------------------------------*/
    /*-------------START: Select Device----------------------------------------------------------*/
    /*-------------------------------------------------------------------------------------------*/
	function selectDevice(chartNum,idInChart,timePoint) {
		if (chartNum==1) {
			var id = idInChart;
			//id = id-0.5;
			r = Math.random(); 
			$.get('detail3.asp',{todo:'get_interval', 
								deviceId:DeviceIntervals[id].DeviceId, 
								deviceType:DeviceIntervals[id].DeviceType,
								timePoint: timePoint,
								r:r},
				function(data){ $('#containerCC').html(data); });
			//alert(DeviceIntervals[id].NumEmv);
		}
		if (chartNum==2) {
			var id = idInChart;
			//id = id-0.5;
			r = Math.random(); 
			$.get('detail3.asp',{todo:'get_warning', 
								deviceId:DeviceWarnings[id].DeviceId, 
								deviceType:DeviceWarnings[id].DeviceType,
								timePoint: timePoint,
								r:r},
				function(data){ $('#containerCC').html(data); });
		}
	}
	/*-------------------------------------------------------------------------------------------*/
	/*-------------END: Select Device------------------------------------------------------------*/
	/*-------------------------------------------------------------------------------------------*/			
			
			function saveComments(tableNum, deviceId, deviceType, startTime) {
			//save_comments Request("tableType"), Request("deviceId"), Request("deviceType"), Request("startTime"), Request("comm1"), Request("comm2")
				var comm1 = $('#comm1').val();
				//if (comm1.Length>0) { comm1 = converterhex(comm1); }
				var comm2 = $('#comm2').val();
				//if (comm2.Length>0) { comm2 = converterhex(comm2); }
				r = Math.random(); 
				$.get('detail3.asp',{todo:'save_comments', 
									tableType: tableNum,
									deviceId: deviceId, 
									deviceType: deviceType,
									startTime: startTime,
									comm1: converterhex(comm1),
									comm2: converterhex(comm2),
									r:r},
					function(data){ /*location.reload();*/ alert('Ok'); });
			}
						

			$(document).ready(function() {
				
				
				// Первый график
			    chart1 = new Highcharts.Chart({
			        chart: {
			            renderTo: 'container1',
						zoomType: 'x',
						ignoreHiddenSeries : false,
						marginLeft: 130,
			            marginRight: 20
			        },
			        credits: {enabled: false},
			        legend:  {enabled: false},
			        tooltip: {enabled: false},
			        title:   {align: 'right', text: 'Интервалы неработоспособности'},
			        subtitle: {align: 'left', text: 'ID'},
			        xAxis: {
			            min: Date.UTC(<%=CurrentStratTime %>),
			            max: Date.UTC(<%=CurrentEndTime %>),
			            type: 'datetime',
			            tickInterval: 3600*1000,
						labels: {
							autoRotation: 0,
							format: '{value:%H:%M}',
							style: {color: '#FFFFFF', font: 'normal 12px Arial' }
						},
			            dateTimeLabelFormats: { // don't display the dummy year
			                hour: '%H:%M'
			            },
						plotLines: [{
										color: '#C0C0C0',
										width: 1,
										value: Date.UTC(<%=CurrentTime %>)
									}]
						
			        },
			        yAxis: {
			            min: -1,
			            max: <% if (VDCount<10) then
									Response.write VDCount
								else 
									response.write 10
								end if%>,
			            tickInterval: 1,
						scrollbar: {
							enabled: true,
							showFull: false
						},
			            labels: { formatter: function() 
			            {
			                var t;
<%
							for i=0 to (VDCount-1)
							  'Response.write "if (this.value == "&i&") { t='"&ATMID(VDCount-i-1,0)&"' };"
								if (ATMID(VDCount-i-1,1)<>"") then
									Response.write "if (this.value == "&i&") { t='("&ATMID(VDCount-i-1,1)&") "&ATMID(VDCount-i-1,0)
									if (ATMID(VDCount-i-1,3)>0) then
										Response.write "<span class=""hc-label"" >("&ATMID(VDCount-i-1,3)&")</span>"
									end if
									Response.write "' };"
								end if

							  'if (ATMID(VDCount-i-1,3)>0) then
							  '	Response.write "if (this.value == "&i&") { t='"&ATMID(VDCount-i-1,0)&"<span class=""hc-label"" >("&ATMID(VDCount-i-1,3)&")</span>' };"
							  'else
							  '  Response.write "if (this.value == "&i&") { t='"&ATMID(VDCount-i-1,0)&"' };"
							  'end if
							next
%>							
			                return t;
			            },
							verticalAlign: 'middle',
							style: {color: '#FFFFFF', font: 'normal 14px Arial' },
			                align: 'left',							
			                x: -120
			            },
			            title: { margin: 100, text: ' '},
			            plotLines: [{
			                value: 0,
			                width: 1,
			                color: '#808080'
			            }]
			        },
			        plotOptions: {
						series: {
							cursor: 'pointer',
							states: {
								hover: {
									enabled: false,
									halo: {
										size: 0
									}
								}
							},
							point: {
								events: {
									click: function (e) { 
										selectDevice(1,this.y,Highcharts.dateFormat('%d.%m.%Y %H:%M', this.x));
										//alert(this.y);
									}
								}
							},
							marker: {
								enabled: true, 
			                    symbol: 'circle'
							}
						}					
			        },
			        series: [
					 <% =AllSeries %> 
                    ]
				});
			
				// Второй график
			    chart2 = new Highcharts.Chart({
			        chart: {
			            renderTo: 'container2',
						zoomType: 'x',
						ignoreHiddenSeries : false,
						marginLeft: 130,
			            marginRight: 20
			        },
			        credits: {enabled: false},
			        legend:  {enabled: false},
			        tooltip: {enabled: false},
			        title:   {align: 'right', text: 'Интервалы предупреждений'},
			        subtitle: {align: 'left', text: 'ID'},
			        xAxis: {
			            min: Date.UTC(<%=CurrentStratTime %>),
			            max: Date.UTC(<%=CurrentEndTime %>),
			            type: 'datetime',
			            tickInterval: 3600*1000,
						labels: {
							autoRotation: 0,
							format: '{value:%H:%M}',
							style: {color: '#FFFFFF', font: 'normal 12px Arial' }
						},
			            dateTimeLabelFormats: { // don't display the dummy year
			                hour: '%H:%M'
			            },
						plotLines: [{
										color: '#C0C0C0',
										width: 1,
										value: Date.UTC(<%=CurrentTime %>)
									}]
			        },
			        yAxis: {
			            min: -1,
			            max: <% if (VDCount_warn<10) then
									Response.write VDCount_warn
								else 
									response.write 10
								end if%>,
			            tickInterval: 1,
						scrollbar: {
							enabled: true,
							showFull: false
						},
			            labels: { formatter: function() 
			            {
			                var t;
<%
							for i=0 to (VDCount_warn-1)
							  Response.write "if (this.value == "&i&") {t='("&ATMID_warn(VDCount_warn-i-1,1)&") "&ATMID_warn(VDCount_warn-i-1,0)&"'};"
							next
%>							
			                return t;
			            },
							style: {color: '#FFFFFF', font: 'normal 14px Arial' },
			                align: 'left',
			                x: -120
			            },
			            title: { margin: 100, text: ' '},
			            plotLines: [{
			                value: 0,
			                width: 1,
			                color: '#808080'
			            }]
			        },
			        plotOptions: {
						series: {
							cursor: 'pointer',
							states: {
								hover: {
									enabled: false,
									halo: {
										size: 0
									}
								}
							},
							point: {
								events: {
									click: function (e) {
										selectDevice(2,this.y,Highcharts.dateFormat('%d.%m.%Y %H:%M', this.x));
										//lert(this.y);
									}
								}
							},
							marker: {
								enabled: true, 
			                    symbol: 'circle'
							}
						}					
			        },
			        series: [
					 <% =AllSeries_warn %> 
                    ]
				});		


				$('.highcharts-yaxis-labels text').bind('click',function () { 
					//console.info($(this).text()); 
					window.prompt("Копировать: Ctrl+C, Enter", $(this).text());
				});
					
			//$('.highcharts-xaxis-labels text').on('click', function () {   console.info($(this).text()); });


			});

		</script>	
<style  type="text/css">
	.hc-label {
		  fill: #00FFFF;
		  color: #00FFFF;
		}
</style>		
</head>
<body>
<div align="center">
<table border="0" width="100%" style="border: none;">
	<tr>
		<td style="border: none;">
		  <div id="containerAA" style="width: 200px; height: 300px; margin: 0 auto; font-size: 12pt;">
		    <table border="0" height="90%" width="100%" cellspacing="0" >
			  <tr><td style="border-top: solid 1px #C0C0C0; border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center">Всего:</td></tr>
			  <tr><td style="border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center; "><% =ValueAll %></td></tr>
			  <tr><td style="border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0;  text-align: center;"><% =ValueAll_discr %></td></tr>
              <tr><td style="border-top: solid 1px #C0C0C0; border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center" nowrap>Сбойные:</td></tr>
			  <tr><td style="border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center;"><% =ValueFail %></td></tr>
			  <tr><td style="border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center;"><% =ValueFail_discr %></td></tr>
              <tr><td style="border-top: solid 1px #C0C0C0; border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center" nowrap>На контроле:</td></tr>
			  <tr><td style="border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0;  text-align: center;"><% =ValueControl %></td></tr>
			  <tr><td style="border-bottom: solid 1px #C0C0C0; border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center;"><% =ValueControl_discr %></td></tr>
			</table>
		  </div></td>
		<td style="border: none;" colspan="2"><div id="container1"  style="width: 1100px; height: 300px; margin: 0 auto"></div></td>
	</tr>
	
	<tr>
		<td style="border: none;">
		  <div id="containerBB" style="width: 200px; height: 300px; margin: 0 auto; font-size: 12pt;">
		    <table border="0" height="90%" width="100%" cellspacing="0" >
              <tr><td style="border-top: solid 1px #C0C0C0; border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center" nowrap>Предупреждение:</td></tr>
			  <tr><td style="border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center;"><% =ValueWarining %></td></tr>
			  <tr><td style="border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center;"><% =ValueWarining_discr %></td></tr>
              <tr><td style="border-top: solid 1px #C0C0C0; border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center" nowrap>На контроле:</td></tr>
			  <tr><td style="border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0;  text-align: center;"><% =ValueControlWarning %></td></tr>
			  <tr><td style="border-bottom: solid 1px #C0C0C0; border-left: solid 1px #C0C0C0; border-right: solid 1px #C0C0C0; text-align: center;"><% =ValueControlWarning_discr %></td></tr>
			</table>
		  </div></td>
		<td style="border: none;" colspan="2"><div id="container2"  style="width: 1100px; height: 300px; margin: 0 auto"></div></td>
	</tr>
	<tr>
		<td style="width: 200px;">&nbsp;</td>
		<td style="border: none;" align="center" colspan=2 >
			<div id="containerCC" style="width: 850px; margin: 0 auto; font-size: 12pt;"></div>
		</td>
	</tr>	
	<tr>
		<td style="border: none;" colspan=3 >
			<div id="containerDD" style="width: 100%; text-align:left;  margin: 0 auto; font-size: 12pt;"><a href="<%=QOSLink %>">переход в QOS</a></div>
		</td>
	</tr>
</table>
</div>



<%
  Conn.Close
  set Cmd=Nothing
  set Conn = Nothing
  set Rs = Nothing
%>
</body>
</html>
<%
end if
%>