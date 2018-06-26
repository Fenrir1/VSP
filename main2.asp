<%@ language="VBScript"%><!-- #include file="const.asp" --><!-- #include file="common.asp" -->
<%
' Модификация 2017q1:
'	- добавлены ссылки на detail.asp
'	- замена в запросах 1.0/6 на 1.0/12
'	- убрал столбец и диаграмму КЦ
'---------------------------------------------------------------
' Вторая экранная форма, вывод 4,5,6 контролируемых параметров.
'---------------------------------------------------------------

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
  Response.Write("<html><body><div style='text-align: center;'><span style='font-size: 14pt; font-weight: 600; color: #800000}'>Для пользователя "&Auth_Name&" доступ не определен.</span></div></body></html>")
else ' юзер зареген, продолжаем:

' далее (до начала html кода) в рабочие переменные считываем данные из БД для отображения элементов на странице

set Rs=Server.CreateObject("ADODB.Recordset")

Rs.Open "SELECT * FROM Tags WHERE (TagID='Main5SMS')", Conn
if not Rs.Eof then
  Value5=Rs.Fields("Value")
  Color5=clNormal
  Main5_SetHi=Rs.Fields("SetHi")
  Main5_SetHiHi=Rs.Fields("SetHiHi")
  FontColor=""
  if Value5 >= Main5_SetHi then Color5=clWarning end if
  if Value5 >= Main5_SetHiHi then Color5=clError end if
end if
Rs.Close
Rs.Open "SELECT ISNULL(Max(Value), 0) FROM Tags_History WHERE (TagID='Main5SMS') AND (DT > GETDATE()-1.0/12)", Conn
ValueSMS_max=Rs.Fields(0)
Rs.Close
if ValueSMS_max<Main5_SetHiHi then ValueSMS_max=Main5_SetHiHi end if

Rs.Open "SELECT COUNT(*) FROM MV_SMSService", Conn
SMSCountServ=Rs.Fields(0)
Rs.Close
if SMSCountServ=0 then
Table="<tr><td class='A' style='text-align: left; color: red;' nowrap>Нет данных</td></tr>"&vbCrLf
end if 
if SMSCountServ>0 and SMSCountServ<4 then
Rs.Open "SELECT * FROM MV_SMSService ORDER BY [SERVER]", Conn
if not Rs.Eof then
r0="<tr><td class='A' rowspan=2 style='text-align: left; font-size: 22pt; padding-left: 4pt;' nowrap>"&DateTimeFormat(Rs.Fields("DT_FILE"), "dd.mm.yy")&"<br />"&DateTimeFormat(Rs.Fields("DT_FILE"), "hh:nn")&"</td>"
r00="<tr>"
r1="<tr><td class='A' style='text-align: left; padding-left: 4pt;' nowrap>в очереди</td>"
r2="<tr><td class='A' style='text-align: left; padding-left: 4pt;' nowrap>откл. ПЦ</td>"
r3="<tr><td class='A' style='text-align: left; padding-left: 4pt;' nowrap>откл. ОСС</td>"
r4="<tr><td class='A' style='text-align: left; padding-left: 4pt;' nowrap>всего</td>"
  bg="background-color: #363636; color: #99CCFF;"
  r0=r0&"<td class='A' colspan=2 style='text-align: center; font-size: 24pt; font-weight: 700;' nowrap>"&left(Rs.Fields("SERVER"), 4)&"</td>"
  r00=r00&"<td class='A' width='100px' style='text-align: center; font-size: 20pt;'>%</td>"
  r00=r00&"<td class='A' width='100px' style='text-align: center; font-size: 20pt;"&bg&"'>шт.</td>"
  FontColor=""
  v=Rs.Fields("WAIT_COUNT")/Rs.Fields("ALL_COUNT")*100
  if v>Main5_SetHi then FontColor="color: "&clWarning&"; " end if
  if v>Main5_SetHiHi then FontColor="color: "&clError&"; " end if
  r1=r1&"<td class='A' style='"&FontColor&"text-align: right; font-size: 36pt; font-weight: 700;'>"&FormatNumber(v, 0, -1, 0, 0)&"</td>"
  r1=r1&"<td class='A' style='text-align: right;"&bg&"'>"&Rs.Fields("WAIT_COUNT")&"</td>"
  
  FontColor=""
  v=Rs.Fields("REJECTED_COUNT")/Rs.Fields("ALL_COUNT")*100
  if v>Main5_SetHi then FontColor="color: "&clWarning&"; " end if
  if v>Main5_SetHiHi then FontColor="color: "&clError&"; " end if
  r2=r2&"<td class='A' style='"&FontColor&"text-align: right; font-size: 36pt; font-weight: 700;' nowrap>"&FormatNumber(v, 0, -1, 0, 0)&"</td>"
  r2=r2&"<td class='A' style='text-align: right;"&bg&"'>"&Rs.Fields("REJECTED_COUNT")&"</td>"

  FontColor=""
  v=Rs.Fields("DECLINED_COUNT")/Rs.Fields("ALL_COUNT")*100
  if v>Main5_SetHi then FontColor="color: "&clWarning&"; " end if
  if v>Main5_SetHiHi then FontColor="color: "&clError&"; " end if
  r3=r3&"<td class='A' style='"&FontColor&"text-align: right; font-size: 36pt; font-weight: 700;' nowrap>"&FormatNumber(v, 0, -1, 0, 0)&"</td>"
  r3=r3&"<td class='A' style='text-align: right;"&bg&"'>"&Rs.Fields("DECLINED_COUNT")&"</td>"

  r4=r4&"<td class='A' colspan=2 style='text-align: right;"&bg&" font-size: 28pt;' nowrap>"&Rs.Fields("ALL_COUNT")&"</td>"
  Rs.MoveNext
  if not Rs.Eof then
    r0=r0&"<td class='A' colspan=2 style='text-align: center; font-size: 24pt; font-weight: 700;' nowrap>"&left(Rs.Fields("SERVER"), 4)&"</td>"
    r00=r00&"<td class='A' width='100px' style='text-align: center; font-size: 20pt;'>%</td>"
    r00=r00&"<td class='A' width='100px' style='text-align: center; font-size: 20pt;"&bg&"'>шт.</td>"
    FontColor=""
    v=Rs.Fields("WAIT_COUNT")/Rs.Fields("ALL_COUNT")*100
    if v>Main5_SetHi then FontColor="color: "&clWarning&"; " end if
    if v>Main5_SetHiHi then FontColor="color: "&clError&"; " end if
    r1=r1&"<td class='A' style='"&FontColor&"text-align: right; font-size: 36pt; font-weight: 700;'>"&FormatNumber(v, 0, -1, 0, 0)&"</td>"
    r1=r1&"<td class='A' style='text-align: right;"&bg&"'>"&Rs.Fields("WAIT_COUNT")&"</td>"
	
    FontColor=""
    v=Rs.Fields("REJECTED_COUNT")/Rs.Fields("ALL_COUNT")*100
    if v>Main5_SetHi then FontColor="color: "&clWarning&"; " end if
    if v>Main5_SetHiHi then FontColor="color: "&clError&"; " end if
    r2=r2&"<td class='A' style='"&FontColor&"text-align: right; font-size: 36pt; font-weight: 700;' nowrap>"&FormatNumber(v, 0, -1, 0, 0)&"</td>"
    r2=r2&"<td class='A' style='text-align: right;"&bg&"'>"&Rs.Fields("REJECTED_COUNT")&"</td>"

    FontColor=""
     v=Rs.Fields("DECLINED_COUNT")/Rs.Fields("ALL_COUNT")*100
     if v>Main5_SetHi then FontColor="color: "&clWarning&"; " end if
     if v>Main5_SetHiHi then FontColor="color: "&clError&"; " end if
     r3=r3&"<td class='A' style='"&FontColor&"text-align: right; font-size: 36pt; font-weight: 700;' nowrap>"&FormatNumber(v, 0, -1, 0, 0)&"</td>"
     r3=r3&"<td class='A' style='text-align: right;"&bg&"'>"&Rs.Fields("DECLINED_COUNT")&"</td>"

     r4=r4&"<td class='A' colspan=2 style='text-align: right;"&bg&" font-size: 28pt;' nowrap>"&Rs.Fields("ALL_COUNT")&"</td>"
	 Rs.MoveNext
   end if
     'Rs.MoveNext
   if not Rs.Eof then
     r0=r0&"<td class='A' colspan=2 style='text-align: center; font-size: 24pt; font-weight: 700;' nowrap>"&left(Rs.Fields("SERVER"), 4)&"</td>"
     r00=r00&"<td class='A' width='100px' style='text-align: center; font-size: 20pt;'>%</td>"
     r00=r00&"<td class='A' width='100px' style='text-align: center; font-size: 20pt;"&bg&"'>шт.</td>"
     FontColor=""
     v=Rs.Fields("WAIT_COUNT")/Rs.Fields("ALL_COUNT")*100
     if v>Main5_SetHi then FontColor="color: "&clWarning&"; " end if
     if v>Main5_SetHiHi then FontColor="color: "&clError&"; " end if
     r1=r1&"<td class='A' style='"&FontColor&"text-align: right; font-size: 36pt; font-weight: 700;'>"&FormatNumber(v, 0, -1, 0, 0)&"</td>"
     r1=r1&"<td class='A' style='text-align: right;"&bg&"'>"&Rs.Fields("WAIT_COUNT")&"</td>" 
	
     FontColor=""
     v=Rs.Fields("REJECTED_COUNT")/Rs.Fields("ALL_COUNT")*100
     if v>Main5_SetHi then FontColor="color: "&clWarning&"; " end if
     if v>Main5_SetHiHi then FontColor="color: "&clError&"; " end if
     r2=r2&"<td class='A' style='"&FontColor&"text-align: right; font-size: 36pt; font-weight: 700;' nowrap>"&FormatNumber(v, 0, -1, 0, 0)&"</td>"
     r2=r2&"<td class='A' style='text-align: right;"&bg&"'>"&Rs.Fields("REJECTED_COUNT")&"</td>"

     FontColor=""
     v=Rs.Fields("DECLINED_COUNT")/Rs.Fields("ALL_COUNT")*100
     if v>Main5_SetHi then FontColor="color: "&clWarning&"; " end if
     if v>Main5_SetHiHi then FontColor="color: "&clError&"; " end if
     r3=r3&"<td class='A' style='"&FontColor&"text-align: right; font-size: 36pt; font-weight: 700;' nowrap>"&FormatNumber(v, 0, -1, 0, 0)&"</td>"
     r3=r3&"<td class='A' style='text-align: right;"&bg&"'>"&Rs.Fields("DECLINED_COUNT")&"</td>"

     r4=r4&"<td class='A' colspan=2 style='text-align: right;"&bg&" font-size: 28pt;' nowrap>"&Rs.Fields("ALL_COUNT")&"</td>"
   end if
Table=r0&"</tr>"&vbCrLf&r00&"</tr>"&vbCrLf&r1&"</tr>"&vbCrLf&r2&"</tr>"&vbCrLf&r3&"</tr>"&vbCrLf&r4&"</tr>"&vbCrLf
else
Table="<tr><td class='A' style='text-align: left; color: red;' nowrap>Нет данных</td></tr>"&vbCrLf
end if
Rs.Close
End if
if SMSCountServ>3 then
Rs.Open "SELECT (select top 1 DT_FILE FROM MV_SMSService ORDER BY [SERVER]) as DT_FILE, 'Общие сведения' as SERVER, SUM(ALL_COUNT) as ALL_COUNT, SUM(WAIT_COUNT) as WAIT_COUNT, SUM(REJECTED_COUNT) as REJECTED_COUNT, SUM(DECLINED_COUNT) as DECLINED_COUNT FROM MV_SMSService", Conn
 if not Rs.Eof then
 r0="<tr><td class='A' rowspan=2 style='text-align: left; font-size: 22pt; padding-left: 4pt;' nowrap>"&DateTimeFormat(Rs.Fields("DT_FILE"), "dd.mm.yy")&"<br />"&DateTimeFormat(Rs.Fields("DT_FILE"), "hh:nn")&"</td>"
 r00="<tr>"
 r1="<tr><td class='A' style='text-align: left; padding-left: 4pt;' nowrap>в очереди</td>"
 r2="<tr><td class='A' style='text-align: left; padding-left: 4pt;' nowrap>откл. ПЦ</td>"
 r3="<tr><td class='A' style='text-align: left; padding-left: 4pt;' nowrap>откл. ОСС</td>"
 r4="<tr><td class='A' style='text-align: left; padding-left: 4pt;' nowrap>всего</td>"
   bg="background-color: #363636; color: #99CCFF;"
   r0=r0&"<td class='A' colspan=2 style='text-align: center; font-size: 24pt; font-weight: 700;' nowrap>"&left(Rs.Fields("SERVER"), 4)&"</td>"
   r00=r00&"<td class='A' width='100px' style='text-align: center; font-size: 20pt;'>%</td>"
   r00=r00&"<td class='A' width='100px' style='text-align: center; font-size: 20pt;"&bg&"'>шт.</td>"
   FontColor=""
   v=Rs.Fields("WAIT_COUNT")/Rs.Fields("ALL_COUNT")*100
   if v>Main5_SetHi then FontColor="color: "&clWarning&"; " end if
   if v>Main5_SetHiHi then FontColor="color: "&clError&"; " end if
   r1=r1&"<td class='A' style='"&FontColor&"text-align: right; font-size: 36pt; font-weight: 700;'>"&FormatNumber(v, 0, -1, 0, 0)&"</td>"
   r1=r1&"<td class='A' style='text-align: right;"&bg&"'>"&Rs.Fields("WAIT_COUNT")&"</td>"
   
   FontColor=""
   v=Rs.Fields("REJECTED_COUNT")/Rs.Fields("ALL_COUNT")*100
   if v>Main5_SetHi then FontColor="color: "&clWarning&"; " end if
   if v>Main5_SetHiHi then FontColor="color: "&clError&"; " end if
   r2=r2&"<td class='A' style='"&FontColor&"text-align: right; font-size: 36pt; font-weight: 700;' nowrap>"&FormatNumber(v, 0, -1, 0, 0)&"</td>"
   r2=r2&"<td class='A' style='text-align: right;"&bg&"'>"&Rs.Fields("REJECTED_COUNT")&"</td>" 

   FontColor=""
   v=Rs.Fields("DECLINED_COUNT")/Rs.Fields("ALL_COUNT")*100
   if v>Main5_SetHi then FontColor="color: "&clWarning&"; " end if
   if v>Main5_SetHiHi then FontColor="color: "&clError&"; " end if
   r3=r3&"<td class='A' style='"&FontColor&"text-align: right; font-size: 36pt; font-weight: 700;' nowrap>"&FormatNumber(v, 0, -1, 0, 0)&"</td>"
   r3=r3&"<td class='A' style='text-align: right;"&bg&"'>"&Rs.Fields("DECLINED_COUNT")&"</td>" 

   r4=r4&"<td class='A' colspan=2 style='text-align: right;"&bg&" font-size: 28pt;' nowrap>"&Rs.Fields("ALL_COUNT")&"</td>"
 Table=r0&"</tr>"&vbCrLf&r00&"</tr>"&vbCrLf&r1&"</tr>"&vbCrLf&r2&"</tr>"&vbCrLf&r3&"</tr>"&vbCrLf&r4&"</tr>"&vbCrLf
 else
 Table="<tr><td class='A' style='text-align: left; color: red;' nowrap>Нет данных</td></tr>"&vbCrLf
 end if
 Rs.Close

End if

CurrentTime = DateTimeFormat(DateAdd("m", -1, Now), "yyyy, mm, dd, hh, nn")

Rs.Open "SELECT * FROM SkillGroups where Name<>'КЦ' ORDER BY [Name]", Conn
Color7=clNormal
Value6_max=0
LastTime6=Now
if not Rs.Eof then
  r0="<tr><td class='A' style='text-align: left; font-size: 22pt; padding-left: 4pt;'>&nbsp;</td>"
  r1="<tr><td class='A' style='text-align: left; font-size: 24pt; padding-left: 4pt;' nowrap>операторов</td>"
  r2="<tr><td class='A' style='text-align: left; font-size: 24pt; padding-left: 4pt;' nowrap>занято</td>"
  r3="<tr><td class='A' style='text-align: left; font-size: 24pt; padding-left: 4pt;' nowrap>клиентов в очереди</td>"
  r4="<tr><td class='A' style='text-align: left; font-size: 24pt; padding-left: 4pt;' nowrap>вызовов за 10мин</td>"
  mc=0
  
  do while not Rs.Eof
    cc=""
    if Rs.Fields("Name")="ДПП" then cc="color: #66FFFF"
    if Rs.Fields("Name")="КЦ" then cc="color: #FF66FF"
    if Rs.Fields("Name")="ТСП" then cc="color: #FFFF66"
    r0=r0&"<td class='A' style='font-size: 24pt; font-weight: 700; "&cc&"'>"&Rs.Fields("Name")&"</td>"
	
	r1=r1&"<td class='A' style='font-size: 36pt; font-weight: 700;'>"&Rs.Fields("ReadyOperator")&"</td>"
	r2=r2&"<td class='A' style='font-size: 36pt; font-weight: 700;'>"&Rs.Fields("TalkingOperator")&"</td>"

	v=Rs.Fields("QueuedUsers")
	if v>Value6_max then 
	  Value6_max=v 
	end if
	bg=""
	Ro=Rs.Fields("ReadyOperator")
	Qu=Rs.Fields("QueuedUsers")
	if Ro>0 then
	  D=Qu/Ro
	  if (Ro=1) and (Qu=1) then
	  else
	    if D>=0.8 then 
		  bg="background-color: "&clWarning&"; color: #000; "
		  mc=mc or 1 
		end if
	    if D>=1.0 then
		  bg="background-color: "&clError&"; color: #ffffff; "
		  mc=mc or 2 
		end if
	  end if
	else 'обработка исключения деления на 0
	  if Qu>0 then 
	    bg="background-color: "&clError&"; color: #ffffff; "
	    mc=mc or 2
	  end if
	end if

	if isnull(v) then v=0 end if 'TEMP
    r3=r3&"<td class='A' style='"&bg&"font-size: 36pt; font-weight: 700;'>"&FormatNumber(v, 0, -1, 0, 0)&"</td>"

	if isnull(Rs.Fields("Incoming10Min")) then  'TEMP
		r4=r4&"<td class='A' style='font-size: 36pt; font-weight: 700;'>0</td>"
	else
		r4=r4&"<td class='A' style='font-size: 36pt; font-weight: 700;'>"&FormatNumber(Rs.Fields("Incoming10Min"), 0, -1, 0, 0)&"</td>"
	end if
	if Rs.Fields("LastTime")<LastTime6 then
	  LastTime6=Rs.Fields("LastTime")
	end if
    Rs.MoveNext
  loop
  Table3=r0&"</tr>"&vbCrLf&r1&"</tr>"&vbCrLf&r2&"</tr>"&vbCrLf&r3&"</tr>"&vbCrLf&r4&"</tr>"&vbCrLf
  if mc and 1 = 1 then Color7=clWarning
  if mc and 2 = 2 then Color7=clError
else
  Table3="<tr><td class='A' style='text-align: left; color: red;' nowrap>Нет данных</td></tr>"&vbCrLf
end if
Rs.Close


dim series(), CID()
Table2=""
CID_list=""
'Rs.Open "SELECT * FROM vw_Messages ORDER BY 1", Conn
t2FontSize = 24
t2NamesCount = 6
Rs.Open "SELECT count([Name]) namescount FROM Messages_Category ", Conn
if not Rs.Eof then
  if (Rs.Fields("namescount")<7) then
		ReDim series(6)
		ReDim CID(6)
	elseif (Rs.Fields("namescount")=7) then
		t2FontSize = 20
		t2NamesCount = 7
		ReDim series(7)
		ReDim CID(7)
	elseif (Rs.Fields("namescount")>=8 ) then
		t2FontSize = 16
		t2NamesCount = 8
		ReDim series(8)
		ReDim CID(8)
	end if
end if
Rs.Close

i = 1
Rs.Open "SELECT  top 8 [Name] FROM Messages_Category order by [Name]", Conn
if not Rs.Eof then
  do while not Rs.Eof
		if ((t2NamesCount-i)>=0) then
			CID(t2NamesCount-i) = Rs.Fields("Name")
			series(t2NamesCount-i)=""
		end if
		if CID_list<>"" then
			CID_list=CID_list&",'"&Rs.Fields("Name")&"'"
		else 
			CID_list=CID_list&"'"&Rs.Fields("Name")&"'"
		end if 
		i=i+1
		Rs.MoveNext
  loop
end if
Rs.Close


SQL_="SELECT top 8 C.[Name], T.LastState FROM Messages_Category AS C LEFT OUTER JOIN (SELECT CategoryCode, MAX(LastState) AS LastState "
SQL_=SQL_&" FROM  DBO.Messages_Type GROUP BY CategoryCode) AS T ON C.CategoryCode = T.CategoryCode order by C.[Name]"
Rs.Open SQL_, Conn
MaxState=0
if not Rs.Eof then
  do while not Rs.Eof
    cl="color: "&clNormal&"; "
	if Rs.Fields(1)=1 then cl="color: "&clWarning&"; " end if
	if Rs.Fields(1)=2 then cl="color: "&clError&"; " end if
	if MaxState<Rs.Fields(1) then MaxState=Rs.Fields(1) end if
    Table2=Table2&"<tr><td class='A' style='"&cl&"font-size: "&t2FontSize&"pt; font-weight: 700;' nowrap>"&Rs.Fields(0)&"</td></tr>"
    Rs.MoveNext
  loop
end if
Color6=clNormal
if MaxState=1 then Color6=clWarning end if
if MaxState=2 then Color6=clError end if
Rs.Close

'SQL_="SELECT [DT], [TagID], [Value], CASE [TagID] WHEN 'AGENT' THEN 5.5 WHEN 'CLOSE' THEN 4.5 WHEN 'IR ACCEPT' THEN 3.5 WHEN 'IBSO' THEN 2.5 WHEN 'OTHER' THEN 1.5 WHEN 'WATCH' THEN 0.5 ELSE 0 END AS Y FROM Tags_History WHERE (DT > GETDATE()-1.0/6) AND (TagID not like 'Main%') AND (TagID not like '5%') ORDER BY TagID, DT"
SQL_="SELECT dateAdd(ss,-1*DATEPART(ss, DT),dateAdd(ms,-1*DATEPART(ms, DT),dateAdd(month,-1,DT))) AS DT , "
SQL_=SQL_&" [Value], [TagID] FROM Tags_History WHERE (DT > GETDATE()-1.0/6) AND (TagID not like 'Main%') "
SQL_=SQL_&" AND (TagID in ("&CID_list&") ) "
SQL_=SQL_&" AND (TagID not like '5%') ORDER BY TagID, DT"
Rs.Open SQL_, Conn
LastID=""

'dim series(6), CID(6)
'ReDim series(t2NamesCount)
'ReDim CID(t2NamesCount)

i=-1
if not Rs.Eof then
do while not Rs.Eof
  if LastID <> Rs.Fields("TagID") then 
    i=i+1 
  end if
  if i<6 then
    select case Rs.Fields("Value")
      case "0": m="marker: {fillColor: '#00FF00', lineColor: '#00FF00', radius: 3}, "
      case "1": m="marker: {fillColor: '#FF0000', lineColor: '#FFFF99', radius: 6}, "
      case "2": m="marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 6}, "
    end select

		y=-1
		for j=0 to t2NamesCount
		  if (CID(j)=Rs.Fields("TagID")) then y=j end if
		next

		if (y>0) then
			y=y+0.5
		end if

	'v=Rs.Fields("Y")
	v=cStr(y)
	v=replace(v, ",", ".")
	series(i)=series(i)&vbCrLf&"{color: null, "&m&"x: Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"},"
  end if
  LastID=Rs.Fields("TagID")
  Rs.MoveNext
loop
  AllSeries=""
  for i=0 to 5
    if series(i)<>"" then
      series(i)=left(series(i), len(series(i))-1)
      series(i)="{ name: '"&CID(i)&"', type: 'scatter', data: ["&series(i)&"]},"
    end if
    AllSeries=AllSeries&series(i)
  next
  if AllSeries<>"" then
    AllSeries=left(AllSeries, len(AllSeries)-1)
    AllSeries=", series: ["&AllSeries&"]"
  end if
else
  AllSeries=""
end if
Rs.Close

if (AllSeries="") then
	AllSeries = AllSeries & ", series: [{ name: '00', type: 'scatter', visible: false, data: [ "
	AllSeries = AllSeries & " {color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 2}, "
 	AllSeries = AllSeries & " x: Date.UTC("&DateTimeFormat(Now, "yyyy, mm, dd, hh, nn")&"), y: -1 } ]}] "
end if


%>
<!DOCTYPE HTML>
<html>
<head>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
		<meta http-equiv="X-UA-Compatible" content="ie=edge">
		<!-- <meta content="60; url=http://ufa-qos01ow/vsp/main2.asp" http-equiv=refresh> -->
		<!-- 1. Add these JavaScript inclusions in the head of your page -->
		<script type="text/javascript" src="js/jquery.min.js"></script>
		<script type="text/javascript" src="js/highstock.js"></script>
		<!-- <script type="text/javascript" src="js/highcharts.js"></script>-->
		<script type="text/javascript" src="js/themes/gray.js"></script>
		<!-- 2. Add the JavaScript to initialize the chart on document ready -->
		<script type="text/javascript">
		
			var chart5;
			var chartE;
			var chart6;
			var chartF;
			var chart7;
			var chartG;

			// Первый график
			$(document).ready(function() {
				chart5 = new Highcharts.Chart({
					chart: {
						renderTo: 'container5',
						type: 'line'
					},
					colors: ["<% =Color5 %>"],
					credits: {enabled: false},
					legend:  {enabled: false},
					tooltip: {enabled: false},
					title:   {align: 'right', text: '<span style="text-decoration: underline;" >СМС-сервис</span>'},
					// subtitle: {align: 'left', text: 'Примечание:'},
					xAxis: [{
						max: Date.UTC(<% =CurrentTime %>),
						type: 'datetime',
						tickmarkPlacement: 'on',
						dateTimeLabelFormats: { // don't display the dummy year
							hour: '%H:%M'
						}
					}],
					yAxis: [
					{
						min: 0,
						max: <% =Replace(ValueSMS_max, ",", ".") %>,
						title: {
							text: null
						},
						allowDecimals: false,
						labels: {
							step: 2,
							formatter: function() {return this.value +'%'; }
						},
						plotLines: [{
							value: 0,
							width: 1,
							color: '#808080'
						}]
					}
					],
					plotOptions: {
						line: {
							dataLabels: {
								enabled: true
							},
							enableMouseTracking: false
						}
					},
					series: [
					{   data: [
<%
LastTime5=Now
Rs.Open "SELECT top 1 DT FROM Tags_History WHERE (TagID='Main5SMS') AND (DT > GETDATE()-1.0/12) ORDER BY DT desc", Conn
if not Rs.Eof then
	LastTime5=Rs.Fields("DT")
end if
Rs.Close

Rs.Open "SELECT dateadd(month,-1,DT) DT, TagID, [Value] FROM Tags_History WHERE (TagID='Main5SMS') AND (DT > GETDATE()-1.0/12) ORDER BY DT", Conn
'LastTime5=Now
if not Rs.Eof then
  Response.Write("[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Replace(FormatNumber(Rs.Fields("Value"), 0, -1, 0, 0), ",", ".")&"]")
  Rs.MoveNext
  do while not Rs.Eof
    Response.Write(","&vbCrLf&"[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Replace(FormatNumber(Rs.Fields("Value"), 0, -1, 0, 0), ",", ".")&"]")
	'LastTime5=Rs.Fields("DT")
    Rs.MoveNext
  loop
end if
Rs.Close
%>
						]
					}
					]
				});


  			$('#container5 .highcharts-title').click(function(){   
					window.location.replace("detail.asp?T=5");
				});

				chart5.yAxis[0].addPlotLine({
					value: <% =Main5_SetHiHi%>, color: '#FF9900', dashStyle: 'Dash', width: 2, id: 'plot-line-1'
				});

				chartE = new Highcharts.Chart({
					chart:   {renderTo: 'containerE', type: 'line', margin: [0, 0, 0, 0] },
					credits: {enabled: false},	legend:  {enabled: false},	tooltip: {enabled: false},	title:   {text: ''}
				});
				chartE.renderer.circle(150, 150, 90).attr({
					fill: '<% =Color5 %>',
					stroke: '<% =Color5 %>'
				}).add();
				<%
				if DateDiff("n", LastTime5, Now())>20 then
					Response.Write("chartE.renderer.image('q.gif', 75, 75, 150, 150).add();")
				end if
				%>
				

				chart6 = new Highcharts.Chart({
					chart: {
						renderTo: 'container6'
					},
					credits: {enabled: false},
					legend:  {enabled: false},
					tooltip: {enabled: false},
					title:   {align: 'right', text: '<span style="text-decoration: underline;"  >Автоматизация</span>'},
					xAxis: {
						max: Date.UTC(<% =CurrentTime %>),
						type: 'datetime',
						dateTimeLabelFormats: { // don't display the dummy year
							hour: '%H:%M'
						}
					},
					yAxis: {
					    min: 0,
					    max: <% =t2NamesCount %>,
						tickInterval: 1,
						labels: { formatter: function() 
								  {
								    var t;
<%
										for j=0 to t2NamesCount
											Response.Write "if (this.value == "&j&") {t='"&CID(j)&"'};"
										next
%>
									return t;
								  },
								style: {color: '#FFFFFF'},
								align: 'left',
								x: -150,
								y: -6
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
						}
					}
					<% =AllSeries %> 
					
				});

				$('#container6 .highcharts-title').click(function(){   
					window.location.replace("detail.asp?T=6");
				});
				



				chartF = new Highcharts.Chart({
					chart:   {renderTo: 'containerF', type: 'line', margin: [0, 0, 0, 0] },
					credits: {enabled: false},	legend:  {enabled: false},	tooltip: {enabled: false},	title:   {text: ''}
				});
				chartF.renderer.circle(150, 150, 90).attr({
					fill: '<% =Color6 %>',
					stroke: '<% =Color6 %>'
				}).add();
				
				chart7 = new Highcharts.Chart({
					chart: {
						renderTo: 'container7',
						// defaultSeriesType: 'column'
						type: 'line'
					},
					colors: ['#66FFFF', '#FFFF66'], //'#FF66FF'
					credits: {enabled: false},
					legend:  {enabled: false},
					tooltip: {enabled: false},
					title:   {align: 'right', text: 'Клиентов в очереди'},
					xAxis: [{
						max: Date.UTC(<% =CurrentTime %>),
						type: 'datetime',
						dateTimeLabelFormats: { // don't display the dummy year
							hour: '%H:%M'
						}
					}],
					yAxis: [
					{
					    min: 0,
						title: {
							text: null
						},
						lineColor: '#66FFFF',
						allowDecimals: false,
						plotLines: [{
							value: 0,
							width: 1,
							color: '#808080'
						}]
					}
					],
					plotOptions: {
						line: {
							dataLabels: {
								enabled: true,
								formatter: function() {
									return this.y > 0  ? this.y : null; 
								}
							},
							enableMouseTracking: false
						}
					},
					legend: {
						enabled : false,
						layout: 'horizontal',
						floating: true,
						backgroundColor: '#363636',
						align: 'left',
						verticalAlign: 'top',
						x: 4,
						y: -8,
						borderWidth: 0
					},
					series: [{
						name: 'ДПП'
<%
Rs.Open "SELECT DT, TagID, [Value] FROM Tags_History WHERE (TagID='5198') AND (DT > GETDATE()-1.0/12) ORDER BY DT", Conn
																																			
if not Rs.Eof then
  temp_DT=datepart("yyyy",Rs.Fields("DT"))&", "&(datepart("m",Rs.Fields("DT"))-1)&", "&datepart("d",Rs.Fields("DT"))&", "&datepart("h",Rs.Fields("DT"))&", "&datepart("n",Rs.Fields("DT"))
  Response.Write(", data: [")
  Response.Write("[Date.UTC("&temp_DT&"), "&Rs.Fields("Value")&"]")
  Rs.MoveNext
  do while not Rs.Eof
	temp_DT=datepart("yyyy",Rs.Fields("DT"))&", "&(datepart("m",Rs.Fields("DT"))-1)&", "&datepart("d",Rs.Fields("DT"))&", "&datepart("h",Rs.Fields("DT"))&", "&datepart("n",Rs.Fields("DT"))
    Response.Write(","&vbCrLf&"[Date.UTC("&temp_DT&"), "&Rs.Fields("Value")&"]")
    Rs.MoveNext
  loop
  Response.Write("]"&vbCrLf)
else
	twoHoursBefore = DateAdd("h",-2,Now)
	temp_DT=datepart("yyyy",twoHoursBefore)&", "&(datepart("m",twoHoursBefore)-1)&", "&datepart("d",twoHoursBefore)&", "&datepart("h",twoHoursBefore)&", "&datepart("n",twoHoursBefore)
	DPPSeries = ", type: 'scatter', data: [ "
	DPPSeries = DPPSeries & " {color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 2}, "
 	DPPSeries = DPPSeries & " x: Date.UTC("&temp_DT&"), y: -1 } ] "
	Response.Write DPPSeries
end if
Rs.Close
%>
					}, {
						name: 'ТСП'
<%
Rs.Open "SELECT DT, TagID, [Value] FROM Tags_History WHERE (TagID='5349') AND (DT > GETDATE()-1.0/12) ORDER BY DT", Conn
if not Rs.Eof then
  temp_DT=datepart("yyyy",Rs.Fields("DT"))&", "&(datepart("m",Rs.Fields("DT"))-1)&", "&datepart("d",Rs.Fields("DT"))&", "&datepart("h",Rs.Fields("DT"))&", "&datepart("n",Rs.Fields("DT"))
  Response.Write(", data: [")
  Response.Write("[Date.UTC("&temp_DT&"), "&Rs.Fields("Value")&"]")
  Rs.MoveNext
  do while not Rs.Eof
	temp_DT=datepart("yyyy",Rs.Fields("DT"))&", "&(datepart("m",Rs.Fields("DT"))-1)&", "&datepart("d",Rs.Fields("DT"))&", "&datepart("h",Rs.Fields("DT"))&", "&datepart("n",Rs.Fields("DT"))
    Response.Write(","&vbCrLf&"[Date.UTC("&temp_DT&"), "&Rs.Fields("Value")&"]")
    Rs.MoveNext
  loop
  Response.Write("]"&vbCrLf)
else
 ' Response.Write(", data: [{x: Date.UTC("&DateTimeFormat(Int(Now), "yyyy, mm, dd")&"), y: 0}]")
	twoHoursBefore = DateAdd("h",-2,Now)
	temp_DT=datepart("yyyy",twoHoursBefore)&", "&(datepart("m",twoHoursBefore)-1)&", "&datepart("d",twoHoursBefore)&", "&datepart("h",twoHoursBefore)&", "&datepart("n",twoHoursBefore)
	TSPSeries = ", type: 'scatter', data: [ "
	TSPSeries = TSPSeries & " {color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 2}, "
 	TSPSeries = TSPSeries & " x: Date.UTC("&temp_DT&"), y: -1 } ] "
	Response.Write TSPSeries
end if
Rs.Close
%>
					}
					]
				});


				chartG = new Highcharts.Chart({
					chart:   {renderTo: 'containerG', type: 'line', margin: [0, 0, 0, 0] },
					credits: {enabled: false},	legend:  {enabled: false},	tooltip: {enabled: false},	title:   {text: ''}
				});
				chartG.renderer.circle(150, 150, 90).attr({
					fill: '<% =Color7 %>',
					stroke: '<% =Color7 %>'
				}).add();
				<%
				if (DateDiff("n", LastTime6, Now())>20) then
					Response.Write("chartG.renderer.image('q.gif', 75, 75, 150, 150).add();")
				end if
				%>
				
			});

		</script>
		
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
	text-align: center;
	color: #FFFFFF;
	font-family: Verdana, Arial, helvetica, sans-serif, Geneva;
}
TD.A {
	border: solid 1px #4572A7
}
TD.Head {
	color: #000000;
	font-size: 28pt;
}
TD.Txt {
	color: #FFFFFF;
	font-size: 48pt;
	font-weight: 700;
}
	
	-->
	</style>
</head>
<body>
<div align="center" valign="top">
<table border="0" width="1900" style="border: none;">
	<tr>
		<td style="border: none;" width="320px"><div id="containerE"  style="width: 320px; height: 320px; margin: 0 auto"></div></td>
		<td style="border: none;" width="640px">
		  <div id="containerAA" style="width: 640px; height: 320px; margin: 0 auto; font-size: 28pt;">
		    <table border="0" cellspacing="0" cellpadding="0" height="100%" width="640px">
			<% Response.Write(Table) %>			
 		    </table>
		  </div></td>
		<td style="border: none;" width="940px"><div id="container5"  style="width: 940px; height: 320px; margin: 0 auto"></div></td>
	</tr>
	<tr>
		<td style="border: none;" width="320px"><div id="containerF"  style="width: 320px; height: 320px; margin: 0 auto"></div></td>
		<td style="border: none;" width="640px">
		  <div id="containerAA" style="width: 640px; height: 320px; margin: 0 auto; font-size: 28pt;">
		    <table border="0" cellspacing="2" cellpadding="0" height="100%" width="640px">
			<% Response.Write(Table2) %>

 		    </table>
		  </div></td>
		<td style="border: none;" width="940px"><div id="container6"  style="width: 940px; height: 320px; margin: 0 auto"></div></td>
	</tr>
	<tr>
		<td style="border: none;" width="320px"><div id="containerG"  style="width: 320px; height: 320px; margin: 0 auto"></div></td>
		<td style="border: none;" width="640px">
		  <div id="containerAA" style="width: 640px; height: 320px; margin: 0 auto; font-size: 28pt;">
		    <table border="0" cellspacing="0" cellpadding="0" height="100%" width="640px">
			<% Response.Write(Table3) %>
 		    </table>
		  </div></td>
		<td style="border: none;" width="940px"><div id="container7"  style="width: 940px; height: 320px; margin: 0 auto"></div></td>
	</tr>
</table>
</div>
<%
  Conn.Close
  set Conn = Nothing
  set Rs = Nothing
%>
</body>
</html>
<%
end if
%>
<!-- Разработка: Машков А.В. -->
<!-- Для вывода графиков используется библиотека Highcharts JS - http://highsoft.com/ -->
