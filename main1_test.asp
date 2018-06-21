<%@ language="VBScript"%><!-- #include file="const.asp" --><!-- #include file="common.asp" -->
<%
' Первая экранная форма, вывод 1,2,3 контролируемых параметров.

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
Rs.Open "SELECT top 1 DT_FILE FROM NV_Operations order by DT_FILE desc", Conn 'тут дата последнего обработанного файла "неуспешные операции"
if isnull(Rs.Fields("DT_FILE")) then
 Value1Date=" "
else
 Value1Date=DateTimeFormat(Rs.Fields("DT_FILE"), "dd.mm.yy hh:nn")
end if 
Rs.Close
Rs.Open "SELECT * FROM Tags WHERE (TagID='Main1')", Conn 'тут Main1  это неуспешные операции
Value1=Rs.Fields("Value") ' последнее значение
Value1all=Rs.Fields("ValueDetail") ' Дополнительная информация, относящаяся к последнему значению.
if rs.Fields("ValueDetail")<>"0" then
Value1proc=100*Rs.Fields("Value")/Rs.Fields("ValueDetail") 'вычесление процента
end if
Color1=clNormal ' clNormal  = "#00FF00" цвет индикатора color1  1 индикатор
Main1_SetHiHi=Rs.Fields("SetHiHi") 'Уставка – превышение критичного значения
if Value1proc >= Rs.Fields("SetHi") then Color1=clWarning end if ' clWarning = "#FFFF00"  sethi поле для переключения предупреждения
if Value1proc >= Main1_SetHiHi then Color1=clError end if 'clError   = "#FF0000"
Rs.Close
Rs.Open "SELECT ISNULL(MAX(Value), 0) FROM Tags_History WHERE (TagID='Main1') AND (DT > GETDATE()-1.0/6)", Conn '
Value1max=Rs.Fields(0)
Rs.Close

Color2=0
dim Colors2(4,2)
dim Text2(4,2)
for i=1 to 4
  for j=1 to 2
    Colors2(i, j)=0
	Text2(i,j)=""
  next
next  

Rs.Open "SELECT ISNULL(Max(Value), 0) FROM Tags_History WHERE (TagID='Main2All24') AND (DT > GETDATE()-1.0/6)", Conn 'Недоступность всех АТМ
Value2A_max=Rs.Fields(0)
Rs.Close

Rs.Open "SELECT ISNULL(Max(Value), 0) FROM Tags_History WHERE (TagID='Main2All') AND (DT > GETDATE()-1.0/6)", Conn 'Недоступность всех АТМ
if Value2A_max<Rs.Fields(0) then
Value2A_max=Rs.Fields(0)
end if
Rs.Close

Rs.Open "SELECT ISNULL(Max(Value), 0) FROM Tags_History WHERE (TagID='Main2AllBPT24') AND (DT > GETDATE()-1.0/6)", Conn 'Недоступность всех БПТ
Value2AllBPT_max=Rs.Fields(0)
Rs.Close

Rs.Open "SELECT ISNULL(Max(Value), 0) FROM Tags_History WHERE (TagID='Main2AllBPT') AND (DT > GETDATE()-1.0/6)", Conn 'Недоступность всех БПТ
if Value2AllBPT_max<Rs.Fields(0) then
Value2AllBPT_max=Rs.Fields(0)
end if
Rs.Close

Main2A_SetHiHi=0
Main2AllBPT_SetHiHi=0
Main2F_SetHiHi=0
Rs.Open "SELECT TagID, [Value], SetHi, SetHiHi, [ValueDetail] FROM Tags WHERE (TagID like 'Main2%') ORDER BY TagID", Conn 'заполнение таблицы основных параметров АТМ БПТ
do while not Rs.Eof
  if Rs.Fields(0) = "Main2All" then ''Недоступность всех АТМ
    Text2(1,1)=Rs.Fields(1)
    if Rs.Fields(1)>Rs.Fields(2) then Colors2(1,1)=1 end if
    if Rs.Fields(1)>Rs.Fields(3) then Colors2(1,1)=2 end if
	Main2A_SetHiHi=Rs.Fields(3)
  end if
  if Rs.Fields(0) = "Main2AllBPT" then 'Недоступность всех БПТ
    Text2(1,2)=Rs.Fields(1)
    if Rs.Fields(1)>Rs.Fields(2) then Colors2(1,2)=1 end if
    if Rs.Fields(1)>Rs.Fields(3) then Colors2(1,2)=2 end if
	Main2AllBPT_SetHiHi=Rs.Fields(3)
  end if
  if Rs.Fields(0) = "Main2All24" then ''Недоступность всех АТМ
    Text2(4,1)=Rs.Fields(1)
    if Rs.Fields(1)>Rs.Fields(2) then Colors2(4,1)=1 end if
    if Rs.Fields(1)>Rs.Fields(3) then Colors2(4,1)=2 end if
	Main2A24_SetHiHi=Rs.Fields(3)
  end if
  if Rs.Fields(0) = "Main2AllBPT24" then 'Недоступность всех БПТ
    Text2(4,2)=Rs.Fields(1)
    if Rs.Fields(1)>Rs.Fields(2) then Colors2(4,2)=1 end if
    if Rs.Fields(1)>Rs.Fields(3) then Colors2(4,2)=2 end if
	Main2AllBPT24_SetHiHi=Rs.Fields(3)
  end if
  if Rs.Fields(0) = "Main2Fil" then 'Недоступность АТМ одного филиала'
    Text2(2,1)=Rs.Fields(1)
    if Rs.Fields(1)>Rs.Fields(2) then Colors2(2,1)=1 end if
    if Rs.Fields(1)>Rs.Fields(3) then Colors2(2,1)=2 end if
    if Rs.Fields(4)>0 then Colors2(2,1)=3 end if
	Main2F_SetHiHi=Rs.Fields(3)
  end if
  if Rs.Fields(0) = "Main2FilBPT" then 'Кол-во филиалов с недоступными БПТ
    Text2(2,2)=Rs.Fields(1)
    if Rs.Fields(1)>Rs.Fields(2) then Colors2(2,2)=1 end if
    if Rs.Fields(1)>Rs.Fields(3) then Colors2(2,2)=2 end if
    if Rs.Fields(4)>0 then Colors2(2,2)=2 end if
  end if
  if Rs.Fields(0) = "Main2Centr" then  'Недоступность АТМ по центральной схеме подключения
    Text2(3,1)=Rs.Fields(1)
    if Rs.Fields(1)=1 then Colors2(3,1)=1 end if
    if Rs.Fields(1)>1 then Colors2(3,1)=2 end if
  end if
  Rs.MoveNext
loop
if Value2A_max<Main2A_SetHiHi then Value2A_max=Main2A_SetHiHi end if
if Value2AllBPT_max<Main2AllBPT_SetHiHi then Value2AllBPT_max=Main2AllBPT_SetHiHi end if
castil=Colors2(1,1) 'castil для того что бы "процент неработоспособных банкоматов процессинга" не оказывал влияния на индикатор
castil1=Colors2(1,2) ' castil1 Процент неработоспособных БПТ процессинга сделать информационным 
Colors2(1,1)=0 ' castil для того что бы "процент неработоспособных банкоматов процессинга" не оказывал влияния на индикатор
Colors2(1,2)=0 ' castil1 Процент неработоспособных БПТ процессинга сделать информационным
for i=1 to 4
  for j=1 to 2
    if Color2 < Colors2(i, j) then Color2=Colors2(i, j) end if
	if i=1 and j=1 then Colors2(1,1)=castil end if'castil для того что бы "процент неработоспособных банкоматов процессинга" не оказывал влияния на индикатор
	if i=1 and j=2 then Colors2(1,2)=castil1 end if' castil1 Процент неработоспособных БПТ процессинга сделать информационным
	if Colors2(i, j)=0 then 
	  Colors2(i, j)=""
	else
	  if Colors2(i, j)=1 then 
	    Colors2(i, j)=clWarning
	  else
	    Colors2(i, j)=clError
	  end if
	end if
  next
next

	if Color2=0 then 
	  Color2=clNormal
	else
	  if Color2=1 then 
	    Color2=clWarning
	  else
	    Color2=clError
	  end if
	end if
Rs.Close

Rs.Open "SELECT [Value], [Prop_Crit], count([Value]) FROM [Tags] WHERE ([FileID]='CV') and (Prop_Active=1) GROUP BY [Value], [Prop_Crit] ORDER BY [Prop_Crit]", Conn 'CV - Данные по состоянию каналов
Value3total=0
Value3linkdown=0
Value3proc=0
Color3=clNormal
if not Rs.Eof then 
  do while not Rs.Eof
    if Rs.Fields(0) = 0 then 
	  Value3linkdown=Value3linkdown+Rs.Fields(2)
      if Rs.Fields(1)=1 then Color3=clWarning
	  if ((Rs.Fields(1)=1)and(Rs.Fields(2)>=4)) then Color3=clError
      if Rs.Fields(1)=2 then Color3=clError
	end if
    Value3total=Value3total+Rs.Fields(2)
    Rs.MoveNext
  loop
  Value3proc=round(Value3linkdown*100/Value3total)
end if
Value3=Value3linkdown
Rs.Close

SQL_="SELECT dateAdd(ss,-1*DATEPART(ss, A.DT),dateAdd(ms,-1*DATEPART(ms, A.DT),dateAdd(month,-1,A.DT))) AS DT, A.CHANNEL_ID, A.CHANNEL, A.[VALUE], B.LastValue "&_
"FROM vw_Channel_History AS A INNER JOIN "&_
"	(SELECT TagID, [Value] AS LastValue FROM Tags WHERE (FileID='CV') AND (Prop_Active=1)) AS B ON A.CHANNEL_ID=B.TagID "&_
"WHERE (A.DT > GETDATE()-1.0/6) AND "&_
"	(A.CHANNEL_ID in (SELECT DISTINCT CHANNEL_ID FROM vw_Channel_History WHERE [Value]=0 AND DT > GETDATE()-1.0/6)) "&_
"GROUP BY B.LastValue, A.CHANNEL_ID, A.CHANNEL, dateAdd(ss,-1*DATEPART(ss, A.DT),dateAdd(ms,-1*DATEPART(ms, A.DT),dateAdd(month,-1,A.DT))), A.[VALUE] "&_
"ORDER BY B.LastValue, A.CHANNEL_ID, A.CHANNEL, dateAdd(ss,-1*DATEPART(ss, A.DT),dateAdd(ms,-1*DATEPART(ms, A.DT),dateAdd(month,-1,A.DT))), A.[VALUE]"

Rs.Open SQL_, Conn
Value3Down=""
Value3Up=""
cnt_dn=0
cnt_up=0
LastID=""
dim series(5), CID(5)
for i=1 to 5 
  series(i)=""
  CID(i)=""
next
i=0
do while not Rs.Eof
  if LastID <> Rs.Fields("CHANNEL_ID") then
    if Rs.Fields("LastValue") = 0 then ' Канал лежит
	  if cnt_dn < 5 then Value3Down=Value3Down&Rs.Fields("CHANNEL")&", " end if
	  cnt_dn=cnt_dn+1
    else ' Канал лежал, восстановился
	  if cnt_up < 5 then Value3Up=Value3Up&Rs.Fields("CHANNEL")&", " end if
	  cnt_up=cnt_up+1
    end if
  end if
  if LastID <> Rs.Fields("CHANNEL_ID") then 
    i=i+1 
  end if
  if i<6 then
	CID(i)=Rs.Fields("CHANNEL")
    if Rs.Fields("Value")=0 then 
	  m="marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 6}, "
	else
	  m="marker: {fillColor: '#00FF00', lineColor: '#00FF00', radius: 3}, "
	end if
	v=5.5-i
	v=replace(v, ",", ".")
	if InStr(m, ".png")>0 then
	  series(i)=series(i)&vbCrLf&"{"&m&"x: Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"},"
	else
	  series(i)=series(i)&vbCrLf&"{color: null, "&m&"x: Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"},"
	end if
  end if
  LastID=Rs.Fields("CHANNEL_ID")
  Rs.MoveNext
loop
AllSeries=""
for i=1 to 5
  if series(i)<>"" then
    series(i)=left(series(i), len(series(i))-1)
    series(i)=", { name: '"&CID(i)&"', type: 'scatter', data: ["&series(i)&"]}"  
	CID(i)=left(CID(i), 12)
  end if
  AllSeries=AllSeries+series(i)
next
if cnt_dn>5 then Value3Down=Value3Down&"###, " end if
if InStr(Value3Down, ",")>0 then Value3Down=Left(Value3Down, len(Value3Down)-2) end if
if cnt_up>5 then Value3Up=Value3Up&"###, " end if
if InStr(Value3Up, ",")>0 then Value3Up=Left(Value3Up, len(Value3Up)-2) end if
Rs.Close

'response.write AllSeries
'response.end
'DateAdd("m", -1, Now)
CurrentTime = DateTimeFormat(  DateAdd("m", -1, Now), "yyyy, mm, dd, hh, nn")
%>
<!DOCTYPE HTML>
<html>
<head>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
		<!-- <meta http-equiv='refresh' content='60; url=http://ufa-qos01ow/vsp/main1.asp'> -->
		<!-- 1. Add these JavaScript inclusions in the head of your page -->
		<script type="text/javascript" src="js/jquery.min.js"></script>
		<script type="text/javascript" src="js/highcharts.js"></script>
		<script type="text/javascript" src="js/themes/gray.js"></script>
		<!-- 2. Add the JavaScript to initialize the chart on document ready -->
		<script type="text/javascript">
		
			var chart1;
			var chart2;
			var chart3;
			var chartA;
			var chartB;
			var chartC;
			var FlagOut=1;
			
			// Первый график
			$(document).ready(function() {
				chart1 = new Highcharts.Chart({
					chart: {
						renderTo: 'container1',
						type: 'line',
						marginLeft: 160,
						marginRight: 60
					},
					colors: ["<% =Color1 %>", "#99CCFF"],
					credits: {enabled: false},
					legend:  {enabled: false},
					tooltip: {enabled: false},
					title:   {align: 'right', text: 'Неуспешные операции'},
					// subtitle: {align: 'left', text: 'Примечание:'},
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
						max: <% if Value1max < Main1_SetHiHi then Response.Write(Main1_SetHiHi) else Response.Write(Replace(Value1max, ",", ".")) end if%>,
						title: {
							text: null
						},
						labels: {
							formatter: function() {return this.value +'%'; }
						},
						allowDecimals: false,
						plotLines: [{
							value: 0,
							width: 1,
							color: '#808080'
						}]
					},
					{
						min: 0,
						title: {
							text: null
						},
						opposite: true,
						allowDecimals: false,
						lineColor: '#99CCFF',
						labels: {
							step: 2,
							style: {color: '#99CCFF'}
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
								enabled: true,
								formatter: function() {
									FlagOut=-FlagOut;
									if (FlagOut==1) { return this.y; }
								}
							},
							enableMouseTracking: false
						}
					},
					series: [
					{
<%
Rs.Open "SELECT dateadd(month,-1,DT) DT, TagID, [Value] FROM Tags_History WHERE (TagID='Main1') AND (DT > GETDATE()-1.0/6) ORDER BY DT", Conn
if not Rs.Eof then
  Response.Write("data: [")
  Response.Write("[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Replace(FormatNumber(Rs.Fields("Value"), 1, -1, 0, 0), ",", ".")&"]")
  Rs.MoveNext
  do while not Rs.Eof
    Response.Write(","&vbCrLf&"[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Replace(FormatNumber(Rs.Fields("Value"), 1, -1, 0, 0), ",", ".")&"]")
    Rs.MoveNext
  loop
  Response.Write("]")
end if
Rs.Close
%>
					},
					{
						yAxis: 1
<%
LastTime1=Now
Rs.Open "SELECT top 1 DT FROM Tags_History WHERE (TagID='Main1all') AND (DT > GETDATE()-1.0/6) ORDER BY DT desc", Conn
if not Rs.Eof then
	LastTime1=Rs.Fields("DT")
end if
Rs.Close

Rs.Open "SELECT dateadd(month,-1,DT) DT, TagID, [Value] FROM Tags_History WHERE (TagID='Main1all') AND (DT > GETDATE()-1.0/6) ORDER BY DT", Conn
'LastTime1=Now
if not Rs.Eof then
  Response.Write(", data: [")
  Response.Write("[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Replace(FormatNumber(Rs.Fields("Value"), 1, -1, 0, 0), ",", ".")&"]")
  'LastTime1=Rs.Fields("DT")
  Rs.MoveNext
  do while not Rs.Eof
    Response.Write(","&vbCrLf&"[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Replace(FormatNumber(Rs.Fields("Value"), 1, -1, 0, 0), ",", ".")&"]")
	'LastTime1=Rs.Fields("DT")
    Rs.MoveNext
  loop
  Response.Write("]")
end if
Rs.Close
%>

					}
					]
				});
				
				chart1.yAxis[0].addPlotLine({
					value: <% =Main1_SetHiHi%>, color: '#FF9900', dashStyle: 'Dash', width: 2, id: 'plot-line-1'
				});
			
				// Второй график
				FlagOut=1;
				chart2 = new Highcharts.Chart({
					chart: {
						renderTo: 'container2',
						// defaultSeriesType: 'column'
						type: 'line'
					},
					colors: ['#66FFFF', '#FF66FF', '#FFFF66', '#d7b6b6'],
					credits: {enabled: false},
					legend:  {enabled: false},
					tooltip: {enabled: false},
					title:   {align: 'center', text: 'ATM'},
					//subtitle: {align: 'left', text: 'Примечание:'},
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
						max: <% =Replace(Value2A_max, ",", ".") %>,
						title: {
							text: null
						},
						lineColor: '#66FFFF',
						labels: {
							formatter: function() {return this.value +'%'; },
							style: {color: '#66FFFF'}
						},
						allowDecimals: false,
						plotLines: [{
							value: 0,
							width: 1,
							color: '#808080'
						}]
					},
					{
					    min: 0,
						title: {
							text: null
						},
						opposite: true,
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
									FlagOut=-FlagOut;
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
						name: 'Все АТМ'
<%
LastTime2=Now
Rs.Open "SELECT top 1 DT FROM Tags_History WHERE (TagID='Main2All') AND (DT > GETDATE()-1.0/6) ORDER BY DT desc", Conn
if not Rs.Eof then
	LastTime2=Rs.Fields("DT")
end if
Rs.Close

Rs.Open "SELECT dateadd(month,-1,DT) DT, TagID, [Value] FROM Tags_History WHERE (TagID='Main2All') AND (DT > GETDATE()-1.0/6) ORDER BY DT", Conn
'LastTime2=Now
if not Rs.Eof then
  Response.Write(", data: [")
  Response.Write("[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Rs.Fields("Value")&"]")
  Rs.MoveNext
  do while not Rs.Eof
    Response.Write(","&vbCrLf&"[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Rs.Fields("Value")&"]")
	'LastTime2=Rs.Fields("DT")
    Rs.MoveNext
  loop
  Response.Write("]")
end if
Rs.Close
%>
					}, {
						name: 'АТМ филиалов',
						yAxis: 1
<%
Rs.Open "SELECT dateadd(month,-1,DT) DT, TagID, [Value] FROM Tags_History WHERE (TagID='Main2Fil') AND (DT > GETDATE()-1.0/6) ORDER BY DT", Conn
if not Rs.Eof then
  Response.Write(", data: [")
  Response.Write("[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Rs.Fields("Value")&"]")
  Rs.MoveNext
  do while not Rs.Eof
    Response.Write(","&vbCrLf&"[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Rs.Fields("Value")&"]")
    Rs.MoveNext
  loop
  Response.Write("]")
end if
Rs.Close
%>
					}, {
						name: 'АТМ через ЦСП',
						yAxis: 1
<%
Rs.Open "SELECT dateadd(month,-1,DT) DT, TagID, [Value] FROM Tags_History WHERE (TagID='Main2Centr') AND (DT > GETDATE()-1.0/6) ORDER BY DT", Conn
if not Rs.Eof then
  Response.Write(", data: [")
  Response.Write("[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Rs.Fields("Value")&"]")
  Rs.MoveNext
  do while not Rs.Eof
    Response.Write(","&vbCrLf&"[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Rs.Fields("Value")&"]")
    Rs.MoveNext
  loop
  Response.Write("]")
end if
Rs.Close
%>
					},
					{
						name: 'Все АТМ за последние 24 часа',
<%
Rs.Open "SELECT dateadd(month,-1,DT) DT, TagID, [Value] FROM Tags_History WHERE (TagID='Main2All24') AND (DT > GETDATE()-1.0/6) ORDER BY DT", Conn
if not Rs.Eof then
  Response.Write(" data: [")
  Response.Write("[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Rs.Fields("Value")&"]")
  Rs.MoveNext
  do while not Rs.Eof
    Response.Write(","&vbCrLf&"[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Rs.Fields("Value")&"]")
    Rs.MoveNext
  loop
  Response.Write("]")
end if
Rs.Close
%>
					}
					]
				});
				
				chart2.yAxis[0].addPlotLine({ value: <% = Main2A24_SetHiHi %>, color: '#d7b6b6', dashStyle: 'Dash', width: 2, id: 'plot-line-1' });
				//chart2.yAxis[0].addPlotLine({ value: <% = Main2A_SetHiHi %>, color: '#66FFFF', dashStyle: 'Dash', width: 2, id: 'plot-line-1' });
				//chart2.yAxis[1].addPlotLine({ value: <% = Main2F_SetHiHi %>, color: '#FF66FF', dashStyle: 'Dash', width: 2, id: 'plot-line-1' });
				//chart2.yAxis[1].addPlotLine({ value: <% = Main2C_SetHiHi %>, color: '#FFFF66', dashStyle: 'Dash', width: 2, id: 'plot-line-1' });

				// график для БПТ
				FlagOut=1;
				chart4 = new Highcharts.Chart({
					chart: {
						renderTo: 'container2b',
						// defaultSeriesType: 'column'
						type: 'line'
					},
					colors: ['#66FFFF', '#FF66FF', '#d7b6b6'],
					credits: {enabled: false},
					legend:  {enabled: false},
					tooltip: {enabled: false},
					title:   {align: 'center', text: 'БПТ'},
					//subtitle: {align: 'left', text: 'Примечание:'},
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
						max: <% =Replace(Value2AllBPT_max, ",", ".") %>,
						title: {
							text: null
						},
						allowDecimals: false,
						lineColor: '#66FFFF',
						labels: {
							formatter: function() {return this.value +'%'; },
							style: {color: '#66FFFF'}
						},
						plotLines: [{
							value: 0,
							width: 1,
							color: '#808080'
						}]
					},
					{
					    min: 0,
						title: {
							text: null
						},
						opposite: true,
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
									FlagOut=-FlagOut;
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
						name: 'Все БПТ'
<%
LastTime4=Now
Rs.Open "SELECT top 1 DT FROM Tags_History WHERE (TagID='Main2AllBPT') AND (DT > GETDATE()-1.0/6) ORDER BY DT desc", Conn
if not Rs.Eof then
	LastTime4=Rs.Fields("DT")
end if
Rs.Close

Rs.Open "SELECT dateadd(month,-1,DT) DT, TagID, [Value] FROM Tags_History WHERE (TagID='Main2AllBPT') AND (DT > GETDATE()-1.0/6) ORDER BY DT", Conn
'LastTime4=Now
if not Rs.Eof then
  Response.Write(", data: [")
  Response.Write("[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Rs.Fields("Value")&"]")
  Rs.MoveNext
  do while not Rs.Eof
    Response.Write(","&vbCrLf&"[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Rs.Fields("Value")&"]")
	'LastTime4=Rs.Fields("DT")
    Rs.MoveNext
  loop
  Response.Write("]")
end if
Rs.Close
%>
					}, {
						name: 'БПТ филиалов',
						yAxis: 1
<%
Rs.Open "SELECT dateadd(month,-1,DT) DT, TagID, [Value] FROM Tags_History WHERE (TagID='Main2FilBPT') AND (DT > GETDATE()-1.0/6) ORDER BY DT", Conn
if not Rs.Eof then
  Response.Write(", data: [")
  Response.Write("[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Rs.Fields("Value")&"]")
  Rs.MoveNext
  do while not Rs.Eof
    Response.Write(","&vbCrLf&"[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Rs.Fields("Value")&"]")
    Rs.MoveNext
  loop
  Response.Write("]")
end if
Rs.Close
%>
					},
					{
						name: 'Все БПТ за последние 24 часа',
<%
Rs.Open "SELECT dateadd(month,-1,DT) DT, TagID, [Value] FROM Tags_History WHERE (TagID='Main2AllBPT24') AND (DT > GETDATE()-1.0/6) ORDER BY DT", Conn
if not Rs.Eof then
  Response.Write(" data: [")
  Response.Write("[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Rs.Fields("Value")&"]")
  Rs.MoveNext
  do while not Rs.Eof
    Response.Write(","&vbCrLf&"[Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), "&Rs.Fields("Value")&"]")
    Rs.MoveNext
  loop
  Response.Write("]")
end if
Rs.Close
%>
					}
					]
				});
				
				chart4.yAxis[0].addPlotLine({ value: <% = Main2AllBPT24_SetHiHi %>, color: '#d7b6b6', dashStyle: 'Dash', width: 2, id: 'plot-line-1' });
				//chart4.yAxis[0].addPlotLine({ value: <% = Main2AllBPT_SetHiHi %>, color: '#66FFFF', dashStyle: 'Dash', width: 2, id: 'plot-line-1' });
				//chart4.yAxis[1].addPlotLine({ value: <% = Main2FilBPT_SetHiHi %>, color: '#FF66FF', dashStyle: 'Dash', width: 2, id: 'plot-line-1' });
				
				// Третий график
				chart3 = new Highcharts.Chart({
					chart: {
						renderTo: 'container3',
						marginRight: 60
					},
					credits: {enabled: false},
					legend:  {enabled: false},
					tooltip: {enabled: false},
					title:   {align: 'right', text: 'Упавшие каналы: <span style="color: #FF0000"><% =Value3Down %> </span>'},
					subtitle: {align: 'left', text: 'Восстановились: <span style="color: #00FF00"><% =Value3Up %> </span>'},
					xAxis: {
						max: Date.UTC(<% =CurrentTime %>),
						type: 'datetime',
						dateTimeLabelFormats: { // don't display the dummy year
							hour: '%H:%M'
						}
					},
					yAxis: {
					    min: 0,
					    max: 6,
						tickInterval: 1,
						labels: { formatter: function() 
								  {
								    var t;
									if (this.value == 0) {t='<% =CID(5) %>'};
									if (this.value == 1) {t='<% =CID(4) %>'};
									if (this.value == 2) {t='<% =CID(3) %>'};
									if (this.value == 3) {t='<% =CID(2) %>'};
									if (this.value == 4) {t='<% =CID(1) %>'};
									if (this.value == 5) {t='Инф.'};
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
					},
					series: [
					{
						name: 'Динамика состояния',
						type: 'scatter',
						data: [
<%
LastTime3=Now
Rs.Open "SELECT top 1 DT FROM Tags_History WHERE (TagID='Main3') AND (DT > GETDATE()-1.0/6) ORDER BY DT desc", Conn
if not Rs.Eof then
	LastTime3=Rs.Fields("DT")
end if
Rs.Close

Rs.Open "SELECT dateadd(month,-1,DT) DT, TagID, -1*[Value] as [Value] FROM Tags_History WHERE (TagID='Main3') AND (DT > GETDATE()-1.0/6) ORDER BY DT", Conn
'Rs.Open "SELECT DT, TagID, -1*[Value] as [Value] FROM Tags_History WHERE (TagID='Main3') AND (DT > GETDATE()-1.0/6) ORDER BY DT", Conn
'LastTime3=Now
if not Rs.Eof then
  m="marker: {fillColor: '#00FF00', lineColor: '#00FF00', radius: 3}, "
  if Rs.Fields("Value")=1 then m="marker: {fillColor: '#99FF99', lineColor: '#99FF99', radius: 3}, " end if
  if Rs.Fields("Value")=2 then m="marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 3}, " end if
  Response.Write("{"&m&"x: Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), y: 5.5}")
  Rs.MoveNext
  do while not Rs.Eof
    m="marker: {fillColor: '#00FF00', lineColor: '#00FF00', radius: 3}, "
    if Rs.Fields("Value")=1 then m="marker: {fillColor: '#99FF99', lineColor: '#99FF99', radius: 3}, " end if
    if Rs.Fields("Value")=2 then m="marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 3}, " end if
    Response.Write(","&vbCrLf&"{"&m&"x: Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), y: 5.5}")
	'LastTime3=Rs.Fields("DT")
    Rs.MoveNext
  loop
end if
Rs.Close
%>
						]
					},
					{
						name: 'Кол-во без связи',
						type: 'line',
						data: [
<%
Rs.Open "SELECT dateadd(month,-1,DT) DT, TagID, [Value] FROM Tags_History WHERE (TagID='Main3down') AND (DT > GETDATE()-1.0/6) ORDER BY DT", Conn
if not Rs.Eof then
  v=Rs.Fields("Value")
  if v>1 then vs="name: '"&v&"', " else vs="" end if
  Response.Write("{"&vs&"x: Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), y: 5.5}")
  Rs.MoveNext
  do while not Rs.Eof
    v=Rs.Fields("Value")
    if v>1 then vs="name: '"&v&"', " else vs="" end if
    Response.Write(","&vbCrLf&"{"&vs&"x: Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), y: 5.5}")
    Rs.MoveNext
  loop
end if
Rs.Close
%>
						]
					}
					<%  =AllSeries 	%> 

 				]
				});

				chartA = new Highcharts.Chart({
					chart:   {renderTo: 'containerA', type: 'line', margin: [0, 0, 0, 0] },
					credits: {enabled: false},	legend:  {enabled: false},	tooltip: {enabled: false},	title:   {text: ''}
				});
				chartA.renderer.circle(150, 150, 90).attr({
					fill: '<% =Color1 %>',
					stroke: '<% =Color1 %>'
				}).add();
				
			<%
				if DateDiff("n", LastTime1, Now())>20 then
					Response.Write("chartA.renderer.image('q.gif', 75, 75, 150, 150).add();")
				end if
				%>
				chartB = new Highcharts.Chart({
					chart:   {renderTo: 'containerB', type: 'line', margin: [0, 0, 0, 0] },
					credits: {enabled: false},	legend:  {enabled: false},	tooltip: {enabled: false},	title:   {text: ''}
				});
				chartB.renderer.circle(150, 150, 90).attr({
					fill: '<% =Color2 %>',
					stroke: '<% =Color2 %>'
				}).add();
				<%
				if (DateDiff("n", LastTime2, Now())>20) or (DateDiff("n", LastTime4, Now())>20) then
					Response.Write("chartB.renderer.image('q.gif', 75, 75, 150, 150).add();")
				end if
				%>

				chartC = new Highcharts.Chart({
					chart:   {renderTo: 'containerC', type: 'line', margin: [0, 0, 0, 0] },
					credits: {enabled: false},	legend:  {enabled: false},	tooltip: {enabled: false},	title:   {text: ''}
				});
				chartC.renderer.circle(150, 150, 90).attr({
					fill: '<% =Color3 %>',
					stroke: '<% =Color3 %>'
				}).add();
				<%
				if DateDiff("n", LastTime3, Now())>20 then
					Response.Write("chartC.renderer.image('q.gif', 75, 75, 150, 150).add();")
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
<table border="0" width="1900" height="1060px" style="border: none;">
	<tr>
		<td style="border: none;"><div id="containerA"  style="width: 300px; height: 300px; margin: 0 auto"></div></td>
		<td style="border: none;">
		  <div id="containerAA" style="width: 500px; height: 300px; margin: 0 auto; font-size: 22pt;">
		    <table border="0" height="90%" width="100%" >
			  <tr><td style="text-align: left;" nowrap><% =Value1Date %></td></tr>
			  <tr><td style="text-align: left" nowrap>Всего операций:</td></tr>
			  <tr><td style="text-align: right; font-size: 40pt; font-weight: 600;"><% =Value1all %></td></tr>
			  <tr><td style="text-align: left">Неуспешных:</td></tr>
			  <tr><td style="text-align: right; font-size: 40pt; font-weight: 600;"><% =Value1&"<br>"&FormatNumber(Value1proc, 2, -1, 0, 0)&"%" %></td></tr>
			</table>
		  </div></td>
		<td style="border: none;" colspan="2"><div id="container1"  style="width: 1100px; height: 300px; margin: 0 auto"></div></td>
	</tr>
	<tr>
		<td style="border: none;"><div id="containerC"  style="width: 300px; height: 300px; margin: 0 auto"></div></td>
		<td style="border: none;">
		  <div id="containerCC" style="width: 500px; height: 300px; margin: 0 auto; font-size: 24pt;">
		    <table border="0" height="90%" width="100%" cellspacing="5">
			  <tr><td style="text-align: left" nowrap>Активных каналов:</td></tr>
			  <tr><td style="text-align: right; font-size: 48pt; font-weight: 700;"><% =Value3Total %></td></tr>
			  <tr><td style="text-align: left">Не активных:</td></tr>
			  <tr><td style="text-align: right; font-size: 48pt; font-weight: 700;"><% =Value3 %></td></tr>
			</table>
		  </div></td>
		<td style="border: none;" colspan="2"><div id="container3"  style="width: 1100px; height: 300px; margin: 0 auto"></div></td>
	</tr>
	<tr>
		<td style="border: none;"><div id="containerB"  style="width: 300px; height: 300px; margin: 0 auto"></div></td>
		<td style="border: none;">
		  <div id="containerBB" style="width: 500px; height: 300px; margin: 0 auto; font-size: 20pt;">
			<table border="0" height="90%" width="100%" cellspacing="0">
			  <tr><td>&nbsp;</td>
			      <td style="border: solid 1px #C0C0C0; font-size: 24pt">&nbsp;&nbsp;&nbsp;АТМ&nbsp;&nbsp;&nbsp;</td>
				  <td style="border: solid 1px #C0C0C0; font-size: 24pt">&nbsp;&nbsp;&nbsp;БПТ&nbsp;&nbsp;&nbsp;</td>
			  </tr>
			  <tr>
			      <td class="Head" style="border: solid 1px #C0C0C0; color: #d7b6b6;" nowrap>LINK 24, %</td>
			      <td class="Txt" style="border: solid 1px #C0C0C0; <% =IIF(Colors2(4,1)="", "", "color: #000000; background-color: "&Colors2(4,1)&";") %>"><% =Text2(4,1) %></td>
			      <td class="Txt" style="border: solid 1px #C0C0C0; <% =IIF(Colors2(4,2)="", "", "color: #000000; background-color: "&Colors2(4,2)&";") %>"><% =Text2(4,2) %></td>
			  </tr>
			  <tr>
			      <td class="Head" style="border: solid 1px #C0C0C0; color: #66FFFF;" nowrap>ПЦ, %</td>
			      <td class="Txt" style="border: solid 1px #C0C0C0; <% =IIF(Colors2(1,1)="", "", "color: #000000; background-color: "&Colors2(1,1)&";") %>"><% =Text2(1,1) %></td>
			      <td class="Txt" style="border: solid 1px #C0C0C0; <% =IIF(Colors2(1,2)="", "", "color: #000000; background-color: "&Colors2(1,2)&";") %>"><% =Text2(1,2) %></td>
			  </tr>
			  <tr>
				  <td class="Head" style="border: solid 1px #C0C0C0; color: #FF66FF;" nowrap>&nbsp;ФИ, шт.&nbsp;</td>
			      <td class="Txt" style="border: solid 1px #C0C0C0; <% =IIF(Colors2(2,1)="", "", "color: #000000; background-color: "&Colors2(2,1)&";") %>"><% =Text2(2,1) %></td>
			      <td class="Txt" style="border: solid 1px #C0C0C0; <% =IIF(Colors2(2,2)="", "", "color: #000000; background-color: "&Colors2(2,2)&";") %>"><% =Text2(2,2) %></td>
			  </tr>
			  <tr>
				  <td class="Head" style="border: solid 1px #C0C0C0; color: #FFFF66;" nowrap>ЦСП, шт</td>
			      <td class="Txt" style="border: solid 1px #C0C0C0; <% =IIF(Colors2(3,1)="", "", "color: #000000; background-color: "&Colors2(3,1)&";") %>"><% =Text2(3,1) %></td>
				  <td class="Txt" style="border: solid 1px #C0C0C0;">&nbsp;</td>
			  </tr>
			</table>
		  </div></td>
		<td style="border: none;"><div id="container2"   style="width: 550px; height: 300px; margin: 0 auto"></div></td>
		<td style="border: none;"><div id="container2b"  style="width: 550px; height: 300px; margin: 0 auto"></div></td>
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
