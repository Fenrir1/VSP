<%@ language="VBScript"%><!-- #include file="const.asp" --><!-- #include file="common.asp" -->
<%
' Данный модуль является вспомогательным для остальных модулей. 
' Предоставляет наборы данных из БД или готовые html вставки.
set Conn=Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeout=180
Conn.CommandTimeout=10
Conn.Open(ConnectionString)
set Rs=Server.CreateObject("ADODB.Recordset")
set Cmd=Server.CreateObject("ADODB.Command")
Cmd.ActiveConnection=Conn
Cmd.CommandType=adCmdText

ds=Request("ds")
tag=Request("tag")
prm=Request("prm")
prm2=Request("prm2")

' Змена , на .
function d(v)
	d = replace(v,",",".")
end function

if ds="OperationsHistory" then
  SQL_="SELECT DT_FILE, SUM(QUANTITY) AS Q FROM vw_NV_Operations_History "&_
  "WHERE [NAME]='"&tag&"' AND DIRECTION='"&prm&"' AND IsFailed=0 AND (DT_FILE>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT_FILE<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
  "GROUP BY DT_FILE ORDER BY DT_FILE"
  Rs.Open SQL_, Conn
  if not Rs.Eof then
    Response.Write(DateTimeFormat(Rs.Fields(0), "mm/dd/yyyy hh:nn:ss")&" UTC,"&Rs.Fields(1))
    Rs.MoveNext
    do while not Rs.Eof
      Response.Write(","&DateTimeFormat(Rs.Fields(0), "mm/dd/yyyy hh:nn:ss")&" UTC,"&Rs.Fields(1))
      Rs.MoveNext
    loop
  end if
  Rs.Close
  Response.Write("~")
  SQL_="SELECT DT_FILE, SUM(QUANTITY) AS Q FROM vw_NV_Operations_History "&_
  "WHERE [NAME]='"&tag&"' AND DIRECTION='"&prm&"' AND IsFailed=1 AND (DT_FILE>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT_FILE<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
  "GROUP BY DT_FILE ORDER BY DT_FILE"
  Rs.Open SQL_, Conn
  if not Rs.Eof then
    Response.Write(DateTimeFormat(Rs.Fields(0), "mm/dd/yyyy hh:nn:ss")&" UTC,"&Rs.Fields(1))
    Rs.MoveNext
    do while not Rs.Eof
      Response.Write(","&DateTimeFormat(Rs.Fields(0), "mm/dd/yyyy hh:nn:ss")&" UTC,"&Rs.Fields(1))
      Rs.MoveNext
    loop
  end if
  Rs.Close
  Response.Write("~")
  Response.Write(DateTimeFormat(Int(Now)+prm2, "mm/dd/yyyy hh:nn:ss")&" UTC,0,"&DateTimeFormat(Int(Now)+prm2+1-1/86400, "mm/dd/yyyy hh:nn:ss")&" UTC,0")
end if

if ds="OperationsDiagramm1" then
  Cmd.CommandText="EXEC [sp_Diagram_NV] @DS="&tag&", @OnTime='"&prm&"'"
  set Rs=Cmd.Execute
  if tag=1 then
    TotalOperations=0
    CurrentProc=0
    gr="underfound"
    series1=""
    series2=""
    if not Rs.Eof then DT_FILE=Rs.Fields(4) end if
    do while not Rs.Eof
      TotalOperations=TotalOperations+Rs.Fields(2)
      if gr=Rs.Fields(0) then
        CurrentProc=CurrentProc+Rs.Fields(3)
    	series1=series1&IIF(gr="", "?", replace(gr, " ", "<br />"))&", "&replace(CurrentProc, ",", ".")&","
      end if
      if gr<>Rs.Fields(0) then 
        CurrentProc=Rs.Fields(3)
    	gr=Rs.Fields(0)
      end if
      if Rs.Fields(1)=0 then ' name / y / color
        series2=series2&","&replace(FormatNumber(Rs.Fields(3), 2), ",", ".")&",""#00CC00""|||"
      else
        series2=series2&IIF(Rs.Fields(0)="", "?", Rs.Fields(0))&":<br />"&Rs.Fields(2)&","&replace(FormatNumber(Rs.Fields(3), 2), ",", ".")&",""#FF3300""|||"
      end if
      Rs.MoveNext
    loop
    series1=left(series1, len(series1)-1)
    series2=left(series2, len(series2)-3)
    Response.Write(series1)
    Response.Write("~")
    Response.Write(series2)
  end if
  if (tag=2) or (tag=3) then
    series3=""
    do while not Rs.Eof
  	  series3=series3&Rs.Fields(0)&","&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&","
      Rs.MoveNext
    loop
    series3=left(series3, len(series3)-1)
	Response.Write(series3)
  end if
  Rs.Close
end if

if ds="OperationsDiagramm11" then
  Cmd.CommandText="EXEC [sp_Diagram_NV] @DS="&tag&", @OnTime='"&prm&"'"
  set Rs=Cmd.Execute
  if tag=1 then
    TotalOperations=0
    CurrentProc=0
    gr="underfound"
    series1=""
    series2=""
    if not Rs.Eof then DT_FILE=Rs.Fields(4) end if
    do while not Rs.Eof
      TotalOperations=TotalOperations+Rs.Fields(2)
      if gr=Rs.Fields(0) then
        CurrentProc=CurrentProc+Rs.Fields(3)
    	series1=series1&IIF(gr="", "?", replace(gr, " ", "<br />"))&", "&replace(CurrentProc, ",", ".")&","
      end if
      if gr<>Rs.Fields(0) then 
        CurrentProc=Rs.Fields(3)
    	gr=Rs.Fields(0)
      end if
      if Rs.Fields(1)=0 then ' name / y / color
        series2=series2&","&replace(FormatNumber(Rs.Fields(3), 2), ",", ".")&",""#00CC00""|||"
      else
        series2=series2&IIF(Rs.Fields(0)="", "?", Rs.Fields(0))&":<br />"&Rs.Fields(2)&","&replace(FormatNumber(Rs.Fields(3), 2), ",", ".")&",""#FF3300""|||"
      end if
      Rs.MoveNext
    loop
    series1=left(series1, len(series1)-1)
    series2=left(series2, len(series2)-3)
    Response.Write(series1)
    Response.Write("~")
    Response.Write(series2)
  end if
  if (tag=11) then
    series3=""
    do while not Rs.Eof
      if Rs.Fields(0)="Другие" then
		series3=series3&"{name: '"&Rs.Fields(0)&":<br />"&replace(FormatNumber(Rs.Fields(1), 0, -1), ",", ".")&" / ', y: "&replace(FormatNumber(Rs.Fields(2), 0, -1), ",", ".")&"},"
	  else
		series3=series3&"{name: '"&Rs.Fields(0)&": "&replace(FormatNumber(Rs.Fields(1), 0, -1), ",", ".")&" / ', y: "&replace(FormatNumber(Rs.Fields(2), 0, -1), ",", ".")&"},"
	  end if
      Rs.MoveNext
    loop
    series3=left(series3, len(series3)-1)
	Response.Write(series3)
  end if
  if (tag=2) or (tag=3) then
    series3=""
    do while not Rs.Eof
  	  series3=series3&Rs.Fields(0)&","&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&","
      Rs.MoveNext
    loop
    series3=left(series3, len(series3)-1)
	Response.Write(series3)
  end if
  Rs.Close
end if

if ds="OperationsAtTime" then
  Rs.Open "SELECT [CHANNEL], [NAME], [DIRECTION], [RESPONSE_CODE], [QUANTITY], C.Resp_text FROM NV_Operations_History AS O LEFT OUTER JOIN V_Resp_code AS C ON O.RESPONSE_CODE=C.Resp_code WHERE (O.DT_File='"&prm&"') AND (ISNULL(C.IsFailed, 0)<>0) ORDER BY 2, 3", Conn
  if not Rs.Eof then
    do while not Rs.Eof
      Response.Write("<tr><td>"&Rs.Fields(0)&"</td><td style='text-align: left;'>"&Rs.Fields(1)&"</td><td>"&Rs.Fields(2)&"</td><td>"&Rs.Fields(3)&"</td><td>"&Rs.Fields(4)&"</td><td style='text-align: left;'>"&Rs.Fields(5)&"</td></tr>")
      Rs.MoveNext
    loop
  else
    Response.Write("<tr><td colspan=6>Нет каналов с неуспешными операциями</td></tr>")
  end if
  Rs.Close
end if

if ds="ChannelHistory" then
	if tag="~" and prm="Table" then
	  ' Формирование таблицы
		Response.Write("	    <table cellpadding=""0"" cellspacing=""0"" width=""670px"" style=""font-size: 10pt; table-layout: fixed"">")
		Response.Write("		<colgroup><col width=""55px""><col width=""200px""><col width=""85px""><col width=""165px""><col width=""165px""></colgroup>")
		Response.Write("		<tbody>")

		dim series(8), CID(8), CNM(8)
		for i=1 to 8
			series(i)=""
			CID(i)=""
			CNM(i)=""
		next
		L=0

		SQL_="SELECT TOP (100) PERCENT A.CHANNEL_ID, A.CHANNEL, A.Qdown, A.LastDown, B.DT, B.VALUE FROM "&_
		"(SELECT CHANNEL_ID, CHANNEL, COUNT(*) AS Qdown, MAX(DT) AS LastDown  "&_
		"FROM vw_Channel_History WHERE (DT>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) AND (VALUE=0)  "&_
		"GROUP BY CHANNEL_ID, CHANNEL) AS A LEFT OUTER JOIN  "&_
		"(SELECT DT, CHANNEL_ID, CHANNEL, [VALUE] FROM vw_Channel_History AS vw_Channel_History_2 "&_
		"WHERE (DT = (SELECT MAX(DT) AS Expr1 FROM vw_Channel_History AS vw_Channel_History_1))) AS B ON A.CHANNEL_ID = B.CHANNEL_ID "&_
		"ORDER BY 3 desc, 2"
		Rs.Open SQL_, Conn
		if not Rs.Eof then
			CHID1=Rs.Fields(0)
			CHNM1=Rs.Fields(1)
			do while not Rs.Eof
				if Rs.Fields(5)=0 then cl=clError else cl=clNormal end if
				Response.Write("<tr id=""r"&Rs.Fields(0)&""" onclick=""ChGraph("&Rs.Fields(0)&", '"&Rs.Fields(1)&"', jsDate)"">"&_
				"<td>"&Rs.Fields(0)&"</td>"&_
				"<td>"&Rs.Fields(1)&"</td>"&_
				"<td>"&Rs.Fields(2)&"</td>"&_
				"<td>"&DateTimeFormat(Rs.Fields(3), "dd.mm.yy hh:mm:ss")&"</td>"&_
				"<td style='text-align: left; color: "&cl&"'>"&DateTimeFormat(Rs.Fields(4), "dd.mm.yy hh:mm:ss")&"</td>"&_
				"</tr>")
				if L<8 then
					L=L+1
					CID(L)=Rs.Fields(0)
					CNM(L)=Rs.Fields(1)
				end if
				Rs.MoveNext
			loop
		else
			Response.Write("<tr><td colspan=5>Нет данных</td></tr>")
			CHID1=0
			CHNM1=""
		end if
		Rs.Close
		Response.Write("		</tbody>")
		Response.Write("		</table>")
		Response.Write("~")
		cat="{""categories"":["
		for i=1 to 8
		  cat=cat&""""&CID(9-i)&" "&CNM(9-i)&""","
		next
		cat=cat&"""Инф.""]}"
		Response.Write(cat)
		Response.Write("~")

	series1=""
	SQL_="SELECT DT, TagID, -1*[Value] as [Value] FROM Tags_History WHERE (TagID='Main3') AND "&_
    "(DT>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
	" ORDER BY DT"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		m="marker: {fillColor: '#00FF00', lineColor: '#00FF00'}, "
		if Rs.Fields("Value")=1 then m="marker: {fillColor: '#99FF99', lineColor: '#99FF99'}, " end if
		if Rs.Fields("Value")=2 then m="marker: {fillColor: '#FF0000', lineColor: '#FF0000'}, " end if
		series1=series1&"{"&m&"x: Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), y: 8.5}"
		Rs.MoveNext
		do while not Rs.Eof
			m="marker: {fillColor: '#00FF00', lineColor: '#00FF00'}, "
			if Rs.Fields("Value")=1 then m="marker: {fillColor: '#99FF99', lineColor: '#99FF99'}, " end if
			if Rs.Fields("Value")=2 then m="marker: {fillColor: '#FF0000', lineColor: '#FF0000'}, " end if
			series1=series1&",{"&m&"x: Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), y: 8.5}"
			Rs.MoveNext
		loop
	end if
	Rs.Close
	Response.Write(series1&"~")

	series2=""
	SQL_= "SELECT * FROM Tags_History WHERE (TagID='Main3down') AND "&_
    "(DT>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
	" ORDER BY DT"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		v=Rs.Fields("Value")
		if v>1 then vs="name: '"&v&"', " else vs="" end if
		series2=series2&"{"&vs&"x: Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), y: 8.5}"
		Rs.MoveNext
		do while not Rs.Eof
			v=Rs.Fields("Value")
			if v>1 then vs="name: '"&v&"', " else vs="" end if
			series2=series2&","&vbCrLf&"{"&vs&"x: Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), y: 8.5}"
			Rs.MoveNext
		loop
	else
		series2=series2&"{x: Date.UTC("&DateTimeFormat(Int(Now+prm2), "yyyy, mm, dd")&"), y: 8.5}"
	end if
	Rs.Close
	Response.Write(series2&"~")

	for i=1 to 8
	  v=8.5-i
      v=replace(v, ",", ".")
	  if i<L then
	    series(i)=""
		SQL_="SELECT dateAdd(ss,-1*DATEPART(ss, DT),dateAdd(ms,-1*DATEPART(ms, DT),DT)) AS DT,[CHANNEL_ID],[CHANNEL],[VALUE] FROM vw_Channel_History "&_
		"WHERE (CHANNEL_ID="&CID(i)&") AND "&_
		"(DT>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
		" GROUP BY dateAdd(ss,-1*DATEPART(ss, DT),dateAdd(ms,-1*DATEPART(ms, DT),DT)),[CHANNEL_ID],[CHANNEL],[VALUE] ORDER BY DT"
		Rs.Open SQL_, Conn
		do while not Rs.Eof
          if Rs.Fields("Value")=0 then 
		    series(i)=series(i)&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"},"
	      end if
		  Rs.MoveNext
	    loop
	    Rs.Close
	    if series(i)<>"" then
		  series(i)=left(series(i), len(series(i))-1)
		  'series(i)="{ name: '"&CNM(i)&"', type: 'scatter', data: ["&series(i)&"]}"  
	    end if
	    Response.Write(series(i)&"~")
	  else
		Response.Write("{x: Date.UTC("&DateTimeFormat(Int(Now+prm2), "yyyy, mm, dd")&"), y: "&v&"}")
	  end if
	next
	
	else
	  ' Отрисовка динамики по выбранному каналу
		SQL_= "SELECT dateAdd(ss,-1*DATEPART(ss, DT),dateAdd(ms,-1*DATEPART(ms, DT),DT)) AS DT, [VALUE] FROM vw_Channel_History "&_
		"WHERE (DT>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
		"AND (CHANNEL_ID='"&tag&"') "&_
		"GROUP BY dateAdd(ss,-1*DATEPART(ss, DT),dateAdd(ms,-1*DATEPART(ms, DT),DT)),[VALUE] ORDER BY 1 DESC"
		Rs.Open SQL_, Conn
		if not Rs.Eof then
			Response.Write(DateTimeFormat(Rs.Fields(0), "mm/dd/yyyy hh:nn:ss")&" UTC,"&Rs.Fields(1))
			Rs.MoveNext
			do while not Rs.Eof
				Response.Write(","&DateTimeFormat(Rs.Fields(0), "mm/dd/yyyy hh:nn:ss")&" UTC,"&Rs.Fields(1))
				Rs.MoveNext
			loop
		else
			Response.Write(DateTimeFormat(Int(Now)+prm2, "mm/dd/yyyy hh:nn:ss")&" UTC,0,"&DateTimeFormat(Int(Now)+prm2+1-1/86400, "mm/dd/yyyy hh:nn:ss")&" UTC,0")
		end if
		Rs.Close
	end if
end if

if ds="AtmNoLink" then
  if prm="Graph" then
	SQL_= "SELECT DT_FILE, 100.0*ATM_LINK/ATM AS [VALUE] FROM AV_ATMStat_History "&_
	"WHERE (DT_FILE>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT_FILE<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
	"AND (BRANCH_CODE='"&tag&"') ORDER BY 1 DESC"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		Response.Write("{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
		Rs.MoveNext
		do while not Rs.Eof
			Response.Write(","&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
			Rs.MoveNext
		loop
	else
		Response.Write("{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2, "yyyy, mm, dd, hh, nn")&"), y: 0},{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2+1-1/86400, "yyyy, mm, dd, hh, nn")&"), y: 0}")
	end if
	Rs.Close

	Response.Write("~")

	SQL_="SELECT DT, [Value] FROM Tags_History WHERE (TagID='Main2All') AND "&_
	"(DT>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
	"ORDER BY DT"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		Response.Write("{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
		Rs.MoveNext
		do while not Rs.Eof
			Response.Write(","&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
			Rs.MoveNext
		loop
	else
		Response.Write("{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2, "yyyy, mm, dd, hh, nn")&"), y: 0},{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2+1-1/86400, "yyyy, mm, dd, hh, nn")&"), y: 0}")
	end if
	Rs.Close

	Response.Write("~")
	
	SQL_="SELECT DT, [Value] FROM Tags_History WHERE (TagID='Main2Fil') AND "&_
	"(DT>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
	"ORDER BY DT"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		Response.Write("{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
		Rs.MoveNext
		do while not Rs.Eof
			Response.Write(","&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
			Rs.MoveNext
		loop
	else
		Response.Write("{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2, "yyyy, mm, dd, hh, nn")&"), y: 0},{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2+1-1/86400, "yyyy, mm, dd, hh, nn")&"), y: 0}")
	end if
	Rs.Close
	
	Response.Write("~")
	
		SQL_= "SELECT DT_FILE, 100.0*ATM_LINK24/ATM AS [VALUE] FROM AV_ATMStat_History "&_
	"WHERE (DT_FILE>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT_FILE<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
	"AND (BRANCH_CODE='"&tag&"') ORDER BY 1 DESC"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		Response.Write("{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
		Rs.MoveNext
		do while not Rs.Eof
			Response.Write(","&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
			Rs.MoveNext
		loop
	else
		Response.Write("{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2, "yyyy, mm, dd, hh, nn")&"), y: 0},{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2+1-1/86400, "yyyy, mm, dd, hh, nn")&"), y: 0}")
	end if
	Rs.Close
	
  end if
  if prm="Table" then
	Response.Write("		    <table cellpadding=""0"" cellspacing=""0"" width=""800px"" style=""font-size: 10pt; table-layout: fixed"">")
	Response.Write("				<colgroup><col span=6 width=""58px""><col width=""440px""></colgroup>")
	Response.Write("				<tbody>")
  
	SQL_="SELECT [DT_FILE], A.[BRANCH_CODE], ISNULL([Name], A.[BRANCH_CODE]) AS [NAME],[ATM],[ATM_LINK],[ATM_ERR],[ATM_LINK_ERR],[ATM_LINK24]*100.0/[ATM] "&_
		 "FROM [AV_ATMStat_History] AS A LEFT OUTER JOIN V_Branch_code AS C ON A.[BRANCH_CODE]=C.[Branch_code] "&_
		 "WHERE (DT_FILE='"&prm2&"') ORDER BY 8 DESC"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		DT_FILE=Rs.Fields(0)
		BRANCH_CODE=Rs.Fields(1)
		BRANCH_NAME=Rs.Fields(2)
		do while not Rs.Eof
			Response.Write("<tr id=""r"&Rs.Fields(1)&""" onclick=""ChGraph('"&Rs.Fields(1)&"', '"&replace(Rs.Fields(2), """", "")&"', jsDate)""><td>"&Rs.Fields(1)&"</td>")
			for i=3 to 6
			  Response.Write("<td>"&Rs.Fields(i)&"</td>")
			next
			ww=Round(440*cint(Rs.Fields(7))/100)
			Response.Write("<td>"&FormatNumber(Rs.Fields(7), 1)&"</td><td style=""text-align: left; background-color: #A0A0A0;"">"&_
			  "<img src=""d.gif"" width="""&ww&""" height=""16"" alt="""" style=""background-color: #CCFFFF; margin-top: 1px;"" />"&_
			  "<div style=""line-height: 18px; margin-top: -17px; margin-left: 2px; color: #000000; font-weight: 700"">"&Rs.Fields(2)&"</div>"&_
			  "</td></tr>"&vbCrLf)
			Rs.MoveNext
		loop
	else
		Response.Write("<tr><td colspan=""7"">Нет данных</td></tr>")
	end if
	Rs.Close
  
	Response.Write("				</tbody>")
	Response.Write("			</table>")
  end if
end if

if ds="AtmTypeLink" then
  if prm="Graph" then
	SQL_= "SELECT DT_FILE, 100.0*ATM_OFFLINE/ATM AS [VALUE] FROM LV_ATMStatLink_History "&_
	"WHERE (DT_FILE>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT_FILE<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
	"AND (LINK_TYPE='"&tag&"') ORDER BY 1 DESC"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		Response.Write("{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
		Rs.MoveNext
		do while not Rs.Eof
			Response.Write(","&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
			Rs.MoveNext
		loop
	else
		Response.Write("{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2, "yyyy, mm, dd, hh, nn")&"), y: 0},{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2+1-1/86400, "yyyy, mm, dd, hh, nn")&"), y: 0}")
	end if
	Rs.Close

	Response.Write("~")

	SQL_="SELECT DT, [Value] FROM Tags_History WHERE (TagID='Main2Centr') AND "&_
	"(DT>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
	"ORDER BY DT"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		Response.Write("{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
		Rs.MoveNext
		do while not Rs.Eof
			Response.Write(","&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
			Rs.MoveNext
		loop
	else
		Response.Write("{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2, "yyyy, mm, dd, hh, nn")&"), y: 0},{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2+1-1/86400, "yyyy, mm, dd, hh, nn")&"), y: 0}")
	end if
	Rs.Close
	
  end if
  if prm="Table" then
	Response.Write("		    <table cellpadding=""0"" cellspacing=""0"" width=""800px"" style=""font-size: 10pt; table-layout: fixed"">")
	Response.Write("				<colgroup><col width=""250px""><col span=3 width=""58px""><col width=""*""></colgroup>")
	Response.Write("				<tbody>")
  
	SQL_="SELECT [DT_FILE], [LINK_TYPE], [ATM], [ATM_OFFLINE], [IsCentralSchema], [ATM_OFFLINE]*100.0/[ATM] "&_
		 "FROM [LV_ATMStatLink_History] "&_
		 "WHERE (DT_FILE='"&prm2&"') ORDER BY 6 DESC"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		DT_FILE=Rs.Fields(0)
		LINK_TYPE=Rs.Fields(1)
		do while not Rs.Eof
			Response.Write("<tr id=""r"&Rs.Fields(1)&""" onclick=""ChGraph('"&Rs.Fields(1)&"', jsDate)""><td>"&Rs.Fields(1)&"</td>")
			for i=2 to 4
			  Response.Write("<td>"&Rs.Fields(i)&"</td>")
			next
			ww=Round(440*cint(Rs.Fields(5))/100)
			Response.Write("<td style=""text-align: left; background-color: #A0A0A0;"">"&_
			  "<img src=""d.gif"" width="""&ww&""" height=""16"" alt="""" style=""background-color: #CCFFFF; margin-top: 1px;"" />"&_
			  "<div style=""line-height: 18px; margin-top: -17px; margin-left: 2px; color: #000000; font-weight: 700"">"&FormatNumber(Rs.Fields(5), 1)&"</div>"&_
			  "</td></tr>"&vbCrLf)
			Rs.MoveNext
		loop
	else
		Response.Write("<tr><td colspan=""5"">Нет данных</td></tr>")
	end if
	Rs.Close
  
	Response.Write("				</tbody>")
	Response.Write("			</table>")
  end if
end if

if ds="BPTNoLink" then
  if prm="Graph" then
	SQL_= "SELECT DT_FILE, 100.0*BPT_LINK/BPT AS [VALUE] FROM TV_BPTStat_History "&_
	"WHERE (DT_FILE>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT_FILE<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
	"AND (BRANCH_CODE='"&tag&"') ORDER BY 1 DESC"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		Response.Write("{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
		Rs.MoveNext
		do while not Rs.Eof
			Response.Write(","&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
			Rs.MoveNext
		loop
	else
		Response.Write("{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2, "yyyy, mm, dd, hh, nn")&"), y: 0},{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2+1-1/86400, "yyyy, mm, dd, hh, nn")&"), y: 0}")
	end if
	Rs.Close

	Response.Write("~")

	SQL_="SELECT DT, [Value] FROM Tags_History WHERE (TagID='Main2AllBPT') AND "&_
	"(DT>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
	"ORDER BY DT"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		Response.Write("{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
		Rs.MoveNext
		do while not Rs.Eof
			Response.Write(","&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
			Rs.MoveNext
		loop
	else
		Response.Write("{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2, "yyyy, mm, dd, hh, nn")&"), y: 0},{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2+1-1/86400, "yyyy, mm, dd, hh, nn")&"), y: 0}")
	end if
	Rs.Close

	Response.Write("~")
	
	SQL_="SELECT DT, [Value] FROM Tags_History WHERE (TagID='Main2FilBPT') AND "&_
	"(DT>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
	"ORDER BY DT"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		Response.Write("{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
		Rs.MoveNext
		do while not Rs.Eof
			Response.Write(","&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
			Rs.MoveNext
		loop
	else
		Response.Write("{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2, "yyyy, mm, dd, hh, nn")&"), y: 0},{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2+1-1/86400, "yyyy, mm, dd, hh, nn")&"), y: 0}")
	end if
	Rs.Close
	
	Response.Write("~")

		SQL_= "SELECT DT_FILE, 100.0*BPT_LINK24/BPT AS [VALUE] FROM TV_BPTStat_History "&_
	"WHERE (DT_FILE>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT_FILE<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
	"AND (BRANCH_CODE='"&tag&"') ORDER BY 1 DESC"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		Response.Write("{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
		Rs.MoveNext
		do while not Rs.Eof
			Response.Write(","&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 2), ",", ".")&"}")
			Rs.MoveNext
		loop
	else
		Response.Write("{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2, "yyyy, mm, dd, hh, nn")&"), y: 0},{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2+1-1/86400, "yyyy, mm, dd, hh, nn")&"), y: 0}")
	end if
		Rs.Close
	
  end if
  if prm="Table" then
	Response.Write("		    <table cellpadding=""0"" cellspacing=""0"" width=""800px"" style=""font-size: 10pt; table-layout: fixed"">")
	Response.Write("				<colgroup><col span=6 width=""58px""><col width=""440px""></colgroup>")
	Response.Write("				<tbody>")
  
	SQL_="SELECT [DT_FILE], A.[BRANCH_CODE], ISNULL([Name], A.[BRANCH_CODE]) AS [NAME],[BPT],[BPT_LINK],[BPT_ERR],[BPT_LINK_ERR],[BPT_LINK24]*100.0/[BPT] "&_
		 "FROM [TV_BPTStat_History] AS A LEFT OUTER JOIN V_Branch_code AS C ON A.[BRANCH_CODE]=C.[Branch_code] "&_
		 "WHERE (DT_FILE='"&prm2&"') ORDER BY 8 DESC"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		DT_FILE=Rs.Fields(0)
		BRANCH_CODE=Rs.Fields(1)
		BRANCH_NAME=Rs.Fields(2)
		do while not Rs.Eof
			Response.Write("<tr id=""r"&Rs.Fields(1)&""" onclick=""ChGraph('"&Rs.Fields(1)&"', '"&replace(Rs.Fields(2), """", "")&"', jsDate)""><td>"&Rs.Fields(1)&"</td>")
			for i=3 to 6
			  Response.Write("<td>"&Rs.Fields(i)&"</td>")
			next
			ww=Round(440*cint(Rs.Fields(7))/100)
			Response.Write("<td>"&FormatNumber(Rs.Fields(7), 1)&"</td><td style=""text-align: left; background-color: #A0A0A0;"">"&_
			  "<img src=""d.gif"" width="""&ww&""" height=""16"" alt="""" style=""background-color: #CCFFFF; margin-top: 1px;"" />"&_
			  "<div style=""line-height: 18px; margin-top: -17px; margin-left: 2px; color: #000000; font-weight: 700"">"&Rs.Fields(2)&"</div>"&_
			  "</td></tr>"&vbCrLf)
			Rs.MoveNext
		loop
	else
		Response.Write("<tr><td colspan=""7"">Нет данных</td></tr>")
	end if
	Rs.Close
  
	Response.Write("				</tbody>")
	Response.Write("			</table>")
  end if
end if

if ds="SMSService" then
	SQL_="SELECT COUNT(*) FROM MV_SMSService"
	Rs.Open SQL_, Conn
	SMScount=Rs.Fields(0)
	Rs.Close
  SQL_="SELECT TOP "&SMScount&" [SERVER] FROM MV_SMSService ORDER BY [SERVER]"
  Rs.Open SQL_, Conn
  dim srv(10), fld(3) 'продумать динамическое обьявление массива для srv()
  srv(0)=""
  fld(0)="WAIT_COUNT"
  fld(1)="[REJECTED_COUNT]"
  fld(2)="[DECLINED_COUNT]"
smscounti=0
  do while not Rs.Eof
	if not Rs.Eof then
	  srv(smscounti)=Rs.Fields(0)
	  Response.Write(Rs.Fields(0)&"~")
	else
	  srv(smscounti)=""
	  Response.Write("~")
	end if  
	smscounti=smscounti+1
	Rs.MoveNext
	loop
  Rs.Close
  for i=0 to SMScount-1 
    if srv(i)<>"" then
	  for j=0 to 2 
		SQL_="SELECT DT_FILE, "&fld(j)&"*100/[ALL_COUNT] FROM MV_SMSService_History WHERE ([SERVER]='"&srv(i)&"') AND "&_
		"(DT_FILE>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&"))) AND (DT_FILE<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+("&prm2&")+1)) "&_
		"ORDER BY DT_FILE"
		Rs.Open SQL_, Conn
		if not Rs.Eof then
			Response.Write("{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 1), ",", ".")&"}")
			Rs.MoveNext
			do while not Rs.Eof
				Response.Write(","&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields(0), "yyyy, mm, dd, hh, nn")&"), y: "&replace(FormatNumber(Rs.Fields(1), 1), ",", ".")&"}")
				Rs.MoveNext
			loop
		else
			Response.Write("{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2, "yyyy, mm, dd, hh, nn")&"), y: 0},{x: Date.UTC("&DateTimeFormat(Int(Now)+prm2+1-1/86400, "yyyy, mm, dd, hh, nn")&"), y: 0}")
		end if
		Rs.Close
		Response.Write("~")
	  next
	end if
  next
end if

if ds="MessageTypeByID" then
  dim nm(5)
  nm(0)="Description"
  nm(1)="Instuction"
  nm(2)="Resolution"
  nm(3)="Recipients"
  nm(4)="Period"
  SQL_="SELECT TOP 1 Description,Instuction,Resolution,Recipients,Period FROM Messages_Type WHERE MsgID="&prm
  Rs.Open SQL_, Conn
  if not Rs.Eof then
    Response.Write("<table cellpadding=""0"" cellspacing=""0"" width=""100%"" style=""font-size: 10pt; border-top: solid 1px #4572A7; border-left: solid 1px #4572A7;"">")
	for i=0 to 4
      Response.Write("<tr><td width=""100px"" style=""border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;"">"&nm(i)&_
	  "</td><td style=""text-align: left; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;"">"&Rs.Fields(i)&"</td>")
	next
    Response.Write("</table>")
  end if
  Rs.Close
end if

if ds="Voc" then
	if prm=1 then
		SQL_="SELECT [Branch_code],[Name] FROM [V_Branch_code] ORDER BY [Name] "
		Rs.Open SQL_, Conn
		Response.Write("<table width=""600px"" cellpadding=""0"" cellspacing=""0"" style=""font-size: 10pt; table-layout: fixed"">")
		Response.Write("<colgroup><col width=""50px""><col width=""300px""><col width=""*""></colgroup>")
		Response.Write("<tbody><tr><th>Код</th><th>Финансовый институт</th><th>[ <span style=""font-weight: 400; text-decoration: underline;"" onclick='javascript: NewFin();'>Добавить</span> ]</th></tr></tbody></table>")
		Response.Write("<div style=""OVERFLOW-Y: auto; OVERFLOW-X: hidden; width: 600px; height: 300px;"">")
		Response.Write("<table width=""600px"" cellpadding=""0"" cellspacing=""0"" style=""font-size: 10pt; table-layout: fixed"">")
		Response.Write("<colgroup><col width=""50px""><col width=""300px""><col width=""*""></colgroup>")
		Response.Write("<tbody>")
		if not Rs.Eof then
			ColCount=Rs.Fields.Count
			do while not Rs.Eof
				id=Rs.Fields(0)
				Response.Write("<tr id=""r"&id&""" onclick=""javascript: selectRow('"&id&"');"">")
				for i=0 to ColCount-1
					Response.Write("<td>"&NBSP(Trim(Rs.Fields(i)))&"</td>")
				next
				Response.Write("<td><span onclick=""javascript: EditFin('"&id&"');"">[Изменить]</span> <span onclick=""javascript: DelFin('"&id&"');"">[Удалить]</span></td></tr>"&vbCrLf)
				Rs.MoveNext
			loop
		else
			Response.Write("<tr><td colspan=""3"">Нет данных</td></tr>")
		end if
		Response.Write("</tbody></table></div>")
		Rs.Close
	end if
	if prm=2 then
		SQL_="SELECT [Resp_code],[Resp_text],[IsFailed] FROM [V_Resp_code] ORDER BY 1"
		Rs.Open SQL_, Conn
		Response.Write("<table width=""900px"" cellpadding=""0"" cellspacing=""0"" style=""font-size: 10pt; table-layout: fixed"">")
		Response.Write("<colgroup><col width=""50px""><col width=""300px""><col width=""100px""><col width=""*""></colgroup>")
		Response.Write("<tbody><tr><th>Код</th><th>Описание</th><th>Критичность</th><th>[ <span style=""font-weight: 400; text-decoration: underline;"" onclick='javascript: NewRC();'>Добавить</span> ]</th></tr></tbody></table>")
		Response.Write("<div style=""OVERFLOW-Y: auto; OVERFLOW-X: hidden; OVERFLOW: auto; width: 900px; height: 300px;"">")
		Response.Write("<table width=""900px"" cellpadding=""0"" cellspacing=""0"" style=""font-size: 10pt; table-layout: fixed"">")
		Response.Write("<colgroup><col width=""50px""><col width=""300px""><col width=""100px""><col width=""*""></colgroup>")
		Response.Write("<tbody>")
		if not Rs.Eof then
			ColCount=Rs.Fields.Count
			do while not Rs.Eof
				id=Rs.Fields(0)
				Response.Write("<tr id=""r"&id&""" onclick=""javascript: selectRow('"&id&"');"">")
				for i=0 to ColCount-1
					Response.Write("<td>"&NBSP(Trim(Rs.Fields(i)))&"</td>")
				next
				Response.Write("<td><span onclick=""javascript: EditRC('"&id&"');"">[Изменить]</span> <span onclick=""javascript: DelRC('"&id&"');"">[Удалить]</span></td></tr>"&vbCrLf)
				Rs.MoveNext
			loop
		else
			Response.Write("<tr><td colspan=""3"">Нет данных</td></tr>")
		end if
		Response.Write("</tbody></table></div>")
		Rs.Close
	end if
	if prm=3 then
		SQL_="SELECT [User_ID],[User_Name],[User_Login],[Role],[Phone],[Email] FROM Users ORDER BY [User_Name] "
		Rs.Open SQL_, Conn
		Response.Write("<table width=""900px"" cellpadding=""0"" cellspacing=""0"" style=""font-size: 10pt; table-layout: fixed"">")
		Response.Write("<colgroup><col width=""150px""><col width=""150px""><col width=""120px""><col width=""90px""><col width=""150""><col width=""*""></colgroup>")
		Response.Write("<tbody><tr><th>Пользователь</th><th>Логин</th><th>Роль</th><th>Телефон</th><th>Эл.почта</th><th>[ <span style=""font-weight: 400; text-decoration: underline;"" onclick='javascript: NewUser();'>Новый пользователь</span> ]</th></tr></tbody></table>")
		Response.Write("<div style=""OVERFLOW-Y: auto; OVERFLOW-X: hidden; OVERFLOW: auto; width: 900px; height: 300px;"">")
		Response.Write("<table width=""900px"" cellpadding=""0"" cellspacing=""0"" style=""font-size: 10pt; table-layout: fixed"">")
		Response.Write("<colgroup><col width=""150px""><col width=""150px""><col width=""120px""><col width=""90px""><col width=""150""><col width=""*""></colgroup>")
		Response.Write("<tbody>")
		if not Rs.Eof then
			ColCount=Rs.Fields.Count
			do while not Rs.Eof
				id=Rs.Fields(0)
				Response.Write("<tr id=""r"&id&""" onclick=""javascript: selectRow('"&id&"');"">")
				for i=1 to ColCount-1
					if i=3 then
						Response.Write("<td>"&IIF(Rs.Fields(i)=1, "администратор", "пользователь")&"</td>")
					else
						Response.Write("<td>"&NBSP(Trim(Rs.Fields(i)))&"</td>")
					end if
				next
				Response.Write("<td><span onclick=""javascript: EditUser('"&id&"');"">[Изменить]</span> <span onclick=""javascript: DelUser('"&id&"', '"&Rs.Fields(2)&"');"">[Удалить]</span></td></tr>"&vbCrLf)
				Rs.MoveNext
			loop
		else
			Response.Write("<tr><td colspan=""6"">Нет данных</td></tr>")
		end if
		Response.Write("</tbody></table></div>")
		Rs.Close
	end if
	if prm=4 then
		SQL_="SELECT [TagID],[TagName],[SetHi],[SetHiHi],[FileID],[GroupName],[Prop_Crit],[Prop_Active],[Prop_SignOn],[Prop_Time] FROM [Tags] ORDER BY [TagID]"
		Rs.Open SQL_, Conn
		Response.Write("<table width=""1020px"" cellpadding=""0"" cellspacing=""0"" style=""font-size: 10pt; table-layout: fixed"">")
		Response.Write("<colgroup><col width=""80px""><col width=""280px""><col width=""80px""><col width=""75px""><col width=""40px""><col span=5 width=""50px""><col width=""*""></colgroup>")
		Response.Write("<tbody><tr><th>ID</th><th>Параметр</th><th>Допустимое значение</th><th>Критичное значение</th><th>Тип файла</th><th>Группа</th><th>Критич-ность</th><th>Актив-ность</th><th>SignOn</th><th>Период</th><th>[ <span style=""font-weight: 400; text-decoration: underline;"" onclick='javascript: NewTag();'>Новый параметр</span> ]</th></tr></tbody></table>")
		Response.Write("<div style=""OVERFLOW-Y: auto; OVERFLOW-X: hidden; OVERFLOW: auto; width: 1020px; height: 300px;"">")
		Response.Write("<table width=""1020px"" cellpadding=""0"" cellspacing=""0"" style=""font-size: 10pt; table-layout: fixed"">")
		Response.Write("<colgroup><col width=""80px""><col width=""280px""><col width=""80px""><col width=""75px""><col width=""40px""><col span=5 width=""50px""><col width=""*""></colgroup>")
		Response.Write("<tbody>")
		if not Rs.Eof then
			ColCount=Rs.Fields.Count
			do while not Rs.Eof
				id=Rs.Fields(0)
				Response.Write("<tr id=""r"&id&""" onclick=""javascript: selectRow('"&id&"');"">")
				for i=0 to ColCount-1
					Response.Write("<td>"&NBSP(Trim(Rs.Fields(i)))&"</td>")
				next
				Response.Write("<td><span onclick=""javascript: EditTag('"&id&"');"">[Изменить]</span> <span onclick=""javascript: DelTag('"&id&"');"">[Удалить]</span></td></tr>"&vbCrLf)
				Rs.MoveNext
			loop
		else
			Response.Write("<tr><td colspan=""6"">Нет данных</td></tr>")
		end if
		Response.Write("</tbody></table></div>")
		Rs.Close
	end if
	'-----------Channel Groups--------------------------------------------------------------
	if prm=5 then
		SQL_="SELECT File_Type, Channel_Group, Channel, ISNULL(Warning_Count,0) Warning_Count, ISNULL(Error_Count,0) Error_Count, ISNULL(Min_Count,0) Min_Count FROM  Channel_Config"
		Rs.Open SQL_, Conn
		Response.Write "<table cellpadding=""0"" cellspacing=""0"" style=""font-size: 10pt;"">"
		Response.Write "<tr><th width=""50px"" >Тип файла</th><th width=""100px"" >Группа каналов</th><th width=""200px"">Канал</th>"
		Response.Write "<th width=""50px"">Допустимое значение</th><th width=""50px"">Критичное значение</th>"
		Response.Write "<th width=""50px"" >Минимальный порог по количеству всех операций</th><th>&nbsp;</th></tr>"
		Response.Write("<tbody>")
		if not Rs.Eof then
			ColCount=Rs.Fields.Count
			do while not Rs.Eof
				id=Rs.Fields("Channel_Group")
				Response.Write("<tr id=""r"&id&""" onclick=""javascript: selectRow('"&id&"');"">")
				for i=0 to ColCount-1
					Response.Write("<td>"&NBSP(Trim(Rs.Fields(i)))&"</td>")
				next
				Response.Write("<td><span onclick=""javascript: EditChannelGroup('"&id&"');"">[Изменить]</span></td></tr>"&vbCrLf)
				Rs.MoveNext
			loop
		else
			Response.Write("<tr><td colspan=""6"">Нет данных</td></tr>")
		end if
		Response.Write "</tbody></table>"
		Rs.Close
	end if	
	
end if

'-------START: Channel Groups--------------------------------------------------------------
if ds="GetChannelGroup" then
	SQL_="SELECT File_Type, Channel_Group, Channel, ISNULL(Warning_Count,0) Warning_Count, ISNULL(Error_Count,0) Error_Count, ISNULL(Min_Count,0) Min_Count, ISNULL(Limit_Count,0) Limit_Count,  ISNULL(Lowactivity_start,0) Lowactivity_start, ISNULL(Lowactivity_end,0) Lowactivity_end FROM Channel_Config where Channel_Group='"&prm&"' "
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		Response.Write "{ ""cg_ftype"": """&Rs.Fields("File_Type")&""", ""cg_group"": """&Rs.Fields("Channel_Group")&""", ""cg_channel"": """&Rs.Fields("Channel")&""", ""cg_warning"": "&d(Rs.Fields("Warning_Count"))&", ""cg_error"":  "&d(Rs.Fields("Error_Count"))&", ""cg_minimal"":  "&Rs.Fields("Min_Count")&", ""cg_limit"": "&Rs.Fields("Limit_Count")&", ""cg_lowactivity_start"": "&Rs.Fields("Lowactivity_start")&", ""cg_lowactivity_end"": "&Rs.Fields("Lowactivity_end")&"  }"
	end if
	Rs.Close
end if

if ds="SaveChannelGroup" then
	cg_group=Request("cg_group")
	cg_warning=d(Request("cg_warning"))
	cg_error=d(Request("cg_error"))
	cg_minimal=Request("cg_minimal")

    cg_limit=Request("cg_limit")
    cg_lowactivity_start=Request("cg_lowactivity_start")
    cg_lowactivity_end=Request("cg_lowactivity_end")

	SQL_="UPDATE Channel_Config set Warning_Count="&cg_warning&", Error_Count="&cg_error&", Min_Count="&cg_minimal
    SQL_=SQL_&" , Limit_Count="&cg_limit&", Lowactivity_start="&cg_lowactivity_start&", Lowactivity_end="&cg_lowactivity_end&"  where Channel_Group='"&cg_group&"' "
	Cmd.CommandText=SQL_
	Cmd.Execute
end if
'-------END: Channel Groups--------------------------------------------------------------

if ds="GetUserProp" then
	SQL_="SELECT [User_ID],[User_Name],[User_Login],[Role],[Phone],[Email] FROM Users WHERE [User_ID]="&prm
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		ColCount=Rs.Fields.Count
		for i=0 to ColCount-1
			Response.Write(Rs.Fields(i)&"^")
		next
	end if
	Rs.Close
end if

if ds="GetUserHist" then
		SQL_="SELECT TOP 100 [DT],[User_Login],[Descr] FROM [UsersLog] WHERE [User_Login]=(SELECT User_Login FROM Users WHERE User_ID="&prm&") ORDER BY DT DESC"
		Rs.Open SQL_, Conn
		Response.Write("<table width=""900px"" cellpadding=""0"" cellspacing=""0"" style=""font-size: 10pt; table-layout: fixed"">")
		Response.Write("<colgroup><col width=""200px""><col width=""200px""><col width=""*""></colgroup>")
		Response.Write("<tbody><tr><th>Дата</th><th>Логин</th><th>Событие</th></tr></tbody></table>")
		Response.Write("<div style=""OVERFLOW-Y: auto; OVERFLOW-X: hidden; OVERFLOW: auto; width: 900px; height: 300px;"">")
		Response.Write("<table width=""900px"" cellpadding=""0"" cellspacing=""0"" style=""font-size: 10pt; table-layout: fixed"">")
		Response.Write("<colgroup><col width=""200px""><col width=""200px""><col width=""*""></colgroup>")
		Response.Write("<tbody>")
		if not Rs.Eof then
			ColCount=Rs.Fields.Count
			do while not Rs.Eof
				id=Rs.Fields(0)
				Response.Write("<tr>")
				for i=0 to ColCount-1
					Response.Write("<td>"&NBSP(Trim(Rs.Fields(i)))&"</td>")
				next
				Response.Write("</tr>"&vbCrLf)
				Rs.MoveNext
			loop
		else
			Response.Write("<tr><td colspan=""3"">Нет данных</td></tr>")
		end if
		Response.Write("</tbody></table></div>")
		Rs.Close
end if

if ds="GetFin" then
	SQL_="SELECT [Branch_code], [Name] FROM V_Branch_code WHERE [Branch_code]='"&prm&"'"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		ColCount=Rs.Fields.Count
		for i=0 to ColCount-1
			Response.Write(Rs.Fields(i)&"^")
		next
	end if
	Rs.Close
end if

if ds="GetRC" then
	SQL_="SELECT [Resp_code], [Resp_text], [IsFailed] FROM V_Resp_code WHERE [Resp_code]='"&prm&"'"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		ColCount=Rs.Fields.Count
		for i=0 to ColCount-1
			Response.Write(Rs.Fields(i)&"^")
		next
	end if
	Rs.Close
end if

if ds="GetTag" then
	SQL_="SELECT [TagID],[TagName],[SetHi],[SetHiHi],[FileID],[GroupName],[Prop_Crit],[Prop_Active],[Prop_SignOn],[Prop_Time] FROM [Tags] WHERE [TagID]='"&prm&"'"
	Rs.Open SQL_, Conn
	if not Rs.Eof then
		ColCount=Rs.Fields.Count
		for i=0 to ColCount-1
			Response.Write(Rs.Fields(i)&"^")
		next
	end if
	Rs.Close
end if

'----------------------------------------------------------------------------------------------------
'-------START: Channel Groups------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
if ds="ChannelISS" then
if prm= "Graph" then

	Series = ""
	Series_FAIL = ""
	RequestParam = ""
	if (prm2="ISS_VISA") then
		RequestParam = " SOURCE_CHANNEL='VISA' "
	elseif (prm2="ISS_MC") then
		RequestParam = " SOURCE_CHANNEL='MasterCard' "
	elseif (prm2="ISS_NSPK_VISA") then
		RequestParam = " SOURCE_CHANNEL='NSPK_VISA' "
	elseif (prm2="ISS_NSPK_MC") then
		RequestParam = " SOURCE_CHANNEL='NSPK_MasterCard' "
	elseif (prm2="ISS_MIR") then
		RequestParam = " SOURCE_CHANNEL='NSPK MIR' "	
	end if

	DBefore = Request("DBefore")
	if Request("DBefore")<>"" then
		DBefore = Request("DBefore")
	else
		DBefore = 0
	end if
		
	sqlstr = "set dateformat ymd; SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION,SUM(OPERATION_FAIL) OPERATION_FAIL, SOURCE_CHANNEL FROM LOG_VO "
	sqlstr = sqlstr&" WHERE [TIME]>=convert(datetime,floor(convert(float,DATEADD(DAY,"&DBefore&",GETDATE()) ))) "
	sqlstr = sqlstr&" and [TIME]<convert(datetime,floor(convert(float, DATEADD(DAY,"&DBefore&"+1,GETDATE()) ))) "
	sqlstr = sqlstr&" and "&RequestParam&" GROUP BY [TIME], SOURCE_CHANNEL order by [TIME]"
	Rs.Open sqlstr, Conn
	If not Rs.EOF then
	do while (not Rs.EOF)
	
	    v = Rs.Fields("OPERATION")-Rs.Fields("OPERATION_FAIL")
		v1 = Rs.Fields("OPERATION_FAIL")

		if (Series<>"") then 
			Series = Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			Series_FAIL = Series_FAIL&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v1&"}"
		else
			Series = Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			Series_FAIL = Series_FAIL&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v1&"}"
		end if 
		
		Rs.MoveNext
	loop
	end if
	Rs.Close
	
	Response.Write(Series&"~"&Series_FAIL)

elseif prm= "Table" then
	UpdateChannelGroupsTable ds, Request("DBefore")
end if
end if

if ds="ChannelACQ" then
if prm= "Graph" then

	Series = ""
	Series_FAIL = ""
	RequestParam = ""
	if (prm2="ACQ_VISA") then
		RequestParam = " (TARGET_CHANNEL='VISA' or TARGET_CHANNEL='VISA SMS') "
	elseif (prm2="ACQ_MC") then
		RequestParam = " TARGET_CHANNEL='MasterCard' "
	elseif (prm2="ACQ_NSPK_VISA") then
		RequestParam = " (TARGET_CHANNEL='NSPK_VISA' or TARGET_CHANNEL='NSPK_VISA SMS') "
	elseif (prm2="ACQ_NSPK_MC") then
		RequestParam = " TARGET_CHANNEL='NSPK_MasterCard' "
	elseif (prm2="ACQ_MIR") then
		RequestParam = " TARGET_CHANNEL='NSPK MIR' "	
	end if
		
	sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION,SUM(OPERATION_FAIL) OPERATION_FAIL FROM LOG_VO "
	sqlstr = sqlstr&" WHERE [TIME]>=convert(datetime,floor(convert(float,Getdate()))) and "&RequestParam&" GROUP BY [TIME] order by [TIME]"
	Rs.Open sqlstr, Conn
	If not Rs.EOF then
	do while (not Rs.EOF)
	
	    v = Rs.Fields("OPERATION")-Rs.Fields("OPERATION_FAIL")
		v1 = Rs.Fields("OPERATION_FAIL")

		if (Series<>"") then 
			Series = Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			Series_FAIL = Series_FAIL&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v1&"}"
		else
			Series = Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			Series_FAIL = Series_FAIL&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v1&"}"
		end if 
		
		Rs.MoveNext
	loop
	end if
	Rs.Close
	
	Response.Write(Series&"~"&Series_FAIL)
elseif prm= "Table" then
	UpdateChannelGroupsTable ds, Request("DBefore")
end if
end if

'----------------------------------------------------------------------------------------------------

if ds="ChannelATM" then
if prm= "Graph" then

	Series = ""
	Series_FAIL = ""
	RequestParam = ""
	if (prm2="All_ATM") then
		RequestParam = " SOURCE_CHANNEL='OUR_ATM' "
	elseif (prm2="All_BPT") then
		RequestParam = " SOURCE_CHANNEL='OUR_BPT' "
	elseif (prm2="All_POS") then
		RequestParam = " SOURCE_CHANNEL='OUR_POS' "
	elseif (prm2="All_H2H_RBS") then
		RequestParam = " SOURCE_CHANNEL='H2H_BPCRBS' "
	end if
		
	sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION,SUM(OPERATION_FAIL) OPERATION_FAIL, SOURCE_CHANNEL FROM LOG_VO WHERE [TIME]>=convert(datetime,floor(convert(float,Getdate()))) and "&RequestParam&" GROUP BY [TIME], SOURCE_CHANNEL order by [TIME]"
	Rs.Open sqlstr, Conn
	If not Rs.EOF then
	do while (not Rs.EOF)
	
	    v = Rs.Fields("OPERATION")-Rs.Fields("OPERATION_FAIL")
		v1 = Rs.Fields("OPERATION_FAIL")

		if (Series<>"") then 
			Series = Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			Series_FAIL = Series_FAIL&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v1&"}"
		else
			Series = Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			Series_FAIL = Series_FAIL&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v1&"}"
		end if 
		
		Rs.MoveNext
	loop
	end if
	Rs.Close
	
	Response.Write(Series&"~"&Series_FAIL)
elseif prm= "Table" then
	UpdateChannelGroupsTable ds, Request("DBefore")
end if
end if

'----------------------------------------------------------------------------------------------------

if ds="Channel3DS" then
if prm= "Graph" then
	Series = ""
	Series_FAIL = ""
	RequestParam = ""
	if (prm2="NSPK_VISA") then
		RequestParam = " SOURCE_CHANNEL='NSPK_VISA' and SERVICE='3D-Secure'  "
	elseif (prm2="NSPK_MC") then
		RequestParam = " SOURCE_CHANNEL='NSPK_MasterCard' and SERVICE='3D-Secure'  "
	elseif (prm2="VISA") then
		RequestParam = " SOURCE_CHANNEL='VISA' and SERVICE='3D-Secure' "
	elseif (prm2="MC") then
		RequestParam = " SOURCE_CHANNEL='MasterCard' and SERVICE='3D-Secure' "
	elseif (prm2="SOA_USB") then
		RequestParam = " SOURCE_CHANNEL='RBS' and SERVICE='SOA_USB'  "
	elseif (prm2="SOA_AGENT") then
		RequestParam = " SOURCE_CHANNEL='OUR_POS' and SERVICE='SOA_AGENT' "
	end if
		
	sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION,SUM(OPERATION_FAIL) OPERATION_FAIL, SOURCE_CHANNEL FROM LOG_VS WHERE [TIME]>=convert(datetime,floor(convert(float,Getdate()))) and "&RequestParam&" GROUP BY [TIME], SOURCE_CHANNEL order by [TIME]"
	Rs.Open sqlstr, Conn
	If not Rs.EOF then
	do while (not Rs.EOF)
	
	    v = Rs.Fields("OPERATION")-Rs.Fields("OPERATION_FAIL")
		v1 = Rs.Fields("OPERATION_FAIL")

		if (Series<>"") then 
			Series = Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			Series_FAIL = Series_FAIL&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v1&"}"
		else
			Series = Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			Series_FAIL = Series_FAIL&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v1&"}"
		end if 
		
		Rs.MoveNext
	loop
	end if
	Rs.Close
	
	Response.Write(Series&"~"&Series_FAIL)
elseif prm= "Table" then
	UpdateChannelGroupsTable ds, Request("DBefore")
end if
end if
'----------------------------------------------------------------------------------------------------
'-------END: Channel Groups--------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------

dim warnings(20,7)

if ((ds="ChannelISS") or  (ds="ChannelACQ") or  (ds="ChannelATM")  or  (ds="Channel3DS")) then
		' Заполнение Warings
		i = 0
		sqlstr = "SELECT Channel_Group, ISNULL(Warning_Count,0) Warning_Count, ISNULL(Error_Count,0) Error_Count, ISNULL(Min_Count,0) Min_Count, ISNULL(Limit_Count,0) Limit_Count,  ISNULL(Lowactivity_start,0) Lowactivity_start, ISNULL(Lowactivity_end,0) Lowactivity_end  FROM  Channel_Config"
		Rs.OPEN sqlstr, CONN
		if not RS.EOF then
		do while (not RS.EOF)
			warnings(i,0)=Rs.Fields("Channel_Group") ' Channel_Group
			warnings(i,1)=Rs.Fields("Warning_Count")	' Warning_Count
			warnings(i,2)=Rs.Fields("Error_Count")	' Error_Count
			warnings(i,3)=Rs.Fields("Min_Count")	' Min_Count
					warnings(i,4)=Rs.Fields("Limit_Count")	' Limit_Count
					warnings(i,5)=Rs.Fields("Lowactivity_start")	' Lowactivity_start
					warnings(i,6)=Rs.Fields("Lowactivity_end")	' Lowactivity_end
			i = i+1
			Rs.MoveNext
		loop
		end if
		RS.CLOSE

end if

		function checkWarning(paramName, failCount, totalCount, minutes_val)
			res = "" 'clWarning clError
			if (totalCount>0) then
				failCount = (failCount*100)/totalCount ' проверяем процент сбойных
				for j=0 to UBound(warnings)
					if (warnings(j,0)=paramName) then
						if ((warnings(j,2)>0)and(totalCount>warnings(j,3))and(failCount>warnings(j,2))) then
							res = clError
						elseif ((warnings(j,1)>0)and(totalCount>warnings(j,3))and(failCount>warnings(j,1))) then
							res = clWarning
						end if

										'if (warnings(j,6)>0) then
											'  if (((minutes_val<warnings(j,5))or(minutes_val>warnings(j,6)))and(totalCount<warnings(j,4))) then
											'      res = clError
											'  end if
										'end if

						checkWarning = res
					end if
				next
			end if
			checkWarning = res
		end function

		function checkWarning_all(paramName, failCount, totalCount, minutes_val)
			res = "" 'clWarning clError
			if (totalCount>0) then
				'failCount = (failCount*100)/totalCount ' проверяем процент сбойных
				for j=0 to UBound(warnings)
					if (warnings(j,0)=paramName) then

										if (warnings(j,6)>0) then
												if (((minutes_val<warnings(j,5))or(minutes_val>warnings(j,6)))and(totalCount<warnings(j,4))) then
														res = clError
												end if
										end if

						checkWarning_all = res
					end if
				next
			end if
			checkWarning_all = res
		end function


'----------------------------------------------------------------------------------------------------
'-------Start: Channel Groups Table------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
function UpdateChannelGroupsTable(ChName,DOfset)
		DBefore = 0
		if DOfset<>"" then
			DBefore = DOfset
		else
			DBefore = 0
		end if

		if ((ChName="ChannelISS") or  (ChName="ChannelACQ")) then
'------------T=7----------------------------------------------------------------------------------------
				ISS_VISA = 0
				ISS_VISA_FAIL = 0
				ISS_VISA_FAIL_PC = 0
				ACQ_VISA = 0 
				ACQ_VISA_FAIL = 0
				ACQ_VISA_FAIL_PC = 0
				ISS_MC = 0
				ISS_MC_FAIL = 0
				ISS_MC_FAIL_PC = 0
				ACQ_MC = 0
				ACQ_MC_FAIL = 0
				ACQ_MC_FAIL_PC = 0
				ISS_NSPK_VISA = 0
				ISS_NSPK_VISA_FAIL = 0
				ISS_NSPK_VISA_FAIL_PC = 0
				ACQ_NSPK_VISA = 0
				ACQ_NSPK_VISA_FAIL = 0
				ACQ_NSPK_VISA_FAIL_PC = 0
				ISS_NSPK_MC = 0
				ISS_NSPK_MC_FAIL = 0
				ISS_NSPK_MC_FAIL_PC = 0
				ACQ_NSPK_MC = 0
				ACQ_NSPK_MC_FAIL = 0
				ACQ_NSPK_MC_FAIL_PC = 0
				ISS_MIR = 0
				ISS_MIR_FAIL = 0
				ISS_MIR_FAIL_PC = 0
				ACQ_MIR = 0
				ACQ_MIR_FAIL = 0
				ACQ_MIR_FAIL_PC = 0
				
				ISS_VISA_Color = ""
				ACQ_VISA_Color = "" 
				ISS_MC_Color = ""
				ACQ_MC_Color = ""
				ISS_NSPK_VISA_Color = ""
				ACQ_NSPK_VISA_Color = ""
				ISS_NSPK_MC_Color = ""
				ACQ_NSPK_MC_Color = ""
				ISS_MIR_Color = ""
				ACQ_MIR_Color = ""

				ISS_VISA_Color_all = ""
				ACQ_VISA_Color_all = "" 
				ISS_MC_Color_all = ""
				ACQ_MC_Color_all = ""
				ISS_NSPK_VISA_Color_all = ""
				ACQ_NSPK_VISA_Color_all = ""
				ISS_NSPK_MC_Color_all = ""
				ACQ_NSPK_MC_Color_all = ""
				ISS_MIR_Color_all = ""
				ACQ_MIR_Color_all = ""
				
				sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION,SUM(OPERATION_FAIL) OPERATION_FAIL, SOURCE_CHANNEL  "
					sqlstr = sqlstr&" ,DATEPART(HOUR,[TIME])*60+DATEPART(MINUTE,[TIME]) timeinminutes FROM LOG_VO "
				'sqlstr = sqlstr&" WHERE [TIME]=(select top 1 [TIME] from LOG_VO order by [TIME] desc) "
				sqlstr = sqlstr&" WHERE [TIME]>=convert(datetime,floor(convert(float,DATEADD(DAY,"&DBefore&",GETDATE()) ))) "
				sqlstr = sqlstr&" and [TIME]<convert(datetime,floor(convert(float, DATEADD(DAY,"&DBefore&"+1,GETDATE()) ))) "
				sqlstr = sqlstr&" GROUP BY [TIME],SOURCE_CHANNEL"
				RS.OPEN sqlstr, CONN
				IF NOT RS.EOF THEN
				DO WHILE (NOT RS.EOF)
					if (Rs.Fields("SOURCE_CHANNEL")="VISA") then
						ISS_VISA = Rs.Fields("OPERATION")
						ISS_VISA_Color = checkWarning("VISA_ISS", Rs.Fields("OPERATION_FAIL"), ISS_VISA, Rs.Fields("timeinminutes"))
									ISS_VISA_Color_all = checkWarning_all("VISA_ISS", Rs.Fields("OPERATION_FAIL"), ISS_VISA, Rs.Fields("timeinminutes"))
						ISS_VISA_FAIL = Rs.Fields("OPERATION_FAIL")
						if (ISS_VISA>0) then
							ISS_VISA_FAIL_PC=(ISS_VISA_FAIL*100)/ISS_VISA
						end if
					elseif (Rs.Fields("SOURCE_CHANNEL")="MasterCard") then
						ISS_MC = Rs.Fields("OPERATION")
						ISS_MC_Color = checkWarning("MC_ISS", Rs.Fields("OPERATION_FAIL"), ISS_MC, Rs.Fields("timeinminutes"))
									ISS_MC_Color_all = checkWarning_all("MC_ISS", Rs.Fields("OPERATION_FAIL"), ISS_MC, Rs.Fields("timeinminutes"))
						ISS_MC_FAIL = Rs.Fields("OPERATION_FAIL")
						if (ISS_MC>0) then
							ISS_MC_FAIL_PC=(ISS_MC_FAIL*100)/ISS_MC
						end if
					elseif (Rs.Fields("SOURCE_CHANNEL")="NSPK_VISA") then
						ISS_NSPK_VISA = Rs.Fields("OPERATION")
						ISS_NSPK_VISA_Color = checkWarning("NSPK_VISA_ISS", Rs.Fields("OPERATION_FAIL"), ISS_NSPK_VISA, Rs.Fields("timeinminutes"))
									ISS_NSPK_VISA_Color_all = checkWarning_all("NSPK_VISA_ISS", Rs.Fields("OPERATION_FAIL"), ISS_NSPK_VISA, Rs.Fields("timeinminutes"))
						ISS_NSPK_VISA_FAIL = Rs.Fields("OPERATION_FAIL")
						if (ISS_NSPK_VISA>0) then
							ISS_NSPK_VISA_FAIL_PC=(ISS_NSPK_VISA_FAIL*100)/ISS_NSPK_VISA
						end if
					elseif (Rs.Fields("SOURCE_CHANNEL")="NSPK_MasterCard") then
						ISS_NSPK_MC = Rs.Fields("OPERATION")
						ISS_NSPK_MC_Color = checkWarning("NSPK_MC_ISS", Rs.Fields("OPERATION_FAIL"), ISS_NSPK_MC, Rs.Fields("timeinminutes"))
									ISS_NSPK_MC_Color_all = checkWarning_all("NSPK_MC_ISS", Rs.Fields("OPERATION_FAIL"), ISS_NSPK_MC, Rs.Fields("timeinminutes"))
						ISS_NSPK_MC_FAIL = Rs.Fields("OPERATION_FAIL")
						if (ISS_NSPK_MC>0) then
							ISS_NSPK_MC_FAIL_PC=(ISS_NSPK_MC_FAIL*100)/ISS_NSPK_MC
						end if
					elseif (Rs.Fields("SOURCE_CHANNEL")="NSPK MIR") then
						ISS_MIR = Rs.Fields("OPERATION")
						ISS_MIR_Color = checkWarning("MIR_ISS", Rs.Fields("OPERATION_FAIL"), ISS_MIR, Rs.Fields("timeinminutes"))
									ISS_MIR_Color_all = checkWarning_all("MIR_ISS", Rs.Fields("OPERATION_FAIL"), ISS_MIR, Rs.Fields("timeinminutes"))
						ISS_MIR_FAIL = Rs.Fields("OPERATION_FAIL")
						if (ISS_MIR>0) then
							ISS_MIR_FAIL_PC=(ISS_MIR_FAIL*100)/ISS_MIR
						end if
					end if
					
					Rs.MoveNext
				LOOP
				END IF
				RS.CLOSE
				
				sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION,SUM(OPERATION_FAIL) OPERATION_FAIL, case when TARGET_CHANNEL='NSPK_VISA SMS' then 'NSPK_VISA'  when TARGET_CHANNEL='VISA SMS' then 'VISA'  else TARGET_CHANNEL end as TARGET_CHANNEL "
				sqlstr = sqlstr&" ,DATEPART(HOUR,[TIME])*60+DATEPART(MINUTE,[TIME]) timeinminutes FROM LOG_VO "
				'sqlstr = sqlstr&" WHERE [TIME]=(select top 1 [TIME] from LOG_VO order by [TIME] desc) "
				sqlstr = sqlstr&" WHERE [TIME]>=convert(datetime,floor(convert(float,DATEADD(DAY,"&DBefore&",GETDATE()) ))) "
				sqlstr = sqlstr&" and [TIME]<convert(datetime,floor(convert(float, DATEADD(DAY,"&DBefore&"+1,GETDATE()) ))) "
				sqlstr = sqlstr&" GROUP BY [TIME], case when TARGET_CHANNEL='NSPK_VISA SMS' then 'NSPK_VISA'  when TARGET_CHANNEL='VISA SMS' then 'VISA'  else TARGET_CHANNEL end "
				Rs.OPEN sqlstr, CONN
				If not Rs.EOF then
				do while (not Rs.EOF)
					if ((Rs.Fields("TARGET_CHANNEL")="VISA")or(Rs.Fields("TARGET_CHANNEL")="VISA SMS")) then
						ACQ_VISA = Rs.Fields("OPERATION")
						ACQ_VISA_Color = checkWarning("VISA_ACQ", Rs.Fields("OPERATION_FAIL"), ACQ_VISA, Rs.Fields("timeinminutes"))
									ACQ_VISA_Color_all = checkWarning_all("VISA_ACQ", Rs.Fields("OPERATION_FAIL"), ACQ_VISA, Rs.Fields("timeinminutes"))
						ACQ_VISA_FAIL = Rs.Fields("OPERATION_FAIL")
						if (ACQ_VISA>0) then
							ACQ_VISA_FAIL_PC=(ACQ_VISA_FAIL*100)/ACQ_VISA
						end if
					elseif (Rs.Fields("TARGET_CHANNEL")="MasterCard") then
						ACQ_MC = Rs.Fields("OPERATION")
						ACQ_MC_Color = checkWarning("MC_ACQ", Rs.Fields("OPERATION_FAIL"), ACQ_MC, Rs.Fields("timeinminutes"))
						ACQ_MC_Color_all = checkWarning_all("MC_ACQ", Rs.Fields("OPERATION_FAIL"), ACQ_MC, Rs.Fields("timeinminutes"))
						ACQ_MC_FAIL = Rs.Fields("OPERATION_FAIL")
						if (ACQ_MC>0) then
							ACQ_MC_FAIL_PC=(ACQ_MC_FAIL*100)/ACQ_MC
						end if
					elseif ((Rs.Fields("TARGET_CHANNEL")="NSPK_VISA")or(Rs.Fields("TARGET_CHANNEL")="NSPK_VISA SMS")) then
						ACQ_NSPK_VISA = Rs.Fields("OPERATION")
						ACQ_NSPK_VISA_Color = checkWarning("NSPK_VISA_ACQ", Rs.Fields("OPERATION_FAIL"), ACQ_NSPK_VISA, Rs.Fields("timeinminutes"))
						ACQ_NSPK_VISA_Color_all = checkWarning_all("NSPK_VISA_ACQ", Rs.Fields("OPERATION_FAIL"), ACQ_NSPK_VISA, Rs.Fields("timeinminutes"))
						ACQ_NSPK_VISA_FAIL = Rs.Fields("OPERATION_FAIL")
						if (ACQ_NSPK_VISA>0) then
							ACQ_NSPK_VISA_FAIL_PC=(ACQ_NSPK_VISA_FAIL*100)/ACQ_NSPK_VISA
						end if
					elseif (Rs.Fields("TARGET_CHANNEL")="NSPK_MasterCard") then
						ACQ_NSPK_MC = Rs.Fields("OPERATION")
						ACQ_NSPK_MC_Color = checkWarning("NSPK_MC_ACQ", Rs.Fields("OPERATION_FAIL"), ACQ_NSPK_MC, Rs.Fields("timeinminutes"))
						ACQ_NSPK_MC_Color_all = checkWarning_all("NSPK_MC_ACQ", Rs.Fields("OPERATION_FAIL"), ACQ_NSPK_MC, Rs.Fields("timeinminutes"))
						ACQ_NSPK_MC_FAIL = Rs.Fields("OPERATION_FAIL")
						if (ACQ_NSPK_MC>0) then
							ACQ_NSPK_MC_FAIL_PC=(ACQ_NSPK_MC_FAIL*100)/ACQ_NSPK_MC
						end if
					elseif (Rs.Fields("TARGET_CHANNEL")="NSPK MIR") then
						ACQ_MIR = Rs.Fields("OPERATION")
						ACQ_MIR_Color = checkWarning("MIR_ACQ", Rs.Fields("OPERATION_FAIL"), ACQ_MIR, Rs.Fields("timeinminutes"))
						ACQ_MIR_Color_all = checkWarning_all("MIR_ACQ", Rs.Fields("OPERATION_FAIL"), ACQ_MIR, Rs.Fields("timeinminutes"))
						ACQ_MIR_FAIL = Rs.Fields("OPERATION_FAIL")
						if (ACQ_MIR>0) then
							ACQ_MIR_FAIL_PC=(ACQ_MIR_FAIL*100)/ACQ_MIR
						end if
					end if

					Rs.MoveNext
				LOOP
				END IF
				RS.CLOSE

				if (ISS_VISA_Color<>"") then 
					if (ISS_VISA_Color = clWarning) then
						ISS_VISA_Color = " style=""color: #000000; background: "&ISS_VISA_Color&" "" "
					else
						ISS_VISA_Color = " style=""background: "&ISS_VISA_Color&" "" "
					end if
				end if 
				if (ACQ_VISA_Color<>"") then
					if (ACQ_VISA_Color = clWarning) then
						ACQ_VISA_Color = " style=""color: #000000; background: "&ACQ_VISA_Color&" "" "
					else
						ACQ_VISA_Color = " style=""background: "&ACQ_VISA_Color&" "" "
					end if
				end if 
				if (ISS_MC_Color<>"") then
					if (ISS_MC_Color = clWarning) then
						ISS_MC_Color = " style=""color: #000000; background: "&ISS_MC_Color&" "" "
					else
						ISS_MC_Color = " style=""background: "&ISS_MC_Color&" "" "
					end if
				end if 
				if (ACQ_MC_Color<>"") then 
					if (ACQ_MC_Color = clWarning) then
						ACQ_MC_Color = " style=""color: #000000; background: "&ACQ_MC_Color&" "" "
					else
						ACQ_MC_Color = " style=""background: "&ACQ_MC_Color&" "" "
					end if
				end if 
				if (ISS_NSPK_VISA_Color<>"") then 
					if (ISS_NSPK_VISA_Color = clWarning) then
						ISS_NSPK_VISA_Color = " style=""color: #000000; background: "&ISS_NSPK_VISA_Color&" "" "
					else
						ISS_NSPK_VISA_Color = " style=""background: "&ISS_NSPK_VISA_Color&" "" "
					end if
				end if 
				if (ACQ_NSPK_VISA_Color<>"") then 
					if (ACQ_NSPK_VISA_Color = clWarning) then
						ACQ_NSPK_VISA_Color = " style=""color: #000000; background: "&ACQ_NSPK_VISA_Color&" "" "
					else
						ACQ_NSPK_VISA_Color = " style=""background: "&ACQ_NSPK_VISA_Color&" "" "
					end if
				end if 
				if (ISS_NSPK_MC_Color<>"") then
					if (ISS_NSPK_MC_Color = clWarning) then
						ISS_NSPK_MC_Color = " style=""color: #000000; background: "&ISS_NSPK_MC_Color&" "" "
					else
						ISS_NSPK_MC_Color = " style=""background: "&ISS_NSPK_MC_Color&" "" "
					end if
				end if 
				if (ACQ_NSPK_MC_Color<>"") then 
					if (ACQ_NSPK_MC_Color = clWarning) then
						ACQ_NSPK_MC_Color = " style=""color: #000000; background: "&ACQ_NSPK_MC_Color&" "" "
					else
						ACQ_NSPK_MC_Color = " style=""background: "&ACQ_NSPK_MC_Color&" "" "
					end if
				end if
				if (ISS_MIR_Color<>"") then 
					if (ISS_MIR_Color = clWarning) then
						ISS_MIR_Color = " style=""color: #000000; background: "&ISS_MIR_Color&" "" "
					else
						ISS_MIR_Color = " style=""background: "&ISS_MIR_Color&" "" "
					end if
				end if
				if (ACQ_MIR_Color<>"") then 
					if (ACQ_MIR_Color = clWarning) then
						ACQ_MIR_Color = " style=""color: #000000; background: "&ACQ_MIR_Color&" "" "
					else
						ACQ_MIR_Color = " style=""background: "&ACQ_MIR_Color&" "" "
					end if
				end if	
					'-----------------
					if (ISS_VISA_Color_all<>"") then 
					if (ISS_VISA_Color_all = clWarning) then
						ISS_VISA_Color_all = " style=""color: #000000; background: "&ISS_VISA_Color_all&" "" "
					else
						ISS_VISA_Color_all = " style=""background: "&ISS_VISA_Color_all&" "" "
					end if
				end if 
				if (ACQ_VISA_Color_all<>"") then
					if (ACQ_VISA_Color_all = clWarning) then
						ACQ_VISA_Color_all = " style=""color: #000000; background: "&ACQ_VISA_Color_all&" "" "
					else
						ACQ_VISA_Color_all = " style=""background: "&ACQ_VISA_Color_all&" "" "
					end if
				end if 
				if (ISS_MC_Color_all<>"") then
					if (ISS_MC_Color_all = clWarning) then
						ISS_MC_Color_all = " style=""color: #000000; background: "&ISS_MC_Color_all&" "" "
					else
						ISS_MC_Color_all = " style=""background: "&ISS_MC_Color_all&" "" "
					end if
				end if 
				if (ACQ_MC_Color_all<>"") then 
					if (ACQ_MC_Color_all = clWarning) then
						ACQ_MC_Color_all = " style=""color: #000000; background: "&ACQ_MC_Color_all&" "" "
					else
						ACQ_MC_Color_all = " style=""background: "&ACQ_MC_Color_all&" "" "
					end if
				end if 
				if (ISS_NSPK_VISA_Color_all<>"") then 
					if (ISS_NSPK_VISA_Color_all = clWarning) then
						ISS_NSPK_VISA_Color_all = " style=""color: #000000; background: "&ISS_NSPK_VISA_Color_all&" "" "
					else
						ISS_NSPK_VISA_Color_all = " style=""background: "&ISS_NSPK_VISA_Color_all&" "" "
					end if
				end if 
				if (ACQ_NSPK_VISA_Color_all<>"") then 
					if (ACQ_NSPK_VISA_Color_all = clWarning) then
						ACQ_NSPK_VISA_Color_all = " style=""color: #000000; background: "&ACQ_NSPK_VISA_Color_all&" "" "
					else
						ACQ_NSPK_VISA_Color_all = " style=""background: "&ACQ_NSPK_VISA_Color_all&" "" "
					end if
				end if 
				if (ISS_NSPK_MC_Color_all<>"") then
					if (ISS_NSPK_MC_Color_all = clWarning) then
						ISS_NSPK_MC_Color_all = " style=""color: #000000; background: "&ISS_NSPK_MC_Color_all&" "" "
					else
						ISS_NSPK_MC_Color_all = " style=""background: "&ISS_NSPK_MC_Color_all&" "" "
					end if
				end if 
				if (ACQ_NSPK_MC_Color_all<>"") then 
					if (ACQ_NSPK_MC_Color_all = clWarning) then
						ACQ_NSPK_MC_Color_all = " style=""color: #000000; background: "&ACQ_NSPK_MC_Color_all&" "" "
					else
						ACQ_NSPK_MC_Color_all = " style=""background: "&ACQ_NSPK_MC_Color_all&" "" "
					end if
				end if
				if (ISS_MIR_Color_all<>"") then 
					if (ISS_MIR_Color_all = clWarning) then
						ISS_MIR_Color_all = " style=""color: #000000; background: "&ISS_MIR_Color_all&" "" "
					else
						ISS_MIR_Color_all = " style=""background: "&ISS_MIR_Color_all&" "" "
					end if
				end if
				if (ACQ_MIR_Color_all<>"") then 
					if (ACQ_MIR_Color_all = clWarning) then
						ACQ_MIR_Color_all = " style=""color: #000000; background: "&ACQ_MIR_Color_all&" "" "
					else
						ACQ_MIR_Color_all = " style=""background: "&ACQ_MIR_Color_all&" "" "
					end if
				end if

				tblISS_ACQ = ""

					'-----------SORTING-------------------------------------------------------------------------------------
			'1-ая (верхняя) строка - группа каналов NSPK_VISA (по умолчанию строить график для группы каналов NSPK_VISA_ACQ)
			'2-ая строка - группа каналов NSPK_MC
			'3-ья строка – группа каналов MIR
			'4-ая строка – группа каналов VISA
			'5-ая строка – группа каналов MC

					Dim tbl_tr(6,6)
					Dim str_tr(6)

					tbl_tr(1, 1)=Max(ISS_NSPK_VISA_FAIL_PC, ACQ_NSPK_VISA_FAIL_PC)
					tbl_tr(1, 2)=ISS_NSPK_VISA_FAIL_PC
					tbl_tr(1, 3)=ACQ_NSPK_VISA_FAIL_PC
					tbl_tr(1, 4)="ISS_NSPK_VISA"
					tbl_tr(1, 5)="ACQ_NSPK_VISA"
					tbl_tr(1, 6)=1

					tbl_tr(2, 1)=Max(ISS_NSPK_MC_FAIL_PC, ACQ_NSPK_MC_FAIL_PC)
					tbl_tr(2, 2)=ISS_NSPK_MC_FAIL_PC
					tbl_tr(2, 3)=ACQ_NSPK_MC_FAIL_PC
					tbl_tr(2, 4)="ISS_NSPK_MC"
					tbl_tr(2, 5)="ACQ_NSPK_MC"
					tbl_tr(2, 6)=2

					tbl_tr(3, 1)=Max(ISS_MIR_FAIL_PC, ACQ_MIR_FAIL_PC)
					tbl_tr(3, 2)=ISS_MIR_FAIL_PC
					tbl_tr(3, 3)=ACQ_MIR_FAIL_PC
					tbl_tr(3, 4)="ISS_MIR"
					tbl_tr(3, 5)="ACQ_MIR"
					tbl_tr(3, 6)=3

					tbl_tr(4, 1)=Max(ISS_VISA_FAIL_PC, ACQ_VISA_FAIL_PC)
					tbl_tr(4, 2)=ISS_VISA_FAIL_PC
					tbl_tr(4, 3)=ACQ_VISA_FAIL_PC
					tbl_tr(4, 4)="ISS_VISA"
					tbl_tr(4, 5)="ACQ_VISA"
					tbl_tr(4, 6)=4

					tbl_tr(5, 1)=Max(ISS_MC_FAIL_PC, ACQ_MC_FAIL_PC)
					tbl_tr(5, 2)=ISS_MC_FAIL_PC
					tbl_tr(5, 3)=ACQ_MC_FAIL_PC
					tbl_tr(5, 4)="ISS_MC"
					tbl_tr(5, 5)="ACQ_MC"
					tbl_tr(5, 6)=5


					For j = 1 To 5-1
							For k = j + 1 To 5
									If (tbl_tr(j,1) < tbl_tr(k,1)) or ((tbl_tr(k,1)=0) and (tbl_tr(j,1)=0) and (tbl_tr(j,6) > tbl_tr(k,6)) ) Then
											For l = 1 To 6
													Temp = tbl_tr(j,l)
													tbl_tr(j,l) = tbl_tr(k,l)
													tbl_tr(k,l) = Temp
											Next

									End If
							Next
					Next



			'tr_NSPK_VISA 
				str_tr(1)  = "<tr><td width=""100px"" >NSPK_VISA</td><td onclick=""ChGraph('ISS_NSPK_VISA',daysBefore)"" width=""70px""  "&ISS_NSPK_VISA_Color_all&" >"&ISS_NSPK_VISA&"</td><td onclick=""ChGraph('ISS_NSPK_VISA',daysBefore)""  width=""70px"" >"&ISS_NSPK_VISA_FAIL&"</td><td onclick=""ChGraph('ISS_NSPK_VISA',daysBefore)""  width=""70px"" "&ISS_NSPK_VISA_Color&" >"&Round(ISS_NSPK_VISA_FAIL_PC,3)&"</td>"&_
							"<td onclick=""ChGraph('ACQ_NSPK_VISA',daysBefore)""  width=""70px"" "&ACQ_NSPK_VISA_Color_all&" >"&ACQ_NSPK_VISA&"</td><td onclick=""ChGraph('ACQ_NSPK_VISA',daysBefore)""  width=""70px"" >"&ACQ_NSPK_VISA_FAIL&"</td><td onclick=""ChGraph('ACQ_NSPK_VISA',daysBefore)""  width=""70px"" "&ACQ_NSPK_VISA_Color&" >"&Round(ACQ_NSPK_VISA_FAIL_PC,3)&"</td><tr>"
			'tr_NSPK_MC 
				str_tr(2) = "<tr><td width=""100px"" >NSPK_MC</td><td onclick=""ChGraph('ISS_NSPK_MC',daysBefore)"" width=""70px"" "&ISS_NSPK_MC_Color_all&" >"&ISS_NSPK_MC&"</td><td onclick=""ChGraph('ISS_NSPK_MC',daysBefore)"" width=""70px""  >"&ISS_NSPK_MC_FAIL&"</td><td onclick=""ChGraph('ISS_NSPK_MC',daysBefore)"" width=""70px""  "&ISS_NSPK_MC_Color&" >"&Round(ISS_NSPK_MC_FAIL_PC,3)&"</td>"&_
						"<td onclick=""ChGraph('ACQ_NSPK_MC',daysBefore)""  width=""70px"" "&ACQ_NSPK_MC_Color_all&" >"&ACQ_NSPK_MC&"</td><td onclick=""ChGraph('ACQ_NSPK_MC',daysBefore)""  width=""70px"" >"&ACQ_NSPK_MC_FAIL&"</td><td onclick=""ChGraph('ACQ_NSPK_MC',daysBefore)"" width=""70px""  "&ACQ_NSPK_MC_Color&" >"&Round(ACQ_NSPK_MC_FAIL_PC,3)&"</td><tr>"	
			'tr_MIR 
				str_tr(3) = "<tr><td width=""100px"" >MIR</td><td onclick=""ChGraph('ISS_MIR',daysBefore)""  width=""70px"" "&ISS_MIR_Color_all&" >"&ISS_MIR&"</td><td onclick=""ChGraph('ISS_MIR',daysBefore)""  width=""70px"" >"&ISS_MIR_FAIL&"</td><td onclick=""ChGraph('ISS_MIR',daysBefore)"" width=""70px""  "&ISS_MIR_Color&" >"&Round(ISS_MIR_FAIL_PC,3)&"</td>"&_
					"<td onclick=""ChGraph('ACQ_MIR',daysBefore)""  width=""70px"" "&ACQ_MIR_Color_all&" >"&ACQ_MIR&"</td><td onclick=""ChGraph('ACQ_MIR',daysBefore)""  width=""70px""  >"&ACQ_MIR_FAIL&"</td><td onclick=""ChGraph('ACQ_MIR',daysBefore)"" width=""70px""  "&ACQ_MIR_Color&" >"&Round(ACQ_MIR_FAIL_PC,3)&"</td><tr>"
			'tr_VISA
				str_tr(4) = "<tr><td width=""100px"" >VISA</td><td onclick=""ChGraph('ISS_VISA',daysBefore)"" width=""70px"" "&ISS_VISA_Color_all&" >"&ISS_VISA&"</td><td onclick=""ChGraph('ISS_VISA',daysBefore)"" width=""70px"" >"&ISS_VISA_FAIL&"</td><td onclick=""ChGraph('ISS_VISA',daysBefore)"" width=""70px"" "&ISS_VISA_Color&" >"&Round(ISS_VISA_FAIL_PC,3)&"</td>"&_
						"<td onclick=""ChGraph('ACQ_VISA',daysBefore)"" width=""70px"" "&ACQ_VISA_Color_all&" >"&ACQ_VISA&"</td><td onclick=""ChGraph('ACQ_VISA',daysBefore)"" width=""70px"" >"&ACQ_VISA_FAIL&"</td><td onclick=""ChGraph('ACQ_VISA',daysBefore)"" width=""70px"" "&ACQ_VISA_Color&" >"&Round(ACQ_VISA_FAIL_PC,3)&"</td><tr>"
			'tr_MC
				str_tr(5) = "<tr><td width=""100px"" >MC</td><td onclick=""ChGraph('ISS_MC',daysBefore)""  width=""70px"" "&ISS_MC_Color_all&" >"&ISS_MC&"</td><td onclick=""ChGraph('ISS_MC',daysBefore)""  width=""70px"" >"&ISS_MC_FAIL&"</td><td onclick=""ChGraph('ISS_MC',daysBefore)"" width=""70px""  "&ISS_MC_Color&" >"&Round(ISS_MC_FAIL_PC,3)&"</td>"&_
					"<td onclick=""ChGraph('ACQ_MC',daysBefore)""  width=""70px"" "&ACQ_MC_Color_all&" >"&ACQ_MC&"</td><td onclick=""ChGraph('ACQ_MC',daysBefore)""  width=""70px"" >"&ACQ_MC_FAIL&"</td><td onclick=""ChGraph('ACQ_MC',daysBefore)"" width=""70px""  "&ACQ_MC_Color&" >"&Round(ACQ_MC_FAIL_PC,3)&"</td><tr>"

				For j = 1 To 5
							tblISS_ACQ = tblISS_ACQ&str_tr(tbl_tr(j,6))
				Next
										

				Response.Write tblISS_ACQ

		elseif (ChName="ChannelATM") then
'------------T=8----------------------------------------------------------------------------------------
				All_ATM = 0
				All_BPT = 0
				All_POS = 0
				All_H2H_RBS = 0
				All_ATM_FAIL = 0
				All_BPT_FAIL = 0
				All_POS_FAIL = 0
				All_H2H_RBS_FAIL = 0
				All_ATM_FAIL_PC = 0
				All_BPT_FAIL_PC = 0
				All_POS_FAIL_PC = 0
				All_H2H_RBS_FAIL_PC = 0
				
				All_ATM_Color = ""
				All_BPT_Color = ""
				All_POS_Color = ""
				All_H2H_RBS_Color = ""

					All_ATM_Color_all = ""
				All_BPT_Color_all = ""
				All_POS_Color_all = ""
				All_H2H_RBS_Color_all = ""

				sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION,SUM(OPERATION_FAIL) OPERATION_FAIL, SOURCE_CHANNEL "
				sqlstr = sqlstr&" ,DATEPART(HOUR,[TIME])*60+DATEPART(MINUTE,[TIME]) timeinminutes FROM LOG_VO "
				'sqlstr = sqlstr&" WHERE [TIME]=(select top 1 [TIME] from LOG_VO order by [TIME] desc) "
				sqlstr = sqlstr&" WHERE [TIME]>=convert(datetime,floor(convert(float,DATEADD(DAY,"&DBefore&",GETDATE()) ))) "
			  sqlstr = sqlstr&" and [TIME]<convert(datetime,floor(convert(float, DATEADD(DAY,"&DBefore&"+1,GETDATE()) ))) "
				sqlstr = sqlstr&" GROUP BY [TIME],SOURCE_CHANNEL"
				RS.OPEN sqlstr, CONN
				IF NOT RS.EOF THEN
				DO WHILE (NOT RS.EOF)
					if (Rs.Fields("SOURCE_CHANNEL")="OUR_ATM") then
						All_ATM = Rs.Fields("OPERATION")
						All_ATM_Color = checkWarning("ATM_ACQ", Rs.Fields("OPERATION_FAIL"), All_ATM, Rs.Fields("timeinminutes"))
									All_ATM_Color_all = checkWarning_all("ATM_ACQ", Rs.Fields("OPERATION_FAIL"), All_ATM, Rs.Fields("timeinminutes"))
						All_ATM_FAIL = Rs.Fields("OPERATION_FAIL")
						if (All_ATM>0) then
							All_ATM_FAIL_PC=(All_ATM_FAIL*100)/All_ATM
						end if
					elseif (Rs.Fields("SOURCE_CHANNEL")="OUR_BPT") then
						All_BPT = Rs.Fields("OPERATION")
						All_BPT_Color = checkWarning("BPT_ACQ", Rs.Fields("OPERATION_FAIL"), All_BPT, Rs.Fields("timeinminutes"))
									All_BPT_Color_all = checkWarning_all("BPT_ACQ", Rs.Fields("OPERATION_FAIL"), All_BPT, Rs.Fields("timeinminutes"))
						All_BPT_FAIL = Rs.Fields("OPERATION_FAIL")
						if (All_BPT>0) then
							All_BPT_FAIL_PC=(All_BPT_FAIL*100)/All_BPT
						end if
					elseif (Rs.Fields("SOURCE_CHANNEL")="OUR_POS") then
						All_POS = Rs.Fields("OPERATION")
						All_POS_Color = checkWarning("POS_ACQ", Rs.Fields("OPERATION_FAIL"), All_POS, Rs.Fields("timeinminutes"))
									All_POS_Color_all = checkWarning_all("POS_ACQ", Rs.Fields("OPERATION_FAIL"), All_POS, Rs.Fields("timeinminutes"))
						All_POS_FAIL = Rs.Fields("OPERATION_FAIL")
						if (All_POS>0) then
							All_POS_FAIL_PC=(All_POS_FAIL*100)/All_POS
						end if
					elseif (Rs.Fields("SOURCE_CHANNEL")="H2H_BPCRBS") then
						All_H2H_RBS = Rs.Fields("OPERATION")
						All_H2H_RBS_Color = checkWarning("H2H_RBS", Rs.Fields("OPERATION_FAIL"), All_H2H_RBS, Rs.Fields("timeinminutes"))
									All_H2H_RBS_Color_all = checkWarning_all("H2H_RBS", Rs.Fields("OPERATION_FAIL"), All_H2H_RBS, Rs.Fields("timeinminutes"))
						All_H2H_RBS_FAIL = Rs.Fields("OPERATION_FAIL")
						if (All_H2H_RBS>0) then
							All_H2H_RBS_FAIL_PC=(All_H2H_RBS_FAIL*100)/All_H2H_RBS
						end if
					end if
					
					Rs.MoveNext
				LOOP
				END IF
				RS.CLOSE

				if (All_ATM_Color<>"") then
					if (All_ATM_Color = clWarning) then
						All_ATM_Color = " style=""color: #000000; background: "&All_ATM_Color&" "" "
					else
						All_ATM_Color = " style=""background: "&All_ATM_Color&" "" "
					end if
				end if 
				if (All_BPT_Color<>"") then
					if (All_BPT_Color = clWarning) then
						All_BPT_Color = " style=""color: #000000; background: "&All_BPT_Color&" "" "
					else
						All_BPT_Color = " style=""background: "&All_BPT_Color&" "" "
					end if
				end if 
				if (All_POS_Color<>"") then 
					if (All_POS_Color = clWarning) then
						All_POS_Color = " style=""color: #000000; background: "&All_POS_Color&" "" "
					else
						All_POS_Color = " style=""background: "&All_POS_Color&" "" "
					end if
				end if 
				if (All_H2H_RBS_Color<>"") then 
					if (All_H2H_RBS_Color = clWarning) then
						All_H2H_RBS_Color = " style=""color: #000000; background: "&All_H2H_RBS_Color&" "" "
					else
						All_H2H_RBS_Color = " style=""background: "&All_H2H_RBS_Color&" "" "
					end if
				end if 

				if (All_ATM_Color_all<>"") then
					if (All_ATM_Color_all = clWarning) then
						All_ATM_Color_all = " style=""color: #000000; background: "&All_ATM_Color_all&" "" "
					else
						All_ATM_Color_all = " style=""background: "&All_ATM_Color_all&" "" "
					end if
				end if 
				if (All_BPT_Color_all<>"") then
					if (All_BPT_Color_all = clWarning) then
						All_BPT_Color_all = " style=""color: #000000; background: "&All_BPT_Color_all&" "" "
					else
						All_BPT_Color_all = " style=""background: "&All_BPT_Color_all&" "" "
					end if
				end if 
				if (All_POS_Color_all<>"") then 
					if (All_POS_Color_all = clWarning) then
						All_POS_Color_all = " style=""color: #000000; background: "&All_POS_Color_all&" "" "
					else
						All_POS_Color_all = " style=""background: "&All_POS_Color_all&" "" "
					end if
				end if 
				if (All_H2H_RBS_Color_all<>"") then 
					if (All_H2H_RBS_Color_all = clWarning) then
						All_H2H_RBS_Color_all = " style=""color: #000000; background: "&All_H2H_RBS_Color_all&" "" "
					else
						All_H2H_RBS_Color_all = " style=""background: "&All_H2H_RBS_Color_all&" "" "
					end if
				end if

				tblATM = ""
					'-----------SORTING-------------------------------------------------------------------------------------
			'1-ая (верхняя) строка - группа каналов POS_ACQ 
			'2-ая строка - группа каналов ATM
			'3-ья строка – группа каналов H2H_RBS
			'4-ая строка – группа каналов BPT

					Dim tbl_tr8(4,3)
					Dim str_tr8(4)

					tbl_tr8(1, 1)=All_POS_FAIL_PC
					tbl_tr8(1, 2)="All_POS"
					tbl_tr8(1, 3)=1

					tbl_tr8(2, 1)=All_ATM_FAIL_PC
					tbl_tr8(2, 2)="All_ATM"
					tbl_tr8(2, 3)=2

					tbl_tr8(3, 1)=All_H2H_RBS_FAIL_PC
					tbl_tr8(3, 2)="All_H2H_RBS"
					tbl_tr8(3, 3)=3

					tbl_tr8(4, 1)=All_BPT_FAIL_PC
					tbl_tr8(4, 2)="All_BPT"
					tbl_tr8(4, 3)=4


					For j = 1 To 4-1
							For k = j + 1 To 4
									If (tbl_tr8(j,1) < tbl_tr8(k,1)) or ((tbl_tr8(k,1)=0) and (tbl_tr8(j,1)=0) and (tbl_tr8(j,3) > tbl_tr8(k,3)) ) Then
											For l = 1 To 3
													Temp = tbl_tr8(j,l)
													tbl_tr8(j,l) = tbl_tr8(k,l)
													tbl_tr8(k,l) = Temp
											Next

									End If
							Next
					Next

				str_tr8(1) ="<tr><td width=""100px"">POS</td><td onclick=""ChGraph('All_POS',daysBefore)"" width=""70px"" "&All_POS_Color_all&" >"&All_POS&"</td><td onclick=""ChGraph('All_POS',daysBefore)""  width=""70px"" >"&All_POS_FAIL&"</td><td onclick=""ChGraph('All_POS',daysBefore)""  width=""70px"" "&All_POS_Color&" >"&Round(All_POS_FAIL_PC,3)&"</td><tr>"
				str_tr8(2)  = "<tr><td width=""100px"" >ATM</td><td onclick=""ChGraph('All_ATM',daysBefore)"" width=""70px""  "&All_ATM_Color_all&" >"&All_ATM&"</td><td onclick=""ChGraph('All_ATM',daysBefore)"" width=""70px"" >"&All_ATM_FAIL&"</td><td onclick=""ChGraph('All_ATM',daysBefore)"" width=""70px"" "&All_ATM_Color&" >"&Round(All_ATM_FAIL_PC,3)&"</td><tr>"
				str_tr8(3) = "<tr><td width=""100px"" >H2H_RBS</td><td onclick=""ChGraph('All_H2H_RBS',daysBefore)""  width=""70px""  "&All_H2H_RBS_Color_all&" >"&All_H2H_RBS&"</td><td onclick=""ChGraph('All_H2H_RBS',daysBefore)""  width=""70px"" >"&All_H2H_RBS_FAIL&"</td><td onclick=""ChGraph('All_H2H_RBS',daysBefore)""  width=""70px"" "&All_H2H_RBS_Color&" >"&Round(All_H2H_RBS_FAIL_PC,3)&"</td><tr>"
				str_tr8(4) = "<tr><td width=""100px"" >BPT</td><td onclick=""ChGraph('All_BPT',daysBefore)""  width=""70px""  "&All_BPT_Color_all&" >"&All_BPT&"</td><td onclick=""ChGraph('All_BPT',daysBefore)""  width=""70px"" >"&All_BPT_FAIL&"</td><td onclick=""ChGraph('All_BPT',daysBefore)""  width=""70px"" "&All_BPT_Color&" >"&Round(All_BPT_FAIL_PC,3)&"</td><tr>"
			
				For j = 1 To 4
							'tblISS_ACQ = tblISS_ACQ&str_tr8(tbl_tr8(j,3))
							tblATM = tblATM&str_tr8(tbl_tr8(j,3))
				Next		    				
				
				Response.Write tblATM

		elseif (ChName="Channel3DS") then
'------------T=9----------------------------------------------------------------------------------------
				NSPK_VISA = 0
				NSPK_MC = 0
				VISA = 0
				MC = 0
				SOA_USB = 0
				SOA_AGENT = 0
				
				NSPK_VISA_FAIL = 0
				NSPK_MC_FAIL = 0
				VISA_FAIL = 0
				MC_FAIL = 0
				SOA_USB_FAIL = 0
				SOA_AGENT_FAIL = 0
				
				NSPK_VISA_FAIL_PC = 0
				NSPK_MC_FAIL_PC = 0
				VISA_FAIL_PC = 0
				MC_FAIL_PC = 0
				SOA_USB_FAIL_PC = 0
				SOA_AGENT_FAIL_PC = 0
				
				VISA_Color = ""
				NSPK_VISA_Color = ""
				MC_Color = ""
				NSPK_MC_Color = ""
				SOA_AGENT_Color = ""
				SOA_USB_Color = ""

					VISA_Color_all = ""
				NSPK_VISA_Color_all = ""
				MC_Color_all = ""
				NSPK_MC_Color_all = ""
				SOA_AGENT_Color_all = ""
				SOA_USB_Color_all = ""

				sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION, SUM(OPERATION_FAIL) OPERATION_FAIL, SOURCE_CHANNEL "
					sqlstr = sqlstr&" ,DATEPART(HOUR,[TIME])*60+DATEPART(MINUTE,[TIME]) timeinminutes FROM LOG_VS "
				'sqlstr = sqlstr&" WHERE [TIME]=(select top 1 [TIME] from LOG_VS order by [TIME] desc) "
				sqlstr = sqlstr&" WHERE [TIME]>=convert(datetime,floor(convert(float,DATEADD(DAY,"&DBefore&",GETDATE()) ))) "
				sqlstr = sqlstr&" and [TIME]<convert(datetime,floor(convert(float, DATEADD(DAY,"&DBefore&"+1,GETDATE()) ))) "
				sqlstr = sqlstr&" and ((SERVICE='3D-Secure' and SOURCE_CHANNEL in ('NSPK_VISA','NSPK_MasterCard','VISA','MasterCard')) or SERVICE='SOA_AGENT' or SERVICE='SOA_USB') "
				sqlstr = sqlstr&" GROUP BY [TIME],SOURCE_CHANNEL"
				RS.OPEN sqlstr, CONN
				IF NOT RS.EOF THEN
				DO WHILE (NOT RS.EOF)
					if (Rs.Fields("SOURCE_CHANNEL")="NSPK_VISA") then
						NSPK_VISA = Rs.Fields("OPERATION")
						NSPK_VISA_Color = checkWarning("NSPK_VISA_3DS", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"), Rs.Fields("timeinminutes"))
									NSPK_VISA_Color_all = checkWarning_all("NSPK_VISA_3DS", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"), Rs.Fields("timeinminutes"))
						NSPK_VISA_FAIL = Rs.Fields("OPERATION_FAIL")
						if (NSPK_VISA>0) then
							NSPK_VISA_FAIL_PC=(NSPK_VISA_FAIL*100)/NSPK_VISA
						end if
					elseif (Rs.Fields("SOURCE_CHANNEL")="NSPK_MasterCard") then
						NSPK_MC = Rs.Fields("OPERATION")
						NSPK_MC_Color = checkWarning("NSPK_MC_3DS", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"), Rs.Fields("timeinminutes"))
									NSPK_MC_Color_all = checkWarning_all("NSPK_MC_3DS", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"), Rs.Fields("timeinminutes"))
						NSPK_MC_FAIL = Rs.Fields("OPERATION_FAIL")
						if (NSPK_MC>0) then
							NSPK_MC_FAIL_PC=(NSPK_MC_FAIL*100)/NSPK_MC
						end if
					elseif (Rs.Fields("SOURCE_CHANNEL")="VISA") then
						VISA = Rs.Fields("OPERATION")
						VISA_Color = checkWarning("VISA_3DS", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"), Rs.Fields("timeinminutes"))
									VISA_Color_all = checkWarning_all("VISA_3DS", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"), Rs.Fields("timeinminutes"))
						VISA_FAIL = Rs.Fields("OPERATION_FAIL")
						if (VISA>0) then
							VISA_FAIL_PC=(VISA_FAIL*100)/VISA
						end if
					elseif (Rs.Fields("SOURCE_CHANNEL")="MasterCard") then
						MC = Rs.Fields("OPERATION")
						MC_Color = checkWarning("MC_3DS", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"), Rs.Fields("timeinminutes"))
									MC_Color_all = checkWarning_all("MC_3DS", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"), Rs.Fields("timeinminutes"))
						MC_FAIL = Rs.Fields("OPERATION_FAIL")
						if (MC>0) then
							MC_FAIL_PC=(MC_FAIL*100)/MC
						end if		
					elseif (Rs.Fields("SOURCE_CHANNEL")="RBS") then
						SOA_USB = Rs.Fields("OPERATION")
						SOA_USB_Color = checkWarning("SOA_USB", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"), Rs.Fields("timeinminutes"))
									SOA_USB_Color_all = checkWarning_all("SOA_USB", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"), Rs.Fields("timeinminutes"))
						SOA_USB_FAIL = Rs.Fields("OPERATION_FAIL")
						if (SOA_USB>0) then
							SOA_USB_FAIL_PC=(SOA_USB_FAIL*100)/SOA_USB
						end if	
					elseif (Rs.Fields("SOURCE_CHANNEL")="OUR_POS") then
						SOA_AGENT = Rs.Fields("OPERATION")
						SOA_AGENT_Color = checkWarning("SOA_AGENT", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"), Rs.Fields("timeinminutes"))
									SOA_AGENT_Color_all = checkWarning_all("SOA_AGENT", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"), Rs.Fields("timeinminutes"))
						SOA_AGENT_FAIL = Rs.Fields("OPERATION_FAIL")
						if (SOA_AGENT>0) then
							SOA_AGENT_FAIL_PC=(SOA_AGENT_FAIL*100)/SOA_AGENT
						end if			
					end if
					
					Rs.MoveNext
				LOOP
				END IF
				RS.CLOSE

				if (VISA_Color<>"") then 
					if (VISA_Color = clWarning) then 
						VISA_Color = " style=""color: #000000; background: "&VISA_Color&" "" "
					else
						VISA_Color = " style=""background: "&VISA_Color&" "" "
					end if 
				end if 
				if (NSPK_VISA_Color<>"") then 
					if (NSPK_VISA_Color = clWarning) then 
						NSPK_VISA_Color = " style=""color: #000000; background: "&NSPK_VISA_Color&" "" "
					else
						NSPK_VISA_Color = " style=""background: "&NSPK_VISA_Color&" "" "
					end if
				end if 
				if (MC_Color<>"") then 
					if (MC_Color = clWarning) then 
						MC_Color = " style=""color: #000000; background: "&MC_Color&" "" "
					else
						MC_Color = " style=""background: "&MC_Color&" "" "
					end if
				end if 
				if (NSPK_MC_Color<>"") then 
					if (NSPK_MC_Color = clWarning) then 
						NSPK_MC_Color = " style=""color: #000000; background: "&NSPK_MC_Color&" "" "
					else
						NSPK_MC_Color = " style=""background: "&NSPK_MC_Color&" "" "
					end if
				end if
				if (SOA_AGENT_Color<>"") then 
					if (SOA_AGENT_Color = clWarning) then
						SOA_AGENT_Color = " style=""color: #000000; background: "&SOA_AGENT_Color&" "" "
					else
						SOA_AGENT_Color = " style=""background: "&SOA_AGENT_Color&" "" "
					end if
				end if
				if (SOA_USB_Color<>"") then 
					if (SOA_USB_Color = clWarning) then
						SOA_USB_Color = " style=""color: #000000; background: "&SOA_USB_Color&" "" "
					else
						SOA_USB_Color = " style=""background: "&SOA_USB_Color&" "" "
					end if
				end if

					if (VISA_Color_all<>"") then 
					if (VISA_Color_all = clWarning) then 
						VISA_Color_all = " style=""color: #000000; background: "&VISA_Color_all&" "" "
					else
						VISA_Color_all = " style=""background: "&VISA_Color_all&" "" "
					end if 
				end if 
				if (NSPK_VISA_Color_all<>"") then 
					if (NSPK_VISA_Color_all = clWarning) then 
						NSPK_VISA_Color_all = " style=""color: #000000; background: "&NSPK_VISA_Color_all&" "" "
					else
						NSPK_VISA_Color_all = " style=""background: "&NSPK_VISA_Color_all&" "" "
					end if
				end if 
				if (MC_Color_all<>"") then 
					if (MC_Color_all = clWarning) then 
						MC_Color_all = " style=""color: #000000; background: "&MC_Color_all&" "" "
					else
						MC_Color_all = " style=""background: "&MC_Color_all&" "" "
					end if
				end if 
				if (NSPK_MC_Color_all<>"") then 
					if (NSPK_MC_Color_all = clWarning) then 
						NSPK_MC_Color_all = " style=""color: #000000; background: "&NSPK_MC_Color_all&" "" "
					else
						NSPK_MC_Color_all = " style=""background: "&NSPK_MC_Color_all&" "" "
					end if
				end if
				if (SOA_AGENT_Color_all<>"") then 
					if (SOA_AGENT_Color_all = clWarning) then
						SOA_AGENT_Color_all = " style=""color: #000000; background: "&SOA_AGENT_Color_all&" "" "
					else
						SOA_AGENT_Color_all = " style=""background: "&SOA_AGENT_Color_all&" "" "
					end if
				end if
				if (SOA_USB_Color_all<>"") then 
					if (SOA_USB_Color_all = clWarning) then
						SOA_USB_Color_all = " style=""color: #000000; background: "&SOA_USB_Color_all&" "" "
					else
						SOA_USB_Color_all = " style=""background: "&SOA_USB_Color_all&" "" "
					end if
				end if

				tbl3DS = ""
				'-----------SORTING-------------------------------------------------------------------------------------
			'1-ая (верхняя) строка - группа каналов SOA_USB 
			'2-ая строка - группа каналов SOA_AGENT 
			'3-ья строка – группа каналов NSPK_VISA (переименовать в NSPK_VISA_3DS)
			'4-ая строка – группа каналов NSPK_MC (переименовать в NSPK_MC_3DS)
			'5-ая строка – группа каналов VISA (переименовать в VISA_3DS)
			'4-ая строка – группа каналов MC (переименовать в MC _3DS)

					Dim tbl_tr9(6,3)
					Dim str_tr9(6)

					tbl_tr9(1, 1)=SOA_USB_FAIL_PC
					tbl_tr9(1, 2)="SOA_USB"
					tbl_tr9(1, 3)=1

					tbl_tr9(2, 1)=SOA_AGENT_FAIL_PC
					tbl_tr9(2, 2)="SOA_AGENT"
					tbl_tr9(2, 3)=2

					tbl_tr9(3, 1)=NSPK_VISA_FAIL_PC
					tbl_tr9(3, 2)="NSPK_VISA"
					tbl_tr9(3, 3)=3

					tbl_tr9(4, 1)=NSPK_MC_FAIL_PC
					tbl_tr9(4, 2)="NSPK_MC"
					tbl_tr9(4, 3)=4

					tbl_tr9(5, 1)=VISA_FAIL_PC
					tbl_tr9(5, 2)="VISA"
					tbl_tr9(5, 3)=5

					tbl_tr9(6, 1)=MC_FAIL_PC
					tbl_tr9(6, 2)="MC"
					tbl_tr9(6, 3)=6


					For j = 1 To 6-1
							For k = j + 1 To 6
									If (tbl_tr9(j,1) < tbl_tr9(k,1)) or ((tbl_tr9(k,1)=0) and (tbl_tr9(j,1)=0) and (tbl_tr9(j,3) > tbl_tr9(k,3)) ) Then
											For l = 1 To 3
													Temp = tbl_tr9(j,l)
													tbl_tr9(j,l) = tbl_tr9(k,l)
													tbl_tr9(k,l) = Temp
											Next

									End If
							Next
					Next

				str_tr9(1) = "<tr><td width=""100px"" >SOA_USB</td><td onclick=""ChGraph('SOA_USB',daysBefore)"" width=""70px""   "&SOA_USB_Color_all&"  >"&SOA_USB&"</td><td onclick=""ChGraph('SOA_USB',daysBefore)"" width=""70px""  >"&SOA_USB_FAIL&"</td><td onclick=""ChGraph('SOA_USB',daysBefore)"" width=""70px""  "&SOA_USB_Color&" >"&Round(SOA_USB_FAIL_PC,3)&"</td><tr>"
				str_tr9(2) = "<tr><td width=""100px"" >SOA_AGENT</td><td onclick=""ChGraph('SOA_AGENT',daysBefore)"" width=""70px""   "&SOA_AGENT_Color_all&"  >"&SOA_AGENT&"</td><td onclick=""ChGraph('SOA_AGENT',daysBefore)"" width=""70px""  >"&SOA_AGENT_FAIL&"</td><td onclick=""ChGraph('SOA_AGENT',daysBefore)"" width=""70px""  "&SOA_AGENT_Color&" >"&Round(SOA_AGENT_FAIL_PC,3)&"</td><tr>"
				str_tr9(3) = "<tr><td width=""100px"" >NSPK_VISA_3DS</td><td onclick=""ChGraph('NSPK_VISA',daysBefore)"" width=""70px""  "&NSPK_VISA_Color_all&"  >"&NSPK_VISA&"</td><td onclick=""ChGraph('NSPK_VISA',daysBefore)"" width=""70px"" >"&NSPK_VISA_FAIL&"</td><td onclick=""ChGraph('NSPK_VISA',daysBefore)"" width=""70px"" "&NSPK_VISA_Color&" >"&Round(NSPK_VISA_FAIL_PC,3)&"</td><tr>"
				str_tr9(4) = "<tr><td width=""100px"" >NSPK_MC_3DS</td><td onclick=""ChGraph('NSPK_MC',daysBefore)""  width=""70px""  "&NSPK_MC_Color_all&" >"&NSPK_MC&"</td><td onclick=""ChGraph('NSPK_MC',daysBefore)"" width=""70px""  >"&NSPK_MC_FAIL&"</td><td onclick=""ChGraph('NSPK_MC',daysBefore)"" width=""70px""  "&NSPK_MC_Color&" >"&Round(NSPK_MC_FAIL_PC,3)&"</td><tr>"
				str_tr9(5) = "<tr><td width=""100px"" >VISA_3DS</td><td onclick=""ChGraph('VISA',daysBefore)""  width=""70px""  "&VISA_Color_all&" >"&VISA&"</td><td onclick=""ChGraph('VISA',daysBefore)"" width=""70px""  >"&VISA_FAIL&"</td><td onclick=""ChGraph('VISA',daysBefore)""  width=""70px"" "&VISA_Color&" >"&Round(VISA_FAIL_PC,3)&"</td><tr>"
				str_tr9(6) = "<tr><td width=""100px"" >MC_3DS</td><td onclick=""ChGraph('MC',daysBefore)""  width=""70px""  "&MC_Color_all&" >"&MC&"</td><td onclick=""ChGraph('MC',daysBefore)"" width=""70px""  >"&MC_FAIL&"</td><td onclick=""ChGraph('MC',daysBefore)"" width=""70px""  "&MC_Color&" >"&Round(MC_FAIL_PC,3)&"</td><tr>"

				For j = 1 To 6
							'tblISS_ACQ = tblISS_ACQ&str_tr9(tbl_tr9(j,3))
							tbl3DS = tbl3DS&str_tr9(tbl_tr9(j,3))
				Next							


				Response.Write tbl3DS

		end if


end function
'----------------------------------------------------------------------------------------------------
'-------END: Channel Groups Table--------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------


Conn.Close
set Cmd = Nothing
set Rs = Nothing
set Conn = Nothing
%>
