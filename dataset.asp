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
	
end if

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

Conn.Close
set Cmd = Nothing
set Rs = Nothing
set Conn = Nothing
%>
