<%@ language="VBScript"%><!-- #include file="const.asp" --><!-- #include file="common.asp" -->
<%
' Вывод детализации по 1-5 параметрам

T = Request("T")
if isNumeric(T) then T=cint(T) else T=0 end if

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

if Len(FUserName)=0 then
  Conn.Close
  set Conn = Nothing
  set Cmd = Nothing
  Response.Write("<html><body><div style='text-align: center;'><span style='font-size: 14pt; font-weight: 600; color: #800000}'>Для пользователя "&Auth_Name&" доступ не определен.</span></div></body></html>")
else

' Dim sajaxText
' Set sajaxText = New Sajax2
' sajaxText.export_function("ShowMessageInfo;")
' sajaxText.handle_client_request()

'function ShowMessageInfo(mesID)
    ' sql="SELECT * FROM [Messages_Type] WHERE [MsgID]="&mesID
    ' set oRs = DAL.ExecuteSQL (sql)
    ' if not oRs.Eof then
        ' do while not oRs.Eof
	   	    ' res=oRs.Fields(0)
            ' oRs.MoveNext
        ' loop
    ' end if
    ' DAL.CloseRS(oRs)
    'ShowMessageInfo = "1231231уамукмек2312132"
'end function

' BEGIN ***********************************************************************
' ФОРМИРОВАНИЕ ГОРИЗОНТАЛЬНОЙ ИНФОРМАЦИОННОЙ ПАНЕЛИ ПО ОСНОВНЫМ ПАРАМЕТРАМ
' *****************************************************************************

set Rs=Server.CreateObject("ADODB.Recordset")
Rs.Open "SELECT * FROM Tags WHERE (TagID='Main1')", Conn
Value1=Rs.Fields("Value")
Value1all=Rs.Fields("ValueDetail")
if Rs.Fields("ValueDetail") = 0 then
  Value1proc=0
else
  Value1proc=100*Rs.Fields("Value")/Rs.Fields("ValueDetail")
end if
Color1=clNormal
Main1_SetHiHi=Rs.Fields("SetHiHi")
if Value1proc >= Rs.Fields("SetHi") then Color1=clWarning end if
if Value1proc >= Main1_SetHiHi then Color1=clError end if
Rs.Close

Color2=0
dim Colors2(3,2)
dim Text2(3,2)
for i=1 to 3
  for j=1 to 2
    Colors2(i, j)=0
	Text2(i,j)=""
  next
next  

Rs.Open "SELECT ISNULL(Max(Value), 0) FROM Tags_History WHERE (TagID='Main2All') AND (DT > GETDATE()-1.0/6)", Conn
Value2A_max=Rs.Fields(0)
Rs.Close

Rs.Open "SELECT ISNULL(Max(Value), 0) FROM Tags_History WHERE (TagID='Main2AllBPT') AND (DT > GETDATE()-1.0/6)", Conn
Value2AllBPT_max=Rs.Fields(0)
Rs.Close

Rs.Open "SELECT ISNULL(Max(Value), 0) FROM Tags_History WHERE (TagID='Main2Centr') AND (DT > GETDATE()-1.0/6)", Conn
Value2Centr_max=Rs.Fields(0)
Rs.Close

Rs.Open "SELECT TOP 1 [NAME] FROM NV_Operations AS O LEFT OUTER JOIN V_Resp_code AS C ON O.RESPONSE_CODE=C.Resp_code WHERE ISNULL(C.IsFailed, 0)<>0 ORDER BY [QUANTITY] DESC, [NAME]", Conn
with RS
  if (not .eof) or (not .bof) then
    Channel1=.Fields(0)
  Else
    Channel1=0
  end if
end with
Rs.Close

Rs.Open "SELECT TagID, [Value], SetHi, SetHiHi, [ValueDetail] FROM Tags WHERE (TagID like 'Main2%') ORDER BY TagID", Conn
do while not Rs.Eof
'  if Rs.Fields(0) = "Main2All" then 
'    Text2(1,1)=Rs.Fields(1)
'    if Rs.Fields(1)>Rs.Fields(2) then Colors2(1,1)=1 end if
'    if Rs.Fields(1)>Rs.Fields(3) then Colors2(1,1)=2 end if
'	Main2A_SetHiHi=Rs.Fields(3)
'  end if
    if Rs.Fields(0) = "Main2All24" then 
    Text2(1,1)=Rs.Fields(1)
    if Rs.Fields(1)>Rs.Fields(2) then Colors2(1,1)=1 end if
    if Rs.Fields(1)>Rs.Fields(3) then Colors2(1,1)=2 end if
	Main2A24_SetHi=Rs.Fields(2)
	Main2A24_SetHiHi=Rs.Fields(3)
  end if
'  if Rs.Fields(0) = "Main2AllBPT" then 
'    Text2(1,2)=Rs.Fields(1)
'    if Rs.Fields(1)>Rs.Fields(2) then Colors2(1,2)=1 end if
'    if Rs.Fields(1)>Rs.Fields(3) then Colors2(1,2)=2 end if
'	Main2AllBPT_SetHiHi=Rs.Fields(3)
'  end if
      if Rs.Fields(0) = "Main2AllBPT24" then 
    Text2(1,2)=Rs.Fields(1)
    if Rs.Fields(1)>Rs.Fields(2) then Colors2(1,2)=1 end if
    if Rs.Fields(1)>Rs.Fields(3) then Colors2(1,2)=2 end if
	Main2AllBPT24_SetHi=Rs.Fields(2)
	Main2AllBPT24_SetHiHi=Rs.Fields(3)
  end if
  if Rs.Fields(0) = "Main2Fil" then 
    Text2(2,1)=Rs.Fields(1)
    if Rs.Fields(1)>Rs.Fields(2) then Colors2(2,1)=1 end if
    if Rs.Fields(1)>Rs.Fields(3) then Colors2(2,1)=2 end if
    if Rs.Fields(4)>0 then Colors2(2,1)=3 end if
	Main2F_SetHi=Rs.Fields(2)
	Main2F_SetHiHi=Rs.Fields(3)
  end if
  if Rs.Fields(0) = "Main2FilBPT" then 
    Text2(2,2)=Rs.Fields(1)
    if Rs.Fields(1)>Rs.Fields(2) then Colors2(2,2)=1 end if
    if Rs.Fields(1)>Rs.Fields(3) then Colors2(2,2)=2 end if
    if Rs.Fields(4)>0 then Colors2(2,2)=2 end if
  end if
 ' if Rs.Fields(0) = "Main2Centr" then 
  '  Text2(3,1)=Rs.Fields(1)
   ' if Rs.Fields(1)>Rs.Fields(2) then Colors2(3,1)=1 end if
    'if Rs.Fields(1)>Rs.Fields(3) then Colors2(3,1)=2 end if
    'if Rs.Fields(4)>0 then Colors2(3,1)=2 end if
	'Main2Centr_SetHi=Rs.Fields(2)
	'Main2Centr_SetHiHi=Rs.Fields(3)
  'end if
  if Rs.Fields(0) = "Main2Centr" then  'Недоступность АТМ по центральной схеме подключения
    Text2(3,1)=Rs.Fields(1)
    if Rs.Fields(1)=1 then Colors2(3,1)=1 end if
    if Rs.Fields(1)>1 then Colors2(3,1)=2 end if
	Main2Centr_SetHi=Rs.Fields(2)
	Main2Centr_SetHiHi=Rs.Fields(3)
  end if  
  Rs.MoveNext
loop
if Value2A_max<Main2A_SetHiHi then Value2A_max=Main2A_SetHiHi end if
if Value2Centr_max<Main2Centr_SetHiHi then Value2Centr_max=Main2Centr_SetHiHi end if
if Value2AllBPT_max<Main2AllBPT_SetHiHi then Value2AllBPT_max=Main2AllBPT_SetHiHi end if

function SetColorByN(ColorN)
  if ColorN=0 then 
    SetColorByN=clNormal
  else
    if ColorN=1 then 
      SetColorByN=clWarning
    else
      SetColorByN=clError
    end if
  end if
end function

ColorATM=0
if Colors2(1, 1) > ColorATM then ColorATM=Colors2(1, 1) end if
if Colors2(2, 1) > ColorATM then ColorATM=Colors2(2, 1) end if
ColorATM=SetColorByN(ColorATM)

ColorBPT=0
if Colors2(1, 2) > ColorBPT then ColorBPT=Colors2(1, 2) end if
if Colors2(2, 2) > ColorBPT then ColorBPT=Colors2(2, 2) end if
ColorBPT=SetColorByN(ColorBPT)

ColorCSP=SetColorByN(Colors2(3, 1))

for i=1 to 3
  for j=1 to 2
    if Color2 < Colors2(i, j) then Color2=Colors2(i, j) end if
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

Color2=SetColorByN(Color2)

Rs.Close

Rs.Open "SELECT [Value], [Prop_Crit], count([Value]) FROM [Tags] WHERE ([FileID]='CV') and (Prop_Active=1) GROUP BY [Value], [Prop_Crit] ORDER BY [Prop_Crit]", Conn
Value3total=0
Value3linkdown=0
Value3proc=0
Color3=clNormal
if not Rs.Eof then 
  do while not Rs.Eof
    if Rs.Fields(0) = 0 then 
	  Value3linkdown=Value3linkdown+Rs.Fields(2)
      if Rs.Fields(1)=1 then Color3=clWarning
      if Rs.Fields(1)=2 then Color3=clError
	end if
    Value3total=Value3total+Rs.Fields(2)
    Rs.MoveNext
  loop
  Value3proc=round(Value3linkdown*100/Value3total)
end if
Value3=Value3linkdown
Rs.Close

Rs.Open "SELECT * FROM Tags WHERE (TagID='Main5SMS')", Conn
Color5=clNormal
if not Rs.Eof then 
  Value5=Rs.Fields("Value")
  Main5_SetHi=Rs.Fields("SetHi")
  Main5_SetHiHi=Rs.Fields("SetHiHi")
  if Value5 >= Main5_SetHi then Color5=clWarning end if
  if Value5 >= Main5_SetHiHi then Color5=clError end if
end if
Rs.Close

Table2=""
Rs.Open "SELECT Max(LastState) FROM vw_Messages", Conn
Color6=clNormal
if Rs.Fields(0)=1 then Color6=clWarning end if
if Rs.Fields(0)=2 then Color6=clError end if
Rs.Close


' *****************************************************************************
' ФОРМИРОВАНИЕ ГОРИЗОНТАЛЬНОЙ ИНФОРМАЦИОННОЙ ПАНЕЛИ ПО ОСНОВНЫМ ПАРАМЕТРАМ
' END *************************************************************************

set Cmd=Server.CreateObject("ADODB.Command")

if T=0 then
  Cmd.ActiveConnection=Conn
  Cmd.CommandType=adCmdText
  ' Cmd.CommandText="EXEC [sp_Diagram_NV] @DS=1"
  ' set Rs=Cmd.Execute
  
  ' TotalOperations=0
  ' CurrentProc=0
  ' gr="underfound"
  ' series1=""
  ' series2=""
  ' if not Rs.Eof then DT_FILE=Rs.Fields(4) end if
  ' do while not Rs.Eof
    ' TotalOperations=TotalOperations+Rs.Fields(2)
    ' if gr=Rs.Fields(0) then
      ' CurrentProc=CurrentProc+Rs.Fields(3)
  	' series1=series1&"['"&IIF(gr="", "?", replace(gr, " ", "<br />"))&"', "&replace(CurrentProc, ",", ".")&"],"
    ' end if
    ' if gr<>Rs.Fields(0) then 
      ' CurrentProc=Rs.Fields(3)
  	' gr=Rs.Fields(0)
    ' end if
    ' if Rs.Fields(1)=0 then
      ' series2=series2&"{name: '', y: "&replace(FormatNumber(Rs.Fields(3), 2, -1), ",", ".")&", color: '#00CC00'}, "
    ' else
      ' series2=series2&"{name: '"&IIF(Rs.Fields(0)="", "?", Rs.Fields(0))&": "&Rs.Fields(2)&"', y: "&replace(FormatNumber(Rs.Fields(3), 2, -1), ",", ".")&", color: '#FF3300'}, "
    ' end if
    ' Rs.MoveNext
  ' loop
  ' series1=left(series1, len(series1)-1)
  ' series2=left(series2, len(series2)-2)
  ' Rs.Close
  
  Cmd.CommandText="EXEC [sp_Diagram_NV] @DS=11"
  set Rs=Cmd.Execute
  series1=""
  if not Rs.Eof then DT_FILE=Rs.Fields("DT_FILE") end if
  do while not Rs.Eof
    if Rs.Fields(0)="Другие" then
	    if (ISNULL(Rs.Fields(1)) or ISNULL(Rs.Fields(2))) then
			series1=series1&"{name: '"&Rs.Fields(0)&":<br />"&replace(FormatNumber(0, 0, -1), ",", ".")&" / ', y: "&replace(FormatNumber(0, 0, -1), ",", ".")&"},"
		else
			series1=series1&"{name: '"&Rs.Fields(0)&":<br />"&replace(FormatNumber(Rs.Fields(1), 0, -1), ",", ".")&" / ', y: "&replace(FormatNumber(Rs.Fields(2), 0, -1), ",", ".")&"},"
		end if
	else
		series1=series1&"{name: '"&Rs.Fields(0)&": "&replace(FormatNumber(Rs.Fields(1), 0, -1), ",", ".")&" / ', y: "&replace(FormatNumber(Rs.Fields(2), 0, -1), ",", ".")&"},"
	end if
    Rs.MoveNext
  loop
  series1=left(series1, len(series1)-1)
  Rs.Close

  Cmd.CommandText="EXEC [sp_Diagram_NV] @DS=2"
  set Rs=Cmd.Execute
  series2=""
  do while not Rs.Eof
  	series2=series2&"{name: '"&Rs.Fields(0)&": "&replace(FormatNumber(Rs.Fields(1), 0, -1), ",", ".")&" / ', y: "&replace(FormatNumber(Rs.Fields(2), 0, -1), ",", ".")&"},"
    Rs.MoveNext
  loop
  if (series2<>"") then
	series2=left(series2, len(series2)-1)
  end if
  Rs.Close
  
  Cmd.CommandText="EXEC [sp_Diagram_NV] @DS=3"
  set Rs=Cmd.Execute
  series3=""
  do while not Rs.Eof
  	series3=series3&"{name: '"&Rs.Fields(0)&": "&replace(FormatNumber(Rs.Fields(1), 0, -1), ",", ".")&" / ', y: "&replace(FormatNumber(Rs.Fields(2), 0, -1), ",", ".")&"},"
    Rs.MoveNext
  loop
  if (series3<>"") then
	series3=left(series3, len(series3)-1)
  end if  
  Rs.Close
end if
CurrentTime = DateTimeFormat(Now, "yyyy, mm, dd, hh, nn")
if T=1 then
	dim series(8), CID(8), CNM(8)
	for i=1 to 8
	  series(i)=""
	  CID(i)=""
	  CNM(i)=""
	next
	L=0
	
	SQL_="SELECT TOP (100) PERCENT A.CHANNEL_ID, A.CHANNEL, A.Qdown, A.LastDown, B.DT, B.VALUE FROM "&_
	"(SELECT CHANNEL_ID, CHANNEL, COUNT(*) AS Qdown, MAX(DT) AS LastDown  "&_
	"FROM vw_Channel_History WHERE (DT>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE())))) AND (DT<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+1)) AND (VALUE=0)  "&_
	"GROUP BY CHANNEL_ID, CHANNEL) AS A LEFT OUTER JOIN  "&_
	"(SELECT DT, CHANNEL_ID, CHANNEL, [VALUE] FROM vw_Channel_History AS vw_Channel_History_2 "&_
	"WHERE (DT = (SELECT MAX(DT) AS Expr1 FROM vw_Channel_History AS vw_Channel_History_1))) AS B ON A.CHANNEL_ID = B.CHANNEL_ID "&_
	"ORDER BY 3 desc, 2"
	Rs.Open SQL_, Conn
	tblChannel=""
	if not Rs.Eof then
		CHID1=Rs.Fields(0)
		CHNM1=Rs.Fields(1)
		do while not Rs.Eof
			if Rs.Fields(5)=0 then cl=clError else cl=clNormal end if
			tblChannel=tblChannel&"<tr id=""r"&Rs.Fields(0)&""" onclick=""ChGraph("&Rs.Fields(0)&", '"&Rs.Fields(1)&"', jsDate)"">"&_
			  "<td>"&Rs.Fields(0)&"</td>"&_
			  "<td>"&Rs.Fields(1)&"</td>"&_
			  "<td>"&Rs.Fields(2)&"</td>"&_
			  "<td>"&DateTimeFormat(Rs.Fields(3), "dd.mm.yy hh:mm:ss")&"</td>"&_
			  "<td style='text-align: left; color: "&cl&"'>"&DateTimeFormat(Rs.Fields(4), "dd.mm.yy hh:mm:ss")&"</td>"&_
			  "</tr>"

			if L<8 then
				L=L+1
				CID(L)=Rs.Fields(0)
				CNM(L)=Rs.Fields(1)
			end if
			Rs.MoveNext
		loop
	else
	  tblChannel=tblChannel&"<tr><td colspan=5>Нет данных</td></tr>"
	  CHID1=0
	  CHNM1=""
	end if
	Rs.Close

	Rs.Open "SELECT Max(DT_FILE) FROM CV_Channel", Conn
	if not Rs.Eof then 
		DT_FILE=Rs.Fields(0)
	end if
	Rs.Close


	series1=""
	SQL_="SELECT DT, TagID, -1*[Value] as [Value] FROM Tags_History WHERE (TagID='Main3') AND "&_
    "(DT>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE())))) AND (DT<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+1)) "&_
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
			series1=series1&","&vbCrLf&"{"&m&"x: Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), y: 8.5}"
			Rs.MoveNext
		loop
	end if
	Rs.Close

	series2=""
	SQL_= "SELECT * FROM Tags_History WHERE (TagID='Main3down') AND "&_
    "(DT>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE())))) AND (DT<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+1)) "&_
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
		series2=series2&"{x: Date.UTC("&DateTimeFormat(Int(Now), "yyyy, mm, dd, hh, nn")&"), y: 8.5},"&_
						"{x: Date.UTC("&DateTimeFormat(Int(Now)+1-1/86400, "yyyy, mm, dd, hh, nn")&"), y: 8.5}"
	end if
	Rs.Close
	
	AllSeries=""
	for i=1 to L
  	  SQL_="SELECT dateAdd(ss,-1*DATEPART(ss, DT),dateAdd(ms,-1*DATEPART(ms, DT),DT)) AS DT,[CHANNEL_ID],[CHANNEL],[VALUE] FROM vw_Channel_History "&_
	  "WHERE (CHANNEL_ID="&CID(i)&") AND "&_
      "(DT>=CONVERT(datetime, FLOOR(CONVERT(float, GETDATE())))) AND (DT<CONVERT(datetime, FLOOR(CONVERT(float, GETDATE()))+1)) "&_
	  "GROUP BY dateAdd(ss,-1*DATEPART(ss, DT),dateAdd(ms,-1*DATEPART(ms, DT),DT)),[CHANNEL_ID],[CHANNEL],[VALUE] ORDER BY DT"
	  Rs.Open SQL_, Conn
	  do while not Rs.Eof
	    v=8.5-i
	    v=replace(v, ",", ".")
        if Rs.Fields("Value")=0 then 
		  series(i)=series(i)&vbCrLf&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("DT"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"},"
	    end if
		Rs.MoveNext
	  loop
	  Rs.Close
	  if series(i)<>"" then
		series(i)=left(series(i), len(series(i))-1)
		series(i)=", { name: '"&CNM(i)&"', type: 'scatter', data: ["&series(i)&"]}"  
	  end if
	  AllSeries=AllSeries+series(i)
	next
	
end if

if T=2 then
	DT_FILE=""
	SQL_="SELECT [DT_FILE], A.[BRANCH_CODE], ISNULL([Name], A.[BRANCH_CODE]) AS [NAME],[ATM],[ATM_LINK],[ATM_ERR],[ATM_LINK_ERR],[ATM_LINK24]*100.0/[ATM], [ATM_LINK_ERR]*100/[ATM] "&_
		 "FROM [AV_ATMStat] AS A LEFT OUTER JOIN V_Branch_code AS C ON A.[BRANCH_CODE]=C.[Branch_code] ORDER BY 8 DESC"
	Rs.Open SQL_, Conn
	tblAtmLink=""
	if not Rs.Eof then
		DT_FILE=Rs.Fields(0)
		BRANCH_CODE=Rs.Fields(1)
		BRANCH_NAME=Rs.Fields(2)
		do while not Rs.Eof
		
				ColorstrATM=""
				ColorstrLNK_ERR=""
				ColorstrLNK_24=""
		
		if (ColorATM<>clNormal) then
		
		 vfieldcol=Rs.Fields(7)
		 if CInt(vfieldcol)>Cint(Main2A24_SetHi)	 then
		    ColorstrLNK_24=clWarning
		 end if
		 if CInt(vfieldcol)>Cint(Main2A24_SetHiHi)	 then
			ColorstrLNK_24=clError
		 end if
		 
		 if Rs.Fields("ATM_LINK_ERR")=Rs.Fields("ATM")	 then
			ColorstrLNK_ERR=clError
		 end if
		
		vfieldcol=Rs.Fields(8)
	    if Cint(vfieldcol)>=Cint(Main2F_SetHi) then
			if ColorstrATM<>clWarning then
				ColorstrATM=clWarning
			end if
		end if
		if Cint(vfieldcol)>=Cint(Main2F_SetHiHi) then
			ColorstrATM=clError
		end if
		if Rs.Fields(8)="100" then
			ColorstrATM=clError
		end if
		
		end if

			tblAtmLink=tblAtmLink&"<tr id=""r"&Rs.Fields(1)&""" onclick=""ChGraph('"&Rs.Fields(1)&"', '"&replace(Rs.Fields(2), """", "")&"', jsDate)""><td>"&Rs.Fields(1)&"</td>"
			for i=3 to 6
			  if (i=6) then
				tblAtmLink=tblAtmLink&"<td style=""background-color:"&ColorstrLNK_ERR&""" >"&Rs.Fields(i)&"</td>"
			  else
				tblAtmLink=tblAtmLink&"<td>"&Rs.Fields(i)&"</td>"
			  end if 

			next
			ww=Round(440*cint(Rs.Fields(7))/100)
			tblAtmLink=tblAtmLink&"<td  style=""background-color:"&ColorstrLNK_24&""" >"&FormatNumber(Rs.Fields(7), 1, -1)&"</td><td style=""text-align: left; background-color: #A0A0A0;"">"&_
			  "<img src=""d.gif"" width="""&ww&""" height=""16"" alt="""" style=""background-color: #CCFFFF; margin-top: 1px;"" />"&_
			  "<div style=""line-height: 18px; margin-top: -17px; margin-left: 2px; color: #000000; font-weight: 700"">"&Rs.Fields(2)&"</div>"&_
			  "</td></tr>"&vbCrLf
			Rs.MoveNext
		loop
	end if
	Rs.Close
end if

if T=3 then
	DT_FILE=""
	SQL_="SELECT [DT_FILE], [LINK_TYPE], [ATM], [ATM_OFFLINE], [IsCentralSchema], [ATM_OFFLINE]*100.0/[ATM] "&_
		 "FROM [LV_ATMStatLink] ORDER BY 6 DESC"
	Rs.Open SQL_, Conn
	tblAtmLink=""
	if not Rs.Eof then
		DT_FILE=Rs.Fields(0)
		LINK_TYPE=Rs.Fields(1)
		do while not Rs.Eof
		 'ColorCSP

				ColorstrLNK_ERR=""
		
		if (ColorCSP<>clNormal) then

		 if Rs.Fields(5)>Main2Centr_SetHiHi	 then
			ColorstrLNK_ERR=clWarning
		 end if
		
		end if
		
			tblAtmLink=tblAtmLink&"<tr id=""r"&Rs.Fields(1)&""" onclick=""ChGraph('"&Rs.Fields(1)&"', jsDate)""><td>"&Rs.Fields(1)&"</td>"
			for i=2 to 4
			  if (i=3) then
			   tblAtmLink=tblAtmLink&"<td style=""background-color:"&ColorstrLNK_ERR&""" >"&Rs.Fields(i)&"</td>"
			  else
			   tblAtmLink=tblAtmLink&"<td>"&Rs.Fields(i)&"</td>"
			  end if
			next
			ww=Round(440*cint(Rs.Fields(5))/100)
			tblAtmLink=tblAtmLink&"<td style=""text-align: left; background-color: #A0A0A0;"">"&_
			  "<img src=""d.gif"" width="""&ww&""" height=""16"" alt="""" style=""background-color: #CCFFFF; margin-top: 1px;"" />"&_
			  "<div style=""line-height: 18px; margin-top: -17px; margin-left: 2px; color: #000000; font-weight: 700"">"&FormatNumber(Rs.Fields(5), 1, -1)&"</div>"&_
			  "</td></tr>"&vbCrLf
			Rs.MoveNext
		loop
	end if
	Rs.Close
end if

if T=4 then
	DT_FILE=""
	SQL_="SELECT [DT_FILE], A.[BRANCH_CODE], ISNULL([Name], A.[BRANCH_CODE]) AS [NAME],[BPT],[BPT_LINK],[BPT_ERR],[BPT_LINK_ERR],[BPT_LINK24]*100.0/[BPT] "&_
		 "FROM [TV_BPTStat] AS A LEFT OUTER JOIN V_Branch_code AS C ON A.[BRANCH_CODE]=C.[Branch_code] ORDER BY 8 DESC"
	Rs.Open SQL_, Conn
	tblAtmLink=""
	if not Rs.Eof then
		DT_FILE=Rs.Fields(0)
		BRANCH_CODE=Rs.Fields(1)
		BRANCH_NAME=Rs.Fields(2)
		do while not Rs.Eof
		
				ColorstrBPT=""
				ColorstrLNK_ERR=""
				ColorstrLNK_24=""
		
		if (ColorBPT<>clNormal) then
		
		 vfieldcol=Rs.Fields(7)
		 if CInt(vfieldcol)>Cint(Main2A24_SetHi)	 then
		    ColorstrLNK_24=clWarning
		 end if
		 if CInt(vfieldcol)>Cint(Main2A24_SetHiHi)	 then
			ColorstrLNK_24=clError
		 end if
		 
		 if Rs.Fields("BPT_LINK_ERR")=Rs.Fields("BPT")	 then
			ColorstrLNK_ERR=clError
		 end if
		
		
		end if		
		
		
			tblAtmLink=tblAtmLink&"<tr id=""r"&Rs.Fields(1)&""" onclick=""ChGraph('"&Rs.Fields(1)&"', '"&replace(Rs.Fields(2), """", "")&"', jsDate)""><td>"&Rs.Fields(1)&"</td>"
			for i=3 to 6
			  if (i=6) then
				tblAtmLink=tblAtmLink&"<td style=""background-color:"&ColorstrLNK_ERR&""" >"&Rs.Fields(i)&"</td>"
			  else
				tblAtmLink=tblAtmLink&"<td>"&Rs.Fields(i)&"</td>"
			  end if 			
			next
			ww=Round(440*cint(Rs.Fields(7))/100)
			tblAtmLink=tblAtmLink&"<td style=""background-color:"&ColorstrLNK_24&""" >"&FormatNumber(Rs.Fields(7), 1, -1)&"</td><td style=""text-align: left; background-color: #A0A0A0;"">"&_
			  "<img src=""d.gif"" width="""&ww&""" height=""16"" alt="""" style=""background-color: #CCFFFF; margin-top: 1px;"" />"&_
			  "<div style=""line-height: 18px; margin-top: -17px; margin-left: 2px; color: #000000; font-weight: 700"">"&Rs.Fields(2)&"</div>"&_
			  "</td></tr>"&vbCrLf
			Rs.MoveNext
		loop
	end if
	Rs.Close
end if

if T=5 then
	DT_FILE=""
	SQL_="SELECT * FROM MV_SMSService ORDER BY [SERVER]"
	Rs.Open SQL_, Conn
	tblSMS=""
	ki=0
	if not Rs.Eof then
		DT_FILE=Rs.Fields(0)
		BRANCH_CODE=Rs.Fields(1)
		BRANCH_NAME=Rs.Fields(2)
		do while not Rs.Eof
			
			if Rs.Fields("WAIT_COUNT")*100.0/Rs.Fields("ALL_COUNT") >= Main5_SetHi or Rs.Fields("REJECTED_COUNT")*100.0/Rs.Fields("ALL_COUNT") >= Main5_SetHi or Rs.Fields("DECLINED_COUNT")*100.0/Rs.Fields("ALL_COUNT") >= Main5_SetHi then
			ColorSMSF=clWarning
			End if
			if Rs.Fields("WAIT_COUNT")*100.0/Rs.Fields("ALL_COUNT") >= Main5_SetHiHi or Rs.Fields("REJECTED_COUNT")*100.0/Rs.Fields("ALL_COUNT") >= Main5_SetHiHi or Rs.Fields("DECLINED_COUNT")*100.0/Rs.Fields("ALL_COUNT") >= Main5_SetHiHi then
			ColorSMSF=clError
			End if
			tblSMS=tblSMS&"<tr style=""background-color: "&ColorSMSF&""" id="""&ki&"""><td>"&DateTimeFormat(Rs.Fields(0), "dd.mm.yy hh:mm")&"</td>"
			ColorSMSF=""
			ki=ki+1
			for i=1 to 8
			  tblSMS=tblSMS&"<td>"&Rs.Fields(i)&"</td>"
			next
			tblSMS=tblSMS&"</tr>"&vbCrLf
			Rs.MoveNext
		loop
	end if
	Rs.Close
	
	SQL_="SELECT COUNT(*) FROM MV_SMSService"
		Rs.Open SQL_, Conn
		SMScount=Rs.Fields(0)
		Rs.Close
		ki=1
		for i=0 to SMScount-1 
		objcountdts=(i+1)*3+SMScount-1 
		SMSServiceStringchrt=SMSServiceStringchrt&" var chart"&i&"; "
		if i=0 then
		SMSServiceStringdiv=SMSServiceStringdiv&"<div id=""container0"" style=""width: 98%; height: 260px; margin: 0 auto; display:""></div>"
		end if
		if i<>0 then
		SMSServiceStringoption=SMSServiceStringoption&" var options"&i&" = options0; "
		SMSServiceStringdiv=SMSServiceStringdiv&" <div id=""container"&i&""" style=""width: 98%; height: 260px; margin: 0 auto; display:none""></div>"
		end if
		SMSServiceString=SMSServiceString&" options"&i&".series[0].data.length=0; "
		SMSServiceString=SMSServiceString&" options"&i&".series[1].data.length=0; "
		SMSServiceString=SMSServiceString&" options"&i&".series[2].data.length=0; "
		SMSServiceString=SMSServiceString&" var obj"&ki&" = eval(""["" + dts["&objcountdts-2&"] + ""]""); "
		SMSServiceString=SMSServiceString&" var obj"&ki+1&" = eval(""["" + dts["&objcountdts-1&"] + ""]""); "
		SMSServiceString=SMSServiceString&" var obj"&ki+2&" = eval(""["" + dts["&objcountdts&"] + ""]""); "
		SMSServiceString=SMSServiceString&" options"&i&".series[0].data = obj"&ki&"; "
		SMSServiceString=SMSServiceString&" options"&i&".series[1].data = obj"&ki+1&"; "
		SMSServiceString=SMSServiceString&" options"&i&".series[2].data = obj"&ki+2&"; "
		ki=ki+3
		SMSServiceString=SMSServiceString&" options"&i&".title.text = dts["&i&"]; "
		SMSServiceString=SMSServiceString&" options"&i&".chart.renderTo='container"&i&"'; "
		SMSServiceString=SMSServiceString&" chart"&i&" = new Highcharts.Chart(options"&i&"); " 
		SMSServiceString=SMSServiceString&" chart"&i&".yAxis[0].addPlotLine({ value:  "&Main5_SetHiHi&", color: '#66FFFF', dashStyle: 'Dash', width: 2, id: 'plot-line-1' }); "
		next
end if

if T=6 then
	Color6aw=clWarning
	BD=Request("BD")
	ED=Request("ED")
	if (BD="") and (ED="") then
	  BD=DateTimeFormat(Now(), "dd.mm.yyyy")
	  ED=DateTimeFormat(Now()+1, "dd.mm.yyyy")
	end if
	CAT_=Request("CAT_")
	if isEmpty(CAT_) then CAT_="" end if

	PRI_=Request("PRI_")
	if isEmpty(PRI_) then PRI_=-1 else if isNumeric(PRI_) then PRI_=cint(PRI_) else PRI_=-1 end if end if

	ELV_=Request("ELV_")
	if isEmpty(ELV_) then ELV_=-1 else if isNumeric(ELV_) then ELV_=cint(ELV_) else ELV_=-1 end if end if

	PROP_=Request("PROP_")
	if isEmpty(PROP_) then PROP_=-1 else if isNumeric(PROP_) then PROP_=cint(PROP_) else PROP_=-1 end if end if

	DT_FILE=""
	SQL_="SELECT M.[MsgTime],C.[Mnemonic],T.[ErrorLevel],T.[Priority],T.[Property],M.[MsgText], T.MsgID, T.Period, T.LastState, T.LastTime "&_
	"FROM [dbo].[Messages] AS M LEFT OUTER JOIN Messages_Type AS T ON CONVERT(int, RIGHT(M.[MsgCode], 3))=T.MsgID LEFT OUTER JOIN Messages_Category AS C ON T.CategoryCode=C.CategoryCode "&_
	"WHERE (M.[MsgTime] > CONVERT(datetime, '"&BD&"', 104)) AND (M.[MsgTime] < CONVERT(datetime, '"&ED&"', 104)) "&_
	" AND (T.MsgID<>10) "
	if CAT_<>"" then
	  SQL_=SQL_&" AND (C.Mnemonic='"&CAT_&"') "
	end if
	if PRI_<>-1 then
	  SQL_=SQL_&" AND (T.[Priority]='"&PRI_&"') "
	end if
	if ELV_<>-1 then
	  SQL_=SQL_&" AND (T.[ErrorLevel]='"&ELV_&"') "
	end if
	if PROP_<>-1 then
	  SQL_=SQL_&" AND (T.[Property]='"&PROP_&"') "
	end if
	SQL_=SQL_&"ORDER BY 1 DESC"
	Rs.Open SQL_, Conn
	tblAuto=""
	rr=0
	if not Rs.Eof then
		DT_FILE=Rs.Fields(0)
		do while not Rs.Eof
			tblAuto=tblAuto&"<tr id='r"&rr&"'"
			tblAuto=tblAuto&" onclick=""javascript: changeColor("&rr&", "&Rs.Fields(6)&");"">"
			tblAuto=tblAuto&"<td "
			if DateDiff("n", Rs.Fields(0), Now())<=Rs.Fields(7) and Rs.Fields(4)="3" then
			   tblAuto=tblAuto&"style='border-bottom: solid 1px "&Color6aw&";'"
			end if
			tblAuto=tblAuto&">"&DateTimeFormat(Rs.Fields(0), "dd.mm.yy hh:nn")&"</td>"
			
			for i=1 to 5
			  tblAuto=tblAuto&"<td "
			  if DateDiff("n", Rs.Fields(0), Now())<=Rs.Fields(7) and Rs.Fields(4)="3" then
			    tblAuto=tblAuto&"style='border-bottom: solid 1px "&Color6aw&";"
				if i=5 then 
				tblAuto=tblAuto&"border-right: solid 1px "&Color6aw&";"'  "id=123456789 style='background-color:"&Color6aw&"; color: black;'"
				
				if Rs.Fields(0)=Rs.Fields(9) and Rs.Fields(8)<>0 then 'Условие для закраски в желлтый цвет
				tblAuto=tblAuto&" background-color:"&Color6aw&"; color: black; "
				end if
				
				end if
				tblAuto=tblAuto&"' "
			  else
			    if i=5 and Rs.Fields(0)=Rs.Fields(9) and Rs.Fields(8)<>0 then 'Условие для закраски в желлтый цвет
				tblAuto=tblAuto&" style='background-color:"&Color6aw&"; color: black;' "
				end if
			  end if
			  tblAuto=tblAuto&">"&Rs.Fields(i)&"</td>"
			next
			tblAuto=tblAuto&"</tr>"&vbCrLf
			rr=rr+1
			Rs.MoveNext
		loop
	end if
	Rs.Close
end if
%>
<!DOCTYPE HTML>
<html>
<head>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
		<!-- 1. Add these JavaScript inclusions in the head of your page -->
		<script type="text/javascript" src="js/jquery.min.js"></script>
		<script type="text/javascript" src="js/highcharts.js"></script>
		<script type="text/javascript" src="js/themes/gray.js"></script>
		
		<!-- 2. Add the JavaScript to initialize the chart on document ready -->
<% if T=0 then %>
		<script type="text/javascript">
		    var jsDate = 0;
		    var jsTime = '<% =right(DT_FILE, 8) %>';
			var jsChannel;
			var chartD1;
			var chartD2;
			var chartD3;
			var chartE;
			var chartF;
			
function ChangeDate(ofs) {
	jsDate=jsDate+ofs;
	if (jsDate>0) jsDate=0;
	graph(jsChannel, jsDate);

	var d = new Date();
	d.setDate(d.getDate()+jsDate);
	dd=((String(d.getDate()).length == 1) ? "0" + d.getDate() : d.getDate());
	mm=d.getMonth()+1;
	mm=((String(mm).length == 1) ? "0" + mm : mm);
	var t = dd+"."+mm +"."+d.getFullYear();
	var t2 = d.getFullYear()+"-"+mm+"-"+dd+" "+jsTime;
	diagr1(t2);
    $('#idDate').html(t);
	jQuery.get('dataset.asp', { ds: 'OperationsAtTime', prm: t2 }, function(ds) {
	  var tbl;
	  tbl='<table cellpadding="0" cellspacing="0" width="600px" style="font-size: 10pt; table-layout: fixed">'+
	  '<colgroup><col width="30px"><col width="160px"><col width="37px"><col width="30px"><col width="49px"><col width="289px"></colgroup>'+
	  '<tbody>'+ds+'</tbody></table>';
	  $('#Table1').html(tbl);
	});
};

function diagr1(dt) {
	jQuery.get('dataset.asp', { ds: 'OperationsDiagramm11', tag: 11, prm: dt }, function(ds) {
		//options_D1.series[0].data.length=0;
		//options_D1.series[1].data.length=0;
		//ds = ds.split('~');
		//var row = [];
		//
		//	row = ds[0].split(',');
		//	for (var i=0; i<row.length; i++) {
		//		if (i%2!=0) {
		//			options_D1.series[0].data.push([row[i-1], row[i]*1 ]);
		//		}
		//	}
		//
		//	//var obj = jQuery.parseJSON('[{"name": "", "y": "36.52", "color": "#00CC00"}, {"name": "?:<br />16", "y": "0.13", "color": "#FF3300"}, {"name": "", "y": "6.73", "color": "#00CC00"}, {"name": "Mastercard:<br />2", "y": "0.02", "color": "#FF3300"}, {"name": "", "y": "47.20", "color": "#00CC00"}, {"name": "Our ATM:<br />2", "y": "0.02", "color": "#FF3300"}, {"name": "", "y": "9.40", "color": "#00CC00"}, {"name": "Our POS:<br />0", "y": "0.00", "color": "#FF3300"}]');
		//	//var obj = jQuery.parseJSON(ds[1]);
		//	//options_D1.series[1].data = obj;
		//	
		//	row = ds[1].split('|||');
		//	for (var i=0; i<row.length; i++) {
		//		var v = {};
		//		var pnt = [];
		//		pnt=row[i].split(',');
		//		v.name=pnt[0];
		//		v.color=pnt[2];
		//		v.y=pnt[1]*1;
		//		options_D1.series[1].data.push(v);
		//	}
		//
		var obj = eval("[" + ds + "]");
		options_D1.series[0].data = obj;
		options_D1.chart.renderTo='container1';
		chartD1 = new Highcharts.Chart(options_D1);
	});
	
	jQuery.get('dataset.asp', { ds: 'OperationsDiagramm1', tag: 2, prm: dt }, function(ds) {
		options_D2.series[0].data.length=0;
		var row = [];
		row = ds.split(',');
		for (var i=0; i<row.length; i++) {
			if (i%2!=0) {
				options_D2.series[0].data.push([row[i-1], row[i]*1 ]);
			}
		}
		options_D2.chart.renderTo='container2';
		chartD2 = new Highcharts.Chart(options_D2);
	});

	jQuery.get('dataset.asp', { ds: 'OperationsDiagramm1', tag: 3, prm: dt }, function(ds) {
		options_D3.series[0].data.length=0;
		var row = [];
		row = ds.split(',');
		for (var i=0; i<row.length; i++) {
			if (i%2!=0) {
				options_D3.series[0].data.push([row[i-1], row[i]*1 ]);
			}
		}
		options_D3.chart.renderTo='container3';
		chartD3 = new Highcharts.Chart(options_D3);
	});
};

function graph(o, d) {
	jsChannel=o;
	jQuery.get('dataset.asp', { ds: 'OperationsHistory', tag: o, prm: 'I', prm2: d }, function(ds) {
		ds = ds.split('~');
		var ymax=[];
		for (var k=0; k<3; k++) {
			var d = [], row = [];
			row = ds[k].split(',');
			ymax[k]=0;
			for (var i=0; i<row.length; i++) {
				if (i%2==0) {
					date = Date.parse(row[i]);
				} else {
					if (ymax[k]*1<row[i]*1) {
					  ymax[k]=row[i]*1;
					}
					d.push([date, row[i]*1 ]);
				}
			}
			options.series[k].data = d;
		}
		options.title.text = o+': вход';
		options.chart.renderTo='containerE';
		chartE = new Highcharts.Chart(options);
	});
	
	jQuery.get('dataset.asp', { ds: 'OperationsHistory', tag: o, prm: 'O', prm2: d }, function(ds) {
		ds = ds.split('~');
		var ymax=[];
		for (var k=0; k<3; k++) {
			var d = [], row = [];
			row = ds[k].split(',');
			ymax[k]=0;
			for (var i=0; i<row.length; i++) {
				if (i%2==0) {
					date = Date.parse(row[i]);
				} else {
					if (ymax[k]*1<row[i]*1) {
					  ymax[k]=row[i]*1;
					}
					d.push([date, row[i]*1 ]);
				}
			}
			options.series[k].data = d;
		}
		options.title.text = o+': выход';
		options.chart.renderTo='containerF';
		chartF = new Highcharts.Chart(options);
	});
};

			// Настройки для второго и третьего графиков
			//chartE = new Highcharts.Chart({
			var options = {
				chart: {
					zoomType: 'x',
					marginLeft: 90,
					marginRight: 60
				},
				colors: ['#00FF00', '#FF0000'],
				credits: {enabled: false},
				legend:  {enabled: false},
				title:   {align: 'center', text: 'A1: in'},
				xAxis: {
					//max: Date.UTC(<% =CurrentTime %>),
					type: 'datetime',
					dateTimeLabelFormats: { hour: '%H:%M' }
				},
				tooltip: {
					shared: true,
					crosshairs: true,
					formatter: function() {
						var s = Highcharts.dateFormat('%d.%m.%y %H:%M', this.x);
                        $.each(this.points, function(i, point) { if (point.series.name != 'None') {
							s += '<br/><br/><span style="font-weight: 700; color: '+point.series.color+'">'+ point.series.name +': '+  point.y+'</span>'}});
						return s;
					}
				},
				yAxis: [{
					min: 0,
					title: {text: null},
					allowDecimals: false,
					lineColor: '#00FF00',
					labels: {
						style: {color: '#00FF00'}
					},
					plotLines: [{
						value: 0,
						width: 1,
						color: '#808080'
					}]
				}
				,{
					min: 0,
					title: {text: null},
					allowDecimals: false,
					opposite: true,
					lineColor: '#FF0000',
					labels: {
						style: {color: '#FF0000'}
					},
					plotLines: [{
						value: 0,
						width: 1,
						color: '#808080'
					}]
				}
				],
				plotOptions: {
					scatter: {
						marker: {enabled: false}
					},
					line: {
						animation: false,
						cursor: 'pointer',
						point: {
							events: {
								click: function() {
									var s = Highcharts.dateFormat('%Y-%m-%d %H:%M', this.x);
									diagr1(s);
									jQuery.get('dataset.asp', { ds: 'OperationsAtTime', prm: s }, function(ds) {
									  var tbl;
									  tbl='<table cellpadding="0" cellspacing="0" width="600px" style="font-size: 10pt; table-layout: fixed">'+
									  '<colgroup><col width="30px"><col width="160px"><col width="37px"><col width="30px"><col width="49px"><col width="289px"></colgroup>'+
									  '<tbody>'+ds+'</tbody></table>';
									  $('#Table1').html(tbl);
									  $('#idDate').html(s.substring(0,10));
									  jsTime=s.substring(11,18);
									  $('#idTime').html(jsTime);
									});
								}
							}
						}
					}
				},
				series: [{
					name: 'Успешных',
					type: 'line'
				}
				, {
					yAxis: 1,
					name: 'Сбойных',
					type: 'line'
				}, {
					name: 'None',
					type: 'scatter'
				}
				]
			};

			var options_D1 = {
				chart: {
					renderTo: 'container1',
					type: 'pie'
				},
				colors: ['#565656', '#669900', '#003399', '#9900CC', '#993333', '#006666'],
				title:   {align: 'left', text: 'Операции по каналам'},
				credits: {enabled: false},
				tooltip: {enabled: false},
				plotOptions: {
					series: {
						animation: false
					},
					pie: {
						animation: false,
						enableMouseTracking: false,
						shadow: false,
						size: '60%'
					}
				},
				series: [{
					data: [<% =series1 %>],
					dataLabels: {
						color: 'white',
						style: { font: 'bold 12px Arial' },
						formatter: function() {
							return this.point.name != '' ? this.point.name + this.y +'%' : null;
						}
					}
				}]
			};
			
			var options_D2 = {
				chart: {
					renderTo: 'container2',
					type: 'pie'
				},
				colors: ['#565656', '#669900', '#003399', '#9900CC', '#993333', '#006666'],
				title:   {align: 'left', text: 'Критичные RC по каналам'},
				credits: {enabled: false},
				tooltip: {enabled: false},
				plotOptions: {
					series: {
						animation: false
					},
					pie: {
						animation: false,
						enableMouseTracking: false,
						shadow: false,
						size: '60%'
					}
				},
				series: [{
					data: [<% =series2 %>],
					dataLabels: {
						color: 'white',
						style: { font: 'bold 12px Arial' },
						formatter: function() {
							return this.point.name != '' ? this.point.name + this.y +'%' : null;
						}
					}
				}]
			};

			var options_D3 = {
				chart: {
					renderTo: 'container3',
					type: 'pie'
				},
				colors: ['#565656', '#669900', '#003399', '#9900CC', '#993333', '#006666'],
				title:   {align: 'left', text: 'Критичные RC'},
				credits: {enabled: false},
				tooltip: {enabled: false},
				plotOptions: {
					series: {
						animation: false
					},
					pie: {
						animation: false,
						enableMouseTracking: false,
						shadow: false,
						size: '60%'
					}
				},
				series: [{
					data: [<% =series3 %>],
					dataLabels: {
						color: 'white',
						style: { font: 'bold 12px Arial' },
						formatter: function() {
							return this.point.name != '' ? 'RC'+this.point.name + this.y +'%'  : null;
						}
					}
				}]
			};

			// Первый график
			$(document).ready(function() {
			  chartD1 = new Highcharts.Chart(options_D1);
			  chartD2 = new Highcharts.Chart(options_D2);
			  chartD3 = new Highcharts.Chart(options_D3);
			  jsChannel = '<% =Channel1 %>';
			  graph(jsChannel, jsDate);
			});

		</script>
<% end if %>

<% if T=1 then %>
		<script type="text/javascript">
		    var jsDate = 0;
		    var jsTime = '<% =right(DT_FILE, 8) %>';
			var jsCHID = <% =CHID1 %>;
			var jsCHNM = '<% =CHNM1 %>';
			var chartG;
			var chartH;
			
			// Настройки графика с каналами
			var options = {
				chart: {
					type: 'line',
					zoomType: 'x'
				},
				colors: ['#F4F4F4'],
				credits: {enabled: false},
				legend:  {enabled: false},
				title:   {text: null},
				xAxis: {
					type: 'datetime',
					dateTimeLabelFormats: { hour: '%H:%M' }
				},
				plotOptions: {
					series: {
						animation: false
					}
				},
				tooltip: {
					shared: true,
					crosshairs: false,
					formatter: function() {
						var s = Highcharts.dateFormat('%d.%m.%y %H:%M', this.x);
                        $.each(this.points, function(i, point) {
							s += '<br/><span style="font-weight: 700; color: '+point.series.color+'">'+ point.series.name +': '+  point.y+'</span>'});
						return s;
					}
				},
					yAxis: {
					    min: 0,
					    max: 1,
						tickInterval: 1,
						title: { margin: 150, text: ' '},
						labels: { 
								formatter: function() {return ' '},
								style: {color: '#FFFFFF', font: 'normal 12px MS Sans Serif' },
								align: 'left',
								x: -150,
								y: -12
						},
						plotLines: [{
							value: 0,
							width: 1,
							color: '#808080'
						}]
					},
				series: [{name: 'Состояние', marker: {enabled: true, fillColor: '#CCCCCC', lineColor: '#FFFFFF', radius: 2} }]
			};
			
			var options2 = {
					chart: {
						renderTo: 'containerH',
						zoomType: 'x' },
					credits: {enabled: false},
					legend:  {enabled: false},
					tooltip: {enabled: false},
					title:   {text: null},
					xAxis: {
						type: 'datetime',
						dateTimeLabelFormats: { hour: '%H:%M' }
					},
					yAxis: {
					    min: 0,
					    max: 9,
						tickInterval: 1,
						title: { margin: 150, text: ' '},
						categories: [
							'<% =CID(8)&" "&CNM(8) %>',
							'<% =CID(7)&" "&CNM(7) %>',
							'<% =CID(6)&" "&CNM(6) %>',
							'<% =CID(5)&" "&CNM(5) %>',
							'<% =CID(4)&" "&CNM(4) %>',
							'<% =CID(3)&" "&CNM(3) %>',
							'<% =CID(2)&" "&CNM(2) %>',
							'<% =CID(1)&" "&CNM(1) %>',
							'Инф.'
						],
						labels: { 
								style: {color: '#FFFFFF', font: 'normal 8px MS Sans Serif' },
								align: 'left',
								x: -150,
								y: -6
						},
						plotLines: [{
							value: 0,
							width: 1,
							color: '#808080'
						}]
					},
					plotOptions: {
						series: {
							animation: false,
							dataLabels: {
								enabled: true,
								align: 'center',
								formatter: function() { return this.point.name; }
							}
						}, 
						line: {
							dataLabels: {enabled: false, color: '#FFFF00', y: -8, style: { font: 'normal 10px Arial' }},
							lineWidth: 0,
							marker: { enabled: false },
							enableMouseTracking: false
						},
						scatter: {
							dataLabels: {
								enabled: false,
								align: 'right',
								style: { font: 'normal 8px Arial' },
								formatter: function() {	return this.point.name; }
							},
							color: null, 
							marker: {
								enabled: true, 
								symbol: 'circle',
								fillColor: '#FF0000', 
								lineColor: '#FF0000', 
								radius: 3
							},
							enableMouseTracking: false
						}
					},
					series: [
					{
						name: 'Динамика состояния',
						type: 'scatter',
						data: [<% =series1 %>]
					},
					{
						name: 'Кол-во без связи',
						type: 'line',
						data: [<% =series2 %>]
					}
					<% =AllSeries %>
					]
				};

			$(document).ready(function() {
			  chartH = new Highcharts.Chart(options2);
			});
			
function ChGraph(chid, chname, chdt) {
	jsCHID=chid;
	jsCHNM=chname;
	jQuery.get('dataset.asp', { ds: 'ChannelHistory', tag: chid, prm: chname, prm2: chdt }, function(ds) {
			var d = [], row = [];
			row = ds.split(',');
			for (var i=0; i<row.length; i++) {
				if (i%2==0) {
					date = Date.parse(row[i]);
				} else {
				    v=row[i]*1;
					d.push([date, v ]);
				}
			}
			options.series[0].data = d;
		options.chart.renderTo='containerG';
		if (window.chartG !== undefined) { chartG.destroy() }
		chartG = new Highcharts.Chart(options);
		chartG.renderer.text(chid+' '+chname, 10, 48)
			.css({ color: '#FFFFFF', fontSize: '12px', fontName: 'MS Sans Serif' })
            .add();
	});
}

function ChTable(d) {
	for (var i=0; i<10; i++){
		if (options2.series[i]!=null) {
		  options2.series[i].data.length=0;
		}
	}
	if (window.chartH !== undefined) { chartH.destroy(); }
	
	jQuery.get('dataset.asp', { ds: 'ChannelHistory', tag: '~', prm: 'Table', prm2: d }, function(ds) {
		var part = [];
		part = ds.split('~');
		$('#idTable1').html(part[0]);
		var cat=jQuery.parseJSON(part[1]);
		options2.yAxis.categories=cat.categories;

		var obj0 = eval("[" + part[2] + "]");
		var obj1 = eval("[" + part[3] + "]");
		options2.series[0].data = obj0;	
		options2.series[1].data = obj1;
		for (var i=1; i<9; i++) {
		  options2.series[i+1].data = eval("[" + part[i+3] + "]");
		}

		//var opt2={chart: {renderTo: 'containerH'},series: [{name: 'Tokyo', data: [7.0, 9.9]}]};
		//var jsStr='{"chart": {"renderTo": "containerH"}, "series": [{"name": "Tokyo", "data": [7.0, 9.9]}]}';
		//var obj2=jQuery.parseJSON(jsStr);
		//chartH = new Highcharts.Chart(obj2);
		chartH = new Highcharts.Chart(options2);
	});
}

function ChangeDate(ofs) {
	jsDate=jsDate+ofs;
	if (jsDate>0) jsDate=0;
	if (ofs != 0) {
	  ChGraph(jsCHID, jsCHNM, jsDate);
	}

	var d = new Date();
	d.setDate(d.getDate()+jsDate);
	dd=((String(d.getDate()).length == 1) ? "0" + d.getDate() : d.getDate());
	mm=d.getMonth()+1;
	mm=((String(mm).length == 1) ? "0" + mm : mm);
	var t = dd+"."+mm +"."+d.getFullYear();
	var t2 = d.getFullYear()+"-"+mm+"-"+dd+" "+jsTime;
	$('#idTable1').html='';
	ChTable(jsDate);
    $('#idDate').html(t);
    $('#idTime').html('');
};

		</script>
<% end if %>		

<% if T=2 then %>
		<script type="text/javascript">
		    var jsDate = 0;
		    var jsTime = '<% =right(DT_FILE, 8) %>';
			var jsBC = '<% =BRANCH_CODE %>';
			var jsBN = '<% =BRANCH_NAME %>';
			var chartE;
			
function ChangeDate(ofs) {
	jsDate=jsDate+ofs;
	if (jsDate>0) jsDate=0;
	if (ofs != 0) {
		ChGraph(jsBC, jsBN, jsDate)
	}

	var d = new Date();
	d.setDate(d.getDate()+jsDate);
	dd=((String(d.getDate()).length == 1) ? "0" + d.getDate() : d.getDate());
	mm=d.getMonth()+1;
	mm=((String(mm).length == 1) ? "0" + mm : mm);
	var t = dd+"."+mm +"."+d.getFullYear();
	var t2 = d.getFullYear()+"-"+mm+"-"+dd+" "+jsTime;
	$('#idTable1').html='';
	ChTable(t2);
    $('#idDate').html(t);
};

function ChGraph(bc, bn, d) {
	jsBC=bc;
	jsBN=bn;
	jQuery.get('dataset.asp', { ds: 'AtmNoLink', tag: bc, prm: 'Graph', prm2: d }, function(ds) {
		dts = ds.split('~');
		options.series[0].data.length=0;
		options.series[1].data.length=0;
		options.series[2].data.length=0;
		options.series[3].data.length=0;
		var obj1 = eval("[" + dts[0] + "]");
		var obj2 = eval("[" + dts[1] + "]");
		var obj3 = eval("[" + dts[2] + "]");
		var obj4 = eval("[" + dts[3] + "]");
		options.series[0].data = obj1;
		options.series[0].name = 'Неработоспособность '+bn+' (%)';
		options.series[1].data = obj2;
		options.series[2].data = obj3;
		options.series[3].data = obj4;
		options.series[3].name = 'Без связи за сутки  '+bn+' (%)' ;
		options.title.text = bn;
		options.chart.renderTo='containerE';
		chartE = new Highcharts.Chart(options);
		chartE.yAxis[0].addPlotLine({ value: <% = Main2A24_SetHiHi %>, color: '#d7b6b6', dashStyle: 'Dash', width: 2, id: 'plot-line-1' });
	});
}
function ChTable(d) {
	jQuery.get('dataset.asp', { ds: 'AtmNoLink', prm: 'Table', prm2: d }, function(ds) {
		$('#idTable1').html(ds);
	});
}

			var options = {
				chart: {
					type: 'line',
					zoomType: 'x'
				},
				colors: ['#F0F0F0', '#66FFFF', '#FF66FF', '#d7b6b6'],
				credits: {enabled: false},
				legend:  {enabled: false},
				title:   {align: 'center', text: 'A1: in'},
				xAxis: {
					//max: Date.UTC(<% =CurrentTime %>),
					type: 'datetime',
					dateTimeLabelFormats: { hour: '%H:%M' }
				},
				tooltip: {
					shared: true,
					crosshairs: true,
					formatter: function() {
						var s = Highcharts.dateFormat('%d.%m.%y %H:%M', this.x);
                        $.each(this.points, function(i, point) {
							s += '<br/><br/><span style="font-weight: 700; color: '+point.series.color+'">'+ point.series.name +': '+  point.y+'</span>'});
						return s;
					}
				},
				yAxis: [{
					min: 0,
					title: {text: null},
					allowDecimals: false,
					lineColor: '#FFFFFF',
					labels: {
						style: {color: '#FFFFFF'}
					},
					plotLines: [{
						value: 0,
						width: 1,
						color: '#808080'
					}]
				}
				,{
					min: 0,
					title: {text: null},
					allowDecimals: false,
					opposite: true,
					lineColor: '#FF66FF',
					labels: {
						style: {color: '#FF66FF'}
					},
					plotLines: [{
						value: 0,
						width: 1,
						color: '#808080'
					}]
				}
				],
				plotOptions: {
					series: {
						animation: false,
						cursor: 'pointer',
						point: {
							events: {
								click: function() {
									var s = Highcharts.dateFormat('%Y-%m-%d %H:%M', this.x);
									ChTable(s);
									jsTime=s.substring(11,18)
									$('#idTime').html('&nbsp;'+jsTime);
									ChangeDate(0);
								}
							}
						}							
					}
				},
				series: [
					{
						name: 'По выбранному ФИ',
						type: 'line',
						data: []
					},
					{
						name: 'По процессингу (%)',
						type: 'line',
						data: []
					},
					{
						name: 'Филиалов без связи (шт)',
						type: 'line',
						yAxis: 1,
						data: []
					},
					{
						name: 'Без связи за сутки (%)',
						type: 'line',
						data: []
					}
				]
			};

			// Первый график
			$(document).ready(function() {
				ChGraph(jsBC, jsBN, jsDate);
			});

		</script>
<% end if %>

<% if T=3 then %>
		<script type="text/javascript">
		    var jsDate = 0;
		    var jsTime = '<% =right(DT_FILE, 8) %>';
			var jsLT = '<% =LINK_TYPE %>';
			var chartE;
			
function ChangeDate(ofs) {
	jsDate=jsDate+ofs;
	if (jsDate>0) jsDate=0;
	if (ofs != 0) {
		ChGraph(jsLT, jsDate)
	}

	var d = new Date();
	d.setDate(d.getDate()+jsDate);
	dd=((String(d.getDate()).length == 1) ? "0" + d.getDate() : d.getDate());
	mm=d.getMonth()+1;
	mm=((String(mm).length == 1) ? "0" + mm : mm);
	var t = dd+"."+mm +"."+d.getFullYear();
	var t2 = d.getFullYear()+"-"+mm+"-"+dd+" "+jsTime;
	$('#idTable1').html='';
	ChTable(t2);
    $('#idDate').html(t);
};

function ChGraph(lt, d) {
	jsLT=lt;
	jQuery.get('dataset.asp', { ds: 'AtmTypeLink', tag: lt, prm: 'Graph', prm2: d }, function(ds) {
		dts = ds.split('~');
		options.series[0].data.length=0;
		options.series[1].data.length=0;
		var obj1 = eval("[" + dts[0] + "]");
		var obj2 = eval("[" + dts[1] + "]");
		options.series[0].data = obj1;
		options.series[0].name = lt+' (%)';
		options.series[1].data = obj2;
		options.title.text = lt;
		options.chart.renderTo='containerE';
		chartE = new Highcharts.Chart(options);
		chartE.yAxis[0].addPlotLine({ value: <% = Main2Centr_SetHiHi %>, color: '#66FFFF', dashStyle: 'Dash', width: 2, id: 'plot-line-1' });
	});
}
function ChTable(d) {
	jQuery.get('dataset.asp', { ds: 'AtmTypeLink', prm: 'Table', prm2: d }, function(ds) {
		$('#idTable1').html(ds);
	});
}

			var options = {
				chart: {
					type: 'line',
					zoomType: 'x'
				},
				colors: ['#F0F0F0', '#66FFFF', '#FF66FF'],
				credits: {enabled: false},
				legend:  {enabled: false},
				title:   {align: 'center', text: 'A1: in'},
				xAxis: {
					//max: Date.UTC(<% =CurrentTime %>),
					type: 'datetime',
					dateTimeLabelFormats: { hour: '%H:%M' }
				},
				tooltip: {
					shared: true,
					crosshairs: true,
					formatter: function() {
						var s = Highcharts.dateFormat('%d.%m.%y %H:%M', this.x);
                        $.each(this.points, function(i, point) {
							s += '<br/><br/><span style="font-weight: 700; color: '+point.series.color+'">'+ point.series.name +': '+  point.y+'</span>'});
						return s;
					}
				},
				yAxis: [{
					min: 0,
					title: {text: null},
					allowDecimals: false,
					lineColor: '#FFFFFF',
					labels: {
						style: {color: '#FFFFFF'}
					},
					plotLines: [{
						value: 0,
						width: 1,
						color: '#808080'
					}]
				}
				,{
					min: 0,
					title: {text: null},
					allowDecimals: false,
					opposite: true,
					lineColor: '#FF66FF',
					labels: {
						style: {color: '#FF66FF'}
					},
					plotLines: [{
						value: 0,
						width: 1,
						color: '#808080'
					}]
				}
				],
				plotOptions: {
					series: {
						animation: false,
						cursor: 'pointer',
						point: {
							events: {
								click: function() {
									var s = Highcharts.dateFormat('%Y-%m-%d %H:%M', this.x);
									ChTable(s);
									jsTime=s.substring(11,18)
									$('#idTime').html('&nbsp;'+jsTime);
									ChangeDate(0);
								}
							}
						}							
					}
				},
				series: [
					{
						name: 'По выбранному ФИ',
						type: 'line',
						data: []
					},
					{
						name: 'По процессингу (%)',
						type: 'line',
						data: []
					}
				]
			};

			// Первый график
			$(document).ready(function() {
				ChGraph(jsLT, jsDate);
			});

		</script>
<% end if %>

<% if T=4 then %>
		<script type="text/javascript">
		    var jsDate = 0;
		    var jsTime = '<% =right(DT_FILE, 8) %>';
			var jsBC = '<% =BRANCH_CODE %>';
			var jsBN = '<% =BRANCH_NAME %>';
			var chartE;
			
function ChangeDate(ofs) {
	jsDate=jsDate+ofs;
	if (jsDate>0) jsDate=0;
	if (ofs != 0) {
		ChGraph(jsBC, jsBN, jsDate)
	}

	var d = new Date();
	d.setDate(d.getDate()+jsDate);
	dd=((String(d.getDate()).length == 1) ? "0" + d.getDate() : d.getDate());
	mm=d.getMonth()+1;
	mm=((String(mm).length == 1) ? "0" + mm : mm);
	var t = dd+"."+mm +"."+d.getFullYear();
	var t2 = d.getFullYear()+"-"+mm+"-"+dd+" "+jsTime;
	$('#idTable1').html='';
	ChTable(t2);
    $('#idDate').html(t);
};

function ChGraph(bc, bn, d) {
	jsBC=bc;
	jsBN=bn;
	jQuery.get('dataset.asp', { ds: 'BPTNoLink', tag: bc, prm: 'Graph', prm2: d }, function(ds) {
		dts = ds.split('~');
		options.series[0].data.length=0;
		options.series[1].data.length=0;
		options.series[2].data.length=0;
		options.series[3].data.length=0;
		var obj1 = eval("[" + dts[0] + "]");
		var obj2 = eval("[" + dts[1] + "]");
		var obj3 = eval("[" + dts[2] + "]");
		var obj4 = eval("[" + dts[3] + "]");
		options.series[0].data = obj1;
		options.series[0].name = 'Неработоспособность  '+bn+' (%)';
		options.series[1].data = obj2;
		options.series[2].data = obj3;
		options.series[3].data = obj4;
		options.series[3].name = 'LINK 24 '+bn+' (%)' ;
		options.title.text = bn;
		options.chart.renderTo='containerE';
		chartE = new Highcharts.Chart(options);
		chartE.yAxis[0].addPlotLine({ value: <% = Main2AllBPT24_SetHiHi %>, color: '#d7b6b6', dashStyle: 'Dash', width: 2, id: 'plot-line-1' });
	});
}
function ChTable(d) {
	jQuery.get('dataset.asp', { ds: 'BPTNoLink', prm: 'Table', prm2: d }, function(ds) {
		$('#idTable1').html(ds);
	});
}

			var options = {
				chart: {
					type: 'line',
					zoomType: 'x'
				},
				colors: ['#F0F0F0', '#66FFFF', '#FF66FF', '#d7b6b6'],
				credits: {enabled: false},
				legend:  {enabled: false},
				title:   {align: 'center', text: 'A1: in'},
				xAxis: {
					//max: Date.UTC(<% =CurrentTime %>),
					type: 'datetime',
					dateTimeLabelFormats: { hour: '%H:%M' }
				},
				tooltip: {
					shared: true,
					crosshairs: true,
					formatter: function() {
						var s = Highcharts.dateFormat('%d.%m.%y %H:%M', this.x);
                        $.each(this.points, function(i, point) {
							s += '<br/><br/><span style="font-weight: 700; color: '+point.series.color+'">'+ point.series.name +': '+  point.y+'</span>'});
						return s;
					}
				},
				yAxis: [{
					min: 0,
					title: {text: null},
					allowDecimals: false,
					lineColor: '#FFFFFF',
					labels: {
						style: {color: '#FFFFFF'}
					},
					plotLines: [{
						value: 0,
						width: 1,
						color: '#808080'
					}]
				}
				,{
					min: 0,
					title: {text: null},
					allowDecimals: false,
					opposite: true,
					lineColor: '#FF66FF',
					labels: {
						style: {color: '#FF66FF'}
					},
					plotLines: [{
						value: 0,
						width: 1,
						color: '#808080'
					}]
				}
				],
				plotOptions: {
					series: {
						animation: false,
						cursor: 'pointer',
						point: {
							events: {
								click: function() {
									var s = Highcharts.dateFormat('%Y-%m-%d %H:%M', this.x);
									ChTable(s);
									jsTime=s.substring(11,18)
									$('#idTime').html('&nbsp;'+jsTime);
									ChangeDate(0);
								}
							}
						}							
					}
				},
				series: [
					{
						name: 'По выбранному ФИ',
						type: 'line',
						data: []
					},
					{
						name: 'По процессингу (%)',
						type: 'line',
						data: []
					},
					{
						name: 'Без связи за сутки (шт)',
						type: 'line',
						yAxis: 1,
						data: []
					},
					{
						name: 'LINK 24 (%)',
						type: 'line',
						data: []
					}
				]
			};

			// Первый график
			$(document).ready(function() {
				ChGraph(jsBC, jsBN, jsDate);
			});

		</script>
<% end if %>

<% if T=5 then %>
		<script type="text/javascript">
		    var jsDate = 0;
			<%=SMSServiceStringchrt%>
			
function ChangeDate(ofs) {
	jsDate=jsDate+ofs;
	if (jsDate>0) jsDate=0;
	if (ofs != 0) {
		ChGraph(jsDate)
	}

	var d = new Date();
	d.setDate(d.getDate()+jsDate);
	dd=((String(d.getDate()).length == 1) ? "0" + d.getDate() : d.getDate());
	mm=d.getMonth()+1;
	mm=((String(mm).length == 1) ? "0" + mm : mm);
	var t = dd+"."+mm +"."+d.getFullYear();
	var t2 = d.getFullYear()+"-"+mm+"-"+dd;
    $('#idDate').html(t);
};

function ChGraph(d) {
	jQuery.get('dataset.asp', { ds: 'SMSService', prm2: d }, function(ds) {
		dts = ds.split('~');
		<%=SMSServiceString%>
	});
}

			var options0 = {
				chart: {
					type: 'line',
					zoomType: 'x'
				},
				colors: ['#F0F0F0', '#66FFFF', '#FF66FF'],
				credits: {enabled: false},
				legend:  {enabled: false},
				title:   {align: 'center', text: 'A1: in'},
				xAxis: {
					type: 'datetime',
					dateTimeLabelFormats: { hour: '%H:%M' }
				},
				tooltip: {
					shared: true,
					crosshairs: true,
					formatter: function() {
						var s = Highcharts.dateFormat('%d.%m.%y %H:%M', this.x);
                        $.each(this.points, function(i, point) {
							s += '<br/><br/><span style="font-weight: 700; color: '+point.series.color+'">'+ point.series.name +': '+  point.y+'</span>'});
						return s;
					}
				},
				yAxis: [{
					min: 0,
					title: {text: null},
					allowDecimals: false,
					lineColor: '#FFFFFF',
					labels: {
						style: {color: '#FFFFFF'}
					},
					plotLines: [{
						value: 0,
						width: 1,
						color: '#808080'
					}]
				}
				],
				plotOptions: {
					series: {
						animation: false,
						cursor: 'pointer'
					}
				},
				series: [
					{name: 'В очереди', type: 'line', data: []},
					{name: 'Отклонено процессингом (%)', type: 'line', data: []},
					{name: 'Отклонено провайдером (%)', type: 'line', data: []}
				]
			};
			<%=SMSServiceStringoption%>
			// Первый график
			$(document).ready(function() {
				ChGraph(jsDate);
			});

		</script>
		<script type="text/javascript">
		$(document).ready(function() {
$("table.SMSService tr").click(
		function(){
//var i = 1;
		id =$(this).attr("id");
	//	alert (id);
		id="container"+id
		Showptions(id);
		})
});
function Showptions(gname)
{
  for (var i = 0; i < <%=SMScount%>; i++)
  {
    var name="container"+i;
    if (name==gname)
	{
	if (document.getElementById(name).style.display=="none")
		{
		document.getElementById(name).style.display="";
		}
	}
	else
	{
	document.getElementById(name).style.display="none";
	}
  }
}
</script>
<% end if %>

<% if T=6 then %>
<link href="js/calendar.css" rel="stylesheet" type="text/css" />
<script src="js/calendar.js" type="text/javascript"></script>
<script src="js/calendar-ru.js" type="text/javascript"></script>

		<script type="text/javascript">
function changeColor(row_ID, mesID)
{
  // обесцвечиваем старый
  var oldRow;
  oldRow = document.getElementById("buffercolor").value;
  document.getElementById("r"+oldRow).style.backgroundColor="#000000";
 
  // подсвечиваем новый
  document.getElementById("r"+row_ID).style.backgroundColor="#808080";
  document.getElementById("buffercolor").value = row_ID;

  jQuery.get('dataset.asp', { ds: 'MessageTypeByID', prm: mesID }, function(ds) {
	  $('#inf').html(ds);
  });
}		
function selected(cal, date) {
  cal.sel.value = date;
}
function closeHandler(cal) {
  cal.hide();
}
function showCalendar(id, format) {
  var el = document.getElementById(id);
  if (calendar != null) {
    calendar.hide();
  } else {
    var cal = new Calendar(false, null, selected, closeHandler);
    cal.weekNumbers = false;
    calendar = cal;                  // remember it in the global var
    cal.setRange(1930, 2030);        // min/max year allowed.
	cal.mondayFirst = true;
    cal.create();
  }
  calendar.setDateFormat(format);    // set the specified date format
  calendar.parseDate(el.value);      // try to parse the text in field
  calendar.sel = el;                 // inform it what input field we use
  calendar.showAtElement(el);        // show the calendar below it
  return false;
}

function SaveChange() {
  var e=document.getElementById('FCAT_');
  if (e) {
    document.getElementById('CAT_').value=e.value;
    document.getElementById('PRI_').value=document.getElementById('FPRI_').value;
    document.getElementById('ELV_').value=document.getElementById('FELV_').value;
    document.getElementById('PROP_').value=document.getElementById('FPROP_').value;
  }
  document.getElementById("Form6").submit();
}

		</script>
<% end if %>

	<style type="text/css">
	<!--
BODY {
	margin: 0px;
	background-color: #242424;
}
TABLE {
	margin-bottom: 0px;
	margin-top: 0px;
	border-top: solid 1px #CCCCCC;
	border-left: solid 1px #CCCCCC;
}
TD {
	padding-top: 1px;
	padding-bottom: 1px;
	padding-left: 2px;
	padding-right: 2px;
	text-align: center;
	color: #FFFFFF;
	font-family: Verdana, Arial, helvetica, sans-serif, Geneva;
	border-bottom: solid 1px #CCCCCC;
	border-right: solid 1px #CCCCCC;
}
TD.Head {
	background-color: #6F8CBF;
	font-weight: 700;
	font-size: 8pt;
}
TD.wb {
	border: none;
}
TH.A {
	border-right: solid 1px black;
	border-bottom: solid 1px black;
	font-weight: 600;
	color: #FFFFFF;
	background-color: #6F8CBF;
	padding-left: 2px;
	padding-right: 2px;
	text-align: left;
}
TD.A {
	border: none;
	text-align: left;
}	-->
	</style>
</head>
<body>
<% if T=6 then %>
<form name="Form6" id="Form6" action="detail.asp" method="post" style="margin-bottom: 0px;">
<input type="hidden" id="buffercolor" value="0">
<input type="hidden" id="T" name="T" value="6">
<input type="hidden" id="CAT_" name="CAT_" value="<%=CAT_%>">
<input type="hidden" id="PRI_" name="PRI_" value="<%=PRI_%>">
<input type="hidden" id="ELV_" name="ELV_" value="<%=ELV_%>">
<input type="hidden" id="PROP_" name="PROP_" value="<%=PROP_%>">
<% end if %>
<div align="center">
<table border="0" cellspacing="2" cellpadding="0" height="24px" style="border: none; margin-bottom: 8px;">
<tr>
<td class="wb" style="cursor: hand; background-color: <% =Color1 %>; color: #000000" onclick="location.replace('detail.asp?T=0');">Операции</td>
<td class="wb" width="24px">&nbsp;</td>
<td class="wb" style="cursor: hand; background-color: <% =Color3 %>; color: #000000" onclick="location.replace('detail.asp?T=1');">Каналы</td>
<td class="wb" width="24px">&nbsp;</td>
<td class="wb" colspan="3" style="cursor: hand; background-color: <% =Color2 %>; color: #000000" onclick="location.replace('detail.asp?T=2');">&nbsp;&nbsp;Устройства&nbsp;&nbsp;</td>
<td class="wb" width="24px">&nbsp;</td>
<td class="wb" style="cursor: hand; background-color: <% =Color5 %>; color: #000000" onclick="location.replace('detail.asp?T=5');">&nbsp;&nbsp;SMS-сервис&nbsp;&nbsp;</td>
<td class="wb" width="24px">&nbsp;</td>
<td class="wb" style="cursor: hand; background-color: <% =Color6 %>; color: #000000" onclick="location.replace('detail.asp?T=6');">&nbsp;&nbsp;Автоматизация ПЦ&nbsp;&nbsp;</td>
<td class="wb" width="24px">&nbsp;</td>
</tr>
<tr>
<td class="wb"><span style="color: #FF0000"><% =Value1&" ("&FormatNumber(Value1proc, 2, -1, 0, 0)&"%)" %></span> / <% =Value1all %></td>
<td class="wb" width="24px">&nbsp;</td>
<td class="wb"><span style="color: #FF0000"><% =Value3&" ("&FormatNumber(Value3proc, 2, -1, 0, 0)&"%)" %></span> / <% =Value3total %></td>
<td class="wb" width="24px">&nbsp;</td>
<td class="wb" style="cursor: hand; border-right: solid 1px #CCC; color: <% =colorATM %>" onclick="location.replace('detail.asp?T=2');">АТМ</td>
<td class="wb" style="cursor: hand; border-right: solid 1px #CCC; color: <% =colorCSP %>" onclick="location.replace('detail.asp?T=3');">ЦСП</td>
<td class="wb" style="cursor: hand; color: <% =colorBPT %>" onclick="location.replace('detail.asp?T=4');">БПТ</td>
<td class="wb" width="24px">&nbsp;</td>
<td class="wb" width="24px">&nbsp;</td>
</tr>
</table>
</div>
<%
' ФОРМИРОВАНИЕ СТРАНИЦЫ ДЛЯ ОТОБРАЖЕНИЯ ДЕТАЛЬНОЙ ИНФОРМАЦИИ ДЛЯ ОСНОВНОГО ПАРАМЕТРА №1 (ОПЕРАЦИИ)
if T=0 then
%>
<div id="container1"  style="width: 416px; height: 380px; margin: 0 auto; float: left;"></div>
<div id="container2"  style="width: 416px; height: 380px; margin: 0 auto; float: left;"></div>
<div id="container3"  style="width: 416px; height: 380px; margin: 0 auto; float: left;"></div>
<br>
<div align="center">
<table border="0" cellspacing="2" cellpadding="0" style="border: none;">
<tr>
<td style="border: none; text-decoration: underline; cursor: hand;" onclick="ChangeDate(-1)">&laquo;</td>
<td style="border: none; font-size: 10pt; font-weight: 700; text-align: center;" nowrap><div style="float: left;">Состояние на &nbsp;</div><div id="idDate" style="float: left;"><% =left(DT_FILE, 10) %></div><div id="idTime" style="float: left;">&nbsp;<% =right(DT_FILE, 8) %></div></td>
<td style="border: none; text-decoration: underline; cursor: hand;" onclick="ChangeDate(1)">&raquo;</td>
</tr>
</table>
</div> 
<div id="tabble1"  style="width: 600px; height: 360px; margin-top: 0px; float: left;">
		    <table cellpadding="0" cellspacing="0" width="600px" style="font-size: 8pt; table-layout: fixed">
			<colgroup><col width="30px"><col width="160px"><col width="37px"><col width="30px"><col width="49px"><col width="*"></colgroup>
			<tbody>
			<tr><td class="Head" height="20px">Код</td>
			<td class="Head" style="padding-left: 0px;">
				<select size="1" id="Ch" name="Ch" onchange="graph(this.value, jsDate)">
				<%
					Rs.Open "SELECT DISTINCT [NAME] FROM NV_Operations ORDER BY 1", Conn
					if not Rs.Eof then
					  do while not Rs.Eof
					    if Rs.Fields(0) = Channel1 then
					      Response.Write("<option value="""&Rs.Fields(0)&""" selected>"&Rs.Fields(0)&"</option>")
						else
					      Response.Write("<option value="""&Rs.Fields(0)&""">"&Rs.Fields(0)&"</option>")
						end if
					    Rs.MoveNext
					  loop
					end if
					Rs.Close
				%>			
				</select>
			</td><td class="Head">Напр</td><td class="Head">Код</td><td class="Head">Кол-во</td><td class="Head">Сообщение</span></td></tr>
			</tbody>
			</table>
			<div style="OVERFLOW-Y: auto; OVERFLOW-X: hidden; OVERFLOW: auto; width: 600px; height: 342px;" id="Table1">
			    <table cellpadding="0" cellspacing="0" width="600px" style="font-size: 10pt; table-layout: fixed">
				<colgroup><col width="30px"><col width="160px"><col width="37px"><col width="30px"><col width="49px"><col width="*"></colgroup>
				<tbody>
					<%
					Rs.Open "SELECT [CHANNEL], [NAME], [DIRECTION], [RESPONSE_CODE], [QUANTITY], C.Resp_text FROM NV_Operations AS O LEFT OUTER JOIN V_Resp_code AS C ON O.RESPONSE_CODE=C.Resp_code WHERE ISNULL(C.IsFailed, 0)<>0 ORDER BY 2, 3", Conn
					if not Rs.Eof then
					  do while not Rs.Eof
					    Response.Write("<tr><td>"&Rs.Fields(0)&"</td><td style='text-align: left;'>"&Rs.Fields(1)&"</td><td>"&Rs.Fields(2)&"</td><td>"&Rs.Fields(3)&"</td><td>"&Rs.Fields(4)&"</td><td style='text-align: left;'>"&Rs.Fields(5)&"</td></tr>")
					    Rs.MoveNext
					  loop
					else
					  Response.Write("<tr><td colspan=6>Нет каналов с неуспешными операциями</td></tr>")
					end if
					Rs.Close
					%>
				</tbody>
				</table>
			</div>
</div>
<div id="graph2"  style="width: 600px; height: 360px; margin-left: 10px; margin-top: 0px; float: left;">
		  <div id="containerE" style="width: 600px; height: 180px; margin: 0 auto"></div>
		  <div id="containerF" style="width: 600px; height: 180px; margin: 0 auto"></div>
</div>

<% end if %>
<%
' ФОРМИРОВАНИЕ СТРАНИЦЫ ДЛЯ ОТОБРАЖЕНИЯ ДЕТАЛЬНОЙ ИНФОРМАЦИИ ДЛЯ ОСНОВНОГО ПАРАМЕТРА №2 (КАНАЛЫ)
if T=1 then
%>
<div id="tabble1"  style="width: 670px; height: 250px; margin-top: 0px; margin-right: 10px; margin-left: 10px;">
		    <table cellpadding="0" cellspacing="0" width="670px" style="font-size: 8pt; table-layout: fixed">
			<colgroup>
				<col width="55px">
				<col width="200px">
				<col width="85px">
				<col width="165px">
				<col width="165px">
			</colgroup>
			<tbody>
				<tr>
					<td class="Head" height="20px">ID</td>
					<td class="Head">Канал</td>
					<td class="Head">Кол-во падений за сутки</td>
					<td class="Head">Последнее падение за сутки</td>
					<td class="Head">Состояние на</td>
				</tr>
			</tbody>
			</table>
			<div style="OVERFLOW-Y: auto; OVERFLOW-X: hidden; OVERFLOW: auto; width: 670px; height: 232px;" id="idTable1">
			    <table cellpadding="0" cellspacing="0" width="670px" style="font-size: 10pt; table-layout: fixed">
				<colgroup>
				<col width="55px">
				<col width="200px">
				<col width="85px">
				<col width="165px">
				<col width="165px">
				</colgroup>
				<tbody>
					<% =tblChannel %>
				</tbody>
				</table>
			</div>
</div>
<br />
<div align="center" style="margin-top: 8px">
<table border="0" cellspacing="2" cellpadding="0" style="border: none">
<tr>
<td style="border: none; text-decoration: underline; cursor: hand;" onclick="ChangeDate(-1)">&laquo;</td>
<td nowrap style="border: none; font-size: 10pt; font-weight: 700; text-align: center;"><div style="float: left;">Состояние на &nbsp;</div><div id="idDate" style="float: left;"><% =left(DT_FILE, 10) %></div><div id="idTime" style="float: left;">&nbsp;<% =right(DT_FILE, 8) %></div></td>
<td style="border: none; text-decoration: underline; cursor: hand;" onclick="ChangeDate(1)">&raquo;</td>
</tr>
</table>
<div id="containerH" style="width: 98%; height: 360px; margin: 0 auto"></div>
<div id="containerG" style="width: 98%; height: 100px; margin: 0 auto"></div>
</div> 
<% end if %>
<%
' ФОРМИРОВАНИЕ СТРАНИЦЫ ДЛЯ ОТОБРАЖЕНИЯ ДЕТАЛЬНОЙ ИНФОРМАЦИИ ДЛЯ ОСНОВНОГО ПАРАМЕТРА №3 (УСТРОЙСТВА)
' БАНКОМАТЫ БЕЗ СВЯЗИ
if T=2 then
%>
<div style="width: 100%; background-color: #363636; color: #FFFFFF; text-align: center; margin-top: 16px; margin-bottom: 16px; font-family: Verdana, Arial, helvetica, sans-serif, Geneva; font-size: 10pt; font-weight: 700; border: solid 1px #000080; ">
Детализация по связи для банкоматов по процессингу и финансовому институту
</div>
<div id="tabble1"  style="width: 800px; height: 248px; margin-top: 0px; margin-right: 10px; margin-left: 10px;">
		    <table cellpadding="0" cellspacing="0" width="800px" style="font-size: 8pt; table-layout: fixed">
				<colgroup><col span=6 width="58px"><col width="*"></colgroup>
				<tbody>
				<tr>
					<td class="Head" height="20px">BRANCH</td><td class="Head">ATM</td><td class="Head">LNK</td><td class="Head">ERR</td><td class="Head">LNK_ERR</td><td class="Head">LNK 24 %</td><td class="Head">NAME</td>
				</tr>
				</tbody>
			</table>
			<div style="OVERFLOW-Y: auto; OVERFLOW-X: hidden; OVERFLOW: auto; width: 800px; height: 230px;" id="idTable1">
			    <table cellpadding="0" cellspacing="0" width="800px" style="font-size: 10pt; table-layout: fixed">
					<colgroup><col span=6 width="58px"><col width="440px"></colgroup>
					<tbody>
					<% =tblAtmLink %>
					</tbody>
				</table>
			</div>
</div>
<div align="center" style="margin-top: 8px">
<table border="0" cellspacing="2" cellpadding="0" style="border: none;">
<tr>
<td style="border: none; text-decoration: underline; cursor: hand;" onclick="ChangeDate(-1)">&laquo;</td>
<td nowrap style="border: none; font-size: 10pt; font-weight: 700; text-align: center;"><div style="float: left;">Состояние на &nbsp;</div><div id="idDate" style="float: left;"><% =left(DT_FILE, 10) %></div><div id="idTime" style="float: left;">&nbsp;<% =right(DT_FILE, 8) %></div></td>
<td style="border: none; text-decoration: underline; cursor: hand;" onclick="ChangeDate(1)">&raquo;</td>
</tr>
</table>
<div id="containerE" style="width: 100%; height: 260px; margin: 0 auto"></div>
</div>

<% end if %>

<%
' ФОРМИРОВАНИЕ СТРАНИЦЫ ДЛЯ ОТОБРАЖЕНИЯ ДЕТАЛЬНОЙ ИНФОРМАЦИИ ДЛЯ ОСНОВНОГО ПАРАМЕТРА №3 (УСТРОЙСТВА)
' БАНКОМАТЫ БЕЗ СВЯЗИ ПО СХЕМАМ ПОДКЛЮЧЕНИЯ
if T=3 then
%>
<div style="width: 100%; background-color: #363636; color: #FFFFFF; text-align: center; margin-top: 16px; margin-bottom: 16px; font-family: Verdana, Arial, helvetica, sans-serif, Geneva; font-size: 10pt; font-weight: 700; border: solid 1px #000080; ">
Детализация по связи для банкоматов по схемам подключения
</div>
<div id="tabble1"  style="width: 800px; height: 248px; margin-top: 0px; margin-right: 10px; margin-left: 10px;">
		    <table cellpadding="0" cellspacing="0" width="800px" style="font-size: 8pt; table-layout: fixed">
				<colgroup><col width="250px"><col span=3 width="58px"><col width="*"></colgroup>
				<tbody>
				<tr>
					<td class="Head" height="20px">LINK_TYPE</td><td class="Head">ATM</td><td class="Head">OFF</td><td class="Head">ЦСП</td><td class="Head">OFF, %</td>
				</tr>
				</tbody>
			</table>
			<div style="OVERFLOW-Y: auto; OVERFLOW-X: hidden; OVERFLOW: auto; width: 800px; height: 230px;" id="idTable1">
			    <table cellpadding="0" cellspacing="0" width="800px" style="font-size: 10pt; table-layout: fixed">
					<colgroup><col width="250px"><col span=3 width="58px"><col width="*"></colgroup>
					<tbody>
					<% =tblAtmLink %>
					</tbody>
				</table>
			</div>
</div>
<div align="center" style="margin-top: 8px">
<table border="0" cellspacing="2" cellpadding="0" style="border: none;">
<tr>
<td style="border: none; text-decoration: underline; cursor: hand;" onclick="ChangeDate(-1)">&laquo;</td>
<td nowrap style="border: none; font-size: 10pt; font-weight: 700; text-align: center;"><div style="float: left;">Состояние на &nbsp;</div><div id="idDate" style="float: left;"><% =left(DT_FILE, 10) %></div><div id="idTime" style="float: left;">&nbsp;<% =right(DT_FILE, 8) %></div></td>
<td style="border: none; text-decoration: underline; cursor: hand;" onclick="ChangeDate(1)">&raquo;</td>
</tr>
</table>
<div id="containerE" style="width: 100%; height: 260px; margin: 0 auto"></div>
</div>
<% end if %>

<%
' ФОРМИРОВАНИЕ СТРАНИЦЫ ДЛЯ ОТОБРАЖЕНИЯ ДЕТАЛЬНОЙ ИНФОРМАЦИИ ДЛЯ ОСНОВНОГО ПАРАМЕТРА №3 (УСТРОЙСТВА)
' БАНКОМАТЫ БЕЗ СВЯЗИ
if T=4 then
%>
<div style="width: 100%; background-color: #363636; color: #FFFFFF; text-align: center; margin-top: 16px; margin-bottom: 16px; font-family: Verdana, Arial, helvetica, sans-serif, Geneva; font-size: 10pt; font-weight: 700; border: solid 1px #000080; ">
Детализация по связи для БПТ по процессингу и финансовому институту
</div>
<div id="tabble1"  style="width: 800px; height: 248px; margin-top: 0px; margin-right: 10px; margin-left: 10px;">
		    <table cellpadding="0" cellspacing="0" width="800px" style="font-size: 8pt; table-layout: fixed">
				<colgroup><col span=6 width="58px"><col width="*"></colgroup>
				<tbody>
				<tr>
					<td class="Head" height="20px">BRANCH</td><td class="Head">BPT</td><td class="Head">LNK</td><td class="Head">ERR</td><td class="Head">LNK_ERR</td><td class="Head">LNK 24 %</td><td class="Head">NAME</td>
				</tr>
				</tbody>
			</table>
			<div style="OVERFLOW-Y: auto; OVERFLOW-X: hidden; OVERFLOW: auto; width: 800px; height: 230px;" id="idTable1">
			    <table cellpadding="0" cellspacing="0" width="800px" style="font-size: 10pt; table-layout: fixed">
					<colgroup><col span=6 width="58px"><col width="440px"></colgroup>
					<tbody>
					<% =tblAtmLink %>
					</tbody>
				</table>
			</div>
</div>
<div align="center" style="margin-top: 8px">
<table border="0" cellspacing="2" cellpadding="0" style="border: none;">
<tr>
<td style="border: none; text-decoration: underline; cursor: hand;" onclick="ChangeDate(-1)">&laquo;</td>
<td nowrap style="border: none; font-size: 10pt; font-weight: 700; text-align: center;"><div style="float: left;">Состояние на &nbsp;</div><div id="idDate" style="float: left;"><% =left(DT_FILE, 10) %></div><div id="idTime" style="float: left;">&nbsp;<% =right(DT_FILE, 8) %></div></td>
<td style="border: none; text-decoration: underline; cursor: hand;" onclick="ChangeDate(1)">&raquo;</td>
</tr>
</table>
<div id="containerE" style="width: 100%; height: 260px; margin: 0 auto"></div>
</div>

<% end if %>


<%
' ФОРМИРОВАНИЕ СТРАНИЦЫ ДЛЯ ОТОБРАЖЕНИЯ ДЕТАЛЬНОЙ ИНФОРМАЦИИ ДЛЯ ОСНОВНОГО ПАРАМЕТРА №5 (СМС-СЕРВИС)
' SMS-Сервис
if T=5 then
%>
<div align="center" style="width: 98%; background-color: #363636; color: #FFFFFF; text-align: center; margin-top: 16px; margin-bottom: 16px; font-family: Verdana, Arial, helvetica, sans-serif, Geneva; font-size: 10pt; font-weight: 700; border: solid 1px #4572A7; ">
Детализация по СМС-сервису
</div>
<div id="tabble1"  style="width: 1030px; height: 248px; margin-top: 0px; margin-right: 10px; margin-left: 10px;">
		    <table cellpadding="0" cellspacing="0" width="1030px" style="font-size: 8pt; table-layout: fixed;">
				<colgroup><col width="180"><col width="290"><col span=7 width="80px"></colgroup>
				<tbody>
				<tr>
					<td class="Head" height="20px">TIME</td><td class="Head">SERVER</td><td class="Head">ALL</td><td class="Head">WAIT</td>
					<td class="Head">REJECTED</td><td class="Head">DECLINED</td><td class="Head">CLOSED</td><td class="Head">INPROC.</td><td class="Head">POSTED</td>
				</tr>
				</tbody>
			</table>
			<div style="OVERFLOW-Y: auto; OVERFLOW-X: hidden; OVERFLOW: auto; width: 1076px; height: 200px;" id="idTable1">
			    <table class="SMSService" cellpadding="0" cellspacing="0" width="1030px" style="font-size: 10pt; table-layout: fixed; font-size: 14pt;">
					<colgroup><col width="180"><col width="290"><col span=7 width="80px"></colgroup>
					<tbody>
					<% =tblSMS %>
					</tbody>
				</table>
			</div>
</div>
<div align="center" style="margin-top: 8px">
<table border="0" cellspacing="2" cellpadding="0" style="border: none;">
<tr>
<td style="border: none; text-decoration: underline; cursor: hand;" onclick="ChangeDate(-1)">&laquo;</td>
<td nowrap style="border: none; font-size: 10pt; font-weight: 700; text-align: center;"><div style="float: left;">Состояние на &nbsp;</div><div id="idDate" style="float: left;"><% =left(DT_FILE, 10) %></div><div id="idTime" style="float: left;">&nbsp;</div></td>
<td style="border: none; text-decoration: underline; cursor: hand;" onclick="ChangeDate(1)">&raquo;</td>
</tr>
</table>
<%=SMSServiceStringdiv%>
</div>

<% end if %>

<%
' ФОРМИРОВАНИЕ СТРАНИЦЫ ДЛЯ ОТОБРАЖЕНИЯ ДЕТАЛЬНОЙ ИНФОРМАЦИИ ДЛЯ ОСНОВНОГО ПАРАМЕТРА №6 (АВТОМАТИЗАЦИЯ ПРОЦЕССИНГА)
' SMS-Сервис
if T=6 then
%>
<div align="center" style="width: 98%; background-color: #363636; color: #FFFFFF; text-align: center; margin-top: 16px; margin-bottom: 16px; font-family: Verdana, Arial, helvetica, sans-serif, Geneva; font-size: 10pt; font-weight: 700; border: solid 1px #4572A7; ">
Детализация по автоматизации процессинга
</div>
<div>
<table class=A cellspacing=1 cellpadding=0 style="{border: none; margin-top: 0px; margin-bottom: 8px;}">
<tr>
  <th class=A width=240 nowrap>Время</th>
  <th class=A width=140 nowrap>Категория</th>
  <th class=A width=140 nowrap>Критичность</th>
  <th class=A width=140 nowrap>Приоритет</th>
  <th class=A nowrap>Свойство</th>
</tr>
<tr>
<td class=A nowrap>с <input id=BD maxlength=10 size=10 value="<%=BD%>" name="BD"><input onclick="return showCalendar('BD', 'dd.mm.y');" type=button value=" ... ">
  по&nbsp;<input id=ED maxlength=10 size=10 value="<%=ED%>" name="ED"><input onclick="return showCalendar('ED', 'dd.mm.y');" type=button value=" ... ">
</td>
<td class=A nowrap>
    <%
    lc="<option value="""""&IIF(CAT_="", " selected", "")&">Все</option>"
	set Rs=Server.CreateObject("ADODB.Recordset")
	Rs.Open "SELECT Mnemonic FROM Messages_Category ORDER BY 1", Conn
      do while not Rs.Eof
        if Rs.Fields(0) = CAT_ then
          lc=lc&"<option value="""&Rs.Fields(0)&""" selected>"&Rs.Fields(0)&"</option>"
          Sel_CAT_=Rs.Fields(0)
        else
          lc=lc&"<option value="""&Rs.Fields(0)&""">"&Rs.Fields(0)&"</option>"
        end if
        Rs.MoveNext
      loop
      Rs.Close
      Response.Write("<select id=""FCAT_"" name=""FCAT_"">"&lc&"</select>")
	%>
</td>
<td class=A nowrap>
    <%
    lc="<option value=""-1"""&IIF(ELV_=-1, " selected", "")&">Все</option>"
    lc=lc&"<option value=""1"" "&IIF(ELV_=1, "selected", "")&">1 - Information</option>"
    lc=lc&"<option value=""2"" "&IIF(ELV_=2, "selected", "")&">2 - Warning</option>"
    lc=lc&"<option value=""3"" "&IIF(ELV_=3, "selected", "")&">3 - Error</option>"
    Response.Write("<select id=""FELV_"" name=""FELV_"">"&lc&"</select>")
	%>
</td>
<td class=A nowrap>
    <%
    lc="<option value=""-1"""&IIF(PRI_=-1, " selected", "")&">Все</option>"
    lc=lc&"<option value=""0"" "&IIF(PRI_=0, "selected", "")&">0 - низкий</option>"
    lc=lc&"<option value=""1"" "&IIF(PRI_=1, "selected", "")&">1 - высокий</option>"
    Response.Write("<select id=""FPRI_"" name=""FPRI_"">"&lc&"</select>")
	%>
</td>
<td class=A nowrap>
    <%
    lc="<option value=""-1"""&IIF(PROP_=-1, " selected", "")&">Все</option>"
    lc=lc&"<option value=""0"" "&IIF(PROP_=0, "selected", "")&">0 – периодическое сообщение</option>"
    lc=lc&"<option value=""1"" "&IIF(PROP_=1, "selected", "")&">1 – событие</option>"
    lc=lc&"<option value=""2"" "&IIF(PROP_=2, "selected", "")&">2 - состояние (up/down)</option>"
    lc=lc&"<option value=""3"" "&IIF(PROP_=3, "selected", "")&">3 – WATCH (уведомление)</option>"
    Response.Write("<select id=""FPROP_"" name=""FPROP_"">"&lc&"</select>")
	%>
  &nbsp;&nbsp;<input type="button" value="Применить" name="Btn_Ok" onclick="SaveChange()">
</td>
</tr>
</table>
</div>
</form>
<div id="tabble1"  style="width: 1030px; height: 288px; margin-top: 0px; margin-right: 10px; margin-left: 10px;">
		    <table cellpadding="0" cellspacing="0" width="1030px" style="font-size: 8pt; table-layout: fixed;">
				<colgroup><col width="180"><col width="80"><col span=3 width="60px"><col width="290"></colgroup>
				<tbody>
				<tr>
					<td class="Head" height="20px">Время</td><td class="Head">Категория</td><td class="Head">Критичность</td>
					<td class="Head">Приоритет</td><td class="Head">Свойство</td><td class="Head">Сообщение</td>
				</tr>
				</tbody>
			</table>
			<div style="OVERFLOW-Y: auto; OVERFLOW-X: hidden; OVERFLOW: auto; width: 1076px; height: 260px;" id="idTable1">
			    <table cellpadding="0" cellspacing="0" width="1030px" style="font-size: 10pt; table-layout: fixed;">
					<colgroup><col width="180"><col width="80"><col span=3 width="60px"><col width="290"></colgroup>
					<tbody>
					<% =tblAuto %>
					</tbody>
				</table>
			</div>
</div>
<div align="center" style="margin-top: 8px">
<div id="inf" style="overflow:auto; width:98%; color: #FFFFFF;"></div>
</div>
<% end if %>

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