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
	Rs.Open "SELECT * FROM Users WHERE (User_Name='"&FUserName&"')", Conn
	usrRole=Rs.Fields("Role")
	Rs.Close
	if usrRole<>1 then
		Conn.Close
		set Conn = Nothing
		set Cmd = Nothing
		set Rs = Nothing
		Response.Write("<html><body><div style='text-align: center;'><span style='font-size: 14pt; font-weight: 600; color: #800000}'>Недостаточно прав для редактирования справочников.</span></div></body></html>")
		Response.End
	end if
		' Права администратора




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


'------------------------------------
'--START: Link to QOS----------------
'------------------------------------	
'Периодичность рассылки FV PeriodFV 
'Периодичность рассылки EV PeriodEV 
'Периодичность рассылки VA PeriodVA 
'Периодичность рассылки VC PeriodVC 

QOSLink = ""
Top5Interval = ""
DaysCountTime = ""
PeriodVC = 0
PeriodVA = 0
PeriodFV = 0
PeriodEV = 0
sqlstr = "select *  from VIP_Config"
Rs.Open sqlstr, Conn
If not Rs.EOF then
	QOSLink = Rs.Fields("QOSLink")
	Top5Interval = Rs.Fields("Top5Interval")
	DaysCountTime = Rs.Fields("DaysCount")

	VIP_Title = Rs.Fields("VIP_Title")
	
	PeriodVC = Rs.Fields("PeriodVC")
	PeriodVA = Rs.Fields("PeriodVA")
	PeriodFV = Rs.Fields("PeriodFV")
	PeriodEV = Rs.Fields("PeriodEV")
end if
Rs.Close
'------------------------------------
'--END: Link to QOS----------------
'------------------------------------	

function save_config()
	ql = ""
	ti = ""
	dc = ""
	ql = Request("ql")
	ti = Request("ti")
	dc = Request("dc")

	vtit = URLDecode(Request("vtit"))
	
	PeriodVC = Request("pvc")
	PeriodVA = Request("pva")
	PeriodFV = Request("pfv")
	PeriodEV = Request("pev")
	
	sqlstr = "UPDATE VIP_Config set QOSLink='"&ql&"', Top5Interval="&ti&", DaysCount="&dc&", VIP_Title='"&vtit&"',"
	sqlstr = sqlstr&" PeriodVC="&PeriodVC&" , PeriodVA="&PeriodVA&" , PeriodFV="&PeriodFV&" , PeriodEV="&PeriodEV&" "
	
	'response.write sqlstr
	Rs.Open sqlstr, Conn
end function

if NOT IsEmpty(Request("todo")) then
	if Request("todo") = "save_config" then
		save_config
	end if
	Response.End
end if

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------

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
    /*-------------START: Save config------------------------------------------------------------*/
    /*-------------------------------------------------------------------------------------------*/	
	function saveConfig(tableNum, deviceId, deviceType, startTime) {
	//save_comments Request("tableType"), Request("deviceId"), Request("deviceType"), Request("startTime"), Request("comm1"), Request("comm2")
		var ql = $('#QOSLink').val();
		var ti = $('#Top5Interval').val();
		ti = ti*60;
		var dc = $('#DaysCountTime').val();

		var vtit = converterhex($('#VIP_Title').val());
		
		var pvc = $('#PeriodVC').val();
		var pva = $('#PeriodVA').val();
		var pfv = $('#PeriodFV').val();
		var pev = $('#PeriodEV').val();
		
		
		r = Math.random(); 
		$.get('config3.asp',{todo:'save_config', 
							ql: ql,
							ti: ti, 
							dc: dc,
							vtit: vtit,
							pvc: pvc,
							pva: pva,
							pfv: pfv,
							pev: pev,
							
							r:r},
			function(data){ /*location.reload();*/ });
	}
			
    /*-------------------------------------------------------------------------------------------*/
    /*-------------END: Save config--------------------------------------------------------------*/
    /*-------------------------------------------------------------------------------------------*/				

	</script>	
</head>
<body>
<div align="center">
<table border="0" width="850px" align='left' style="border: none;">
	<tr><td width='450px' style="text-align: left; border: none; font-size: 14pt; font-weight: 200;">Счетчик дней для интервалов неработоспособности (часы)</td><td style='text-align: left;' ><input id='DaysCountTime' value='<%=DaysCountTime %>' ></td></tr>
	<tr><td style="text-align: left; border: none; font-size: 14pt; font-weight: 200;">Настройка длительности интервала TOP 5 (минуты)</td><td  style='text-align: left;' ><input id='Top5Interval' value='<%=Top5Interval/60 %>' ></td></tr>
	<tr><td style="text-align: left; border: none; font-size: 14pt; font-weight: 200;">Переход в QOS</td><td  style='text-align: left;' ><input style='width: 400px;' id='QOSLink' value='<%=QOSLink %>' ></td></tr>

	<tr><td style="text-align: left; border: none; font-size: 14pt; font-weight: 200;">Заголовок экрана</td><td  style='text-align: left;' ><input style='width: 400px;' id='VIP_Title' value='<%=VIP_Title %>' ></td></tr>
	
	
	<tr><td style="text-align: left; border: none; font-size: 14pt; font-weight: 200;">Периодичность рассылки FV (минуты)</td><td  style='text-align: left;' ><input style='width: 400px;' id='PeriodFV' value='<%=PeriodFV %>' ></td></tr>
	<tr><td style="text-align: left; border: none; font-size: 14pt; font-weight: 200;">Периодичность рассылки EV (минуты)</td><td  style='text-align: left;' ><input style='width: 400px;' id='PeriodEV' value='<%=PeriodEV %>' ></td></tr>
	<tr><td style="text-align: left; border: none; font-size: 14pt; font-weight: 200;">Периодичность рассылки VA (минуты)</td><td  style='text-align: left;' ><input style='width: 400px;' id='PeriodVA' value='<%=PeriodVA %>' ></td></tr>
	<tr><td style="text-align: left; border: none; font-size: 14pt; font-weight: 200;">Периодичность рассылки VC (минуты)</td><td  style='text-align: left;' ><input style='width: 400px;' id='PeriodVC' value='<%=PeriodVC %>' ></td></tr>
	<tr><td colspan=2 style="border: none;"><button onClick='saveConfig()' >Сохранить</button></td></tr>
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