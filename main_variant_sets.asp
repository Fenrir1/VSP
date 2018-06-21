<%@ language="VBScript"%><!-- #include file="const.asp" --><!-- #include file="common.asp" -->
<%
' основной экран мониторинга VIP банкоматов БПТ

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

' далее (до начала html кода) в рабочие переменные считываем данные из БД для отображения элементов на странице
set Rs=Server.CreateObject("ADODB.Recordset")

    '------------------------------------------------------
    '---get Screen ID----------------------------------
    ScreenID=Request("screenid")
    if not isnumeric(ScreenID) then
        ScreenID=0
    end if

    '------------------------------------------------------
    '---get Screen Refresh----------------------------------
    Refresh = 600
    if (ScreenID>0) then
        sqlstr = "select refresh from Screen_Config where screenID="&ScreenID
        Rs.Open sqlstr, Conn
        If not Rs.EOF then
                Refresh = Rs.Fields("refresh")
        end if
        Rs.Close
    end if


%>
<!DOCTYPE HTML>
<html>
<head>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
        <meta http-equiv="X-UA-Compatible" content="ie=edge">
		<!-- 1. Add these JavaScript inclusions in the head of your page -->
		<title>Редактор экранов</title>
    	<script src="js\jquery-3.2.1.min.js"></script>
    	<script src="js\jquery-ui.min.js"></script>
        <script src="js\json2.js"></script>

<style>
        body {
            background: #294959;
            font-family: Arial;
            position: relative;
        }
        .tools-panel {
            width: 100%;
            position: fixed;
            top: 0px;
            background:rgba(0,0,0,0.5);

        }
        .tools-panel>div{
            float: left;
            margin-left: 30px;

        }
        #activeObject input, #activeObject select {
            vertical-align: top;
        }

        .speed-indicator {
            /* radius = R */
            width: 160px;  /* 2R */
            height: 80px;
            background-image:    url(speed-ticks.png);
            background-size:     cover;                 
            background-repeat:   no-repeat;
            font-size: 10pt;
            position: relative;
            cursor: pointer;

        }
        .speed-indicator .tick-label-0 {
            position: absolute;
            left: 5px;
            bottom: 0px;
        }
        .speed-indicator .tick-label-1 {
            position: absolute;
            left: 15px;
            bottom: 25px;
        }
        .speed-indicator .tick-label-2 {
            position: absolute;
            left: 35px;
            bottom: 50px;
        }
        .speed-indicator .tick-label-3 {
            position: absolute;
            left: 75px;
            bottom: 60px;
        }
        .speed-indicator .tick-label-4 {
            position: absolute;
            right: 35px;
            bottom: 50px;
        }
        .speed-indicator .tick-label-5 {
            position: absolute;
            right: 15px;
            bottom: 25px;
        }
        .speed-indicator .tick-label-6 {
            position: absolute;
            right: 5px;
            bottom: 0px;
        }
        .speed-indicator .arrow {
            width: 60px;  /* 2R */
            height: 4px; /* 2R */
            position: absolute;
            left: 20px;
            bottom: 0px;
            -moz-border-radius:  2px; /* R */
            -webkit-border-radius: 2px ; /* R */
            border-radius: 2px; /* R */
            background: #FF0000;

        }
        .speed-indicator .multiplier {
            position: absolute;
            left: 75px;
            bottom: 20px;
        }

        .simple-indicator {
            /* radius = R */
            width: 10px;  /* 2R */
            height: 10px; /* 2R */
            -moz-border-radius: 50%;; /* R */
            -webkit-border-radius: 50%;; /* R */
            border-radius: 50%;; /* R */
            box-shadow: 0 0 2px rgba(0,0,0,0.5);
            cursor: pointer;
        }
        .simple-indicator.status-green {
            /*background: #7fee1d;*/
            background: radial-gradient(farthest-side ellipse at top left, white, #7fee1d);
        }
        .simple-indicator.status-grey {
            /*background: #7fee1d;*/
            background: radial-gradient(farthest-side ellipse at top left, white, #999999);
        }
        .simple-indicator.status-yellow {
            /*background: #7fee1d;*/
            background: radial-gradient(farthest-side ellipse at top left, white, #FFCC33);
        }
        .simple-indicator.status-red {
            /*background: #7fee1d;*/
            background: radial-gradient(farthest-side ellipse at top left, white, #FF0000);
        }
        .string-indicator {
            width: 100px;  
            color: #000000;
            cursor: pointer;
        }
        .string-indicator.status-green {
            background:  #7fee1d;
        }
        .string-indicator.status-grey {
            background:  #999999;
        }
        .string-indicator.status-yellow {
            background: #FFCC33;
        }
        .string-indicator.status-red {
            background: #FF0000;
        }
        .complex-indicator {
            /* radius = R */
            width: 60px;  /* 2R */
            height: 60px; /* 2R */
            -moz-border-radius: 50%;; /* R */
            -webkit-border-radius: 50%;; /* R */
            border-radius: 50%;; /* R */
            cursor: pointer;
            /*box-shadow: 0 0 2px rgba(0,0,0,0.5);*/
        }
        .complex-indicator.status-green {
            /*background: #7fee1d;*/
            background: #7fee1d;
        }
        .complex-indicator.status-yellow {
            /*background: #7fee1d;*/
            background: #FFCC33;
        }
        .complex-indicator.status-red {
            /*background: #7fee1d;*/
            background: #FF0000;
        }
</style>
<script>

    function saveScreenRefresh() {
        var r = Math.random();
        var screenRefresh = $("#screenRefresh").val();

        $.get('editor.asp',{todo:'saveScreenRefresh',
            screenid: <%=ScreenId %>,
            refresh: screenRefresh,
            r:r
        },function(){
            //location.reload();
        }); 
    
    }

</script>
</head>
<body>

    <label for="screenRefresh">Частота обновления</label>
        <input type="text" id="screenRefresh" value="<%=Refresh %>"">
    <button onClick="saveScreenRefresh()">Установить активный экран</button>


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
<!-- Разработка: Берников И.П. -->
<!-- Для вывода графиков используется библиотека Highcharts JS - http://highsoft.com/ -->
