<%@ language="VBScript"%><!-- #include file="const.asp" --><!-- #include file="common.asp" -->
<%
' экран администратора модуля аварий/профработ VSP

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
    '---Init configuration params---------------------------
    incidents_config_init = ""
    incidents_config_html = ""
    incidents_config_save = ""
    incidents_config_update = ""

    api_url = ""
    sdToken = ""
    sdLogin = ""

    '-----------------------------------------------------------
    '---init CONFIG SD API---------------------------------------
    initSDAPI = ""
    sqlstr = "set dateformat dmy; select str_val val, symbol_name from [dbo].[Incidents_config] where [symbol_name] in ( 'sdURL','sdToken','sdLogin') "
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        do while (not Rs.EOF)
            initSDAPI = initSDAPI&" "&Rs.Fields("symbol_name")&"='"&Rs.Fields("val")&"'; "
            if (Rs.Fields("symbol_name")="sdURL") then
                api_url = Rs.Fields("val")
            elseif (Rs.Fields("symbol_name")="sdToken") then
                sdToken = Rs.Fields("val")
            elseif (Rs.Fields("symbol_name")="sdLogin") then
                sdLogin = Rs.Fields("val")
            end if

            Rs.MoveNext
        loop
    end if
    Rs.Close
    '---init CONFIG SD API---------------------------------------
    '-----------------------------------------------------------

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

function  deleteIncident()
    num = Request("num")
    sqlstr = "DELETE Incidents_archive where [number]='"&num&"' "
    Rs.Open sqlstr, Conn
	
	'Log Action
	actionDiscr = "Incidents_admin.asp: delete incident("&num&") from Incidents_archive."
	Cmd.CommandText="sp_LogIncidents"
	Cmd.Parameters.Refresh
	Cmd.Parameters("@Action") = actionDiscr
	Cmd.Parameters("@Userlogin") = Auth_Name
	Cmd.Execute	 
end function

function  isOnlineIncident()
    num = Request("num")
    val = Request("isOnline")
    sqlstr = "UPDATE Incidents_archive set [online]="&val&" where [number]='"&num&"' "
    Rs.Open sqlstr, Conn
	
	'Log Action
	actionDiscr = "Incidents_admin.asp: update incident("&num&") in Incidents_archive. Set isOnline="&val
	Cmd.CommandText="sp_LogIncidents"
	Cmd.Parameters.Refresh
	Cmd.Parameters("@Action") = actionDiscr
	Cmd.Parameters("@Userlogin") = Auth_Name
	Cmd.Execute
end function

function  writeIncidentToDB(actionNum)
    num = Request("num")
    detType = Request("det_type")
    detLength = Request("det_length")
    detStart = Request("det_start")
    detStop = Request("det_stop")
    detTheme = URLDecode(Request("det_theme"))
    detPrior = URLDecode(Request("det_prior"))
    detKP = URLDecode(Request("det_KP"))
    detStatus = URLDecode(Request("det_status"))

    ' 1 - activate
    ' 2 - register
    ' 3 - udpate
    ' 4 - close
    ' 5 - udpate Archive
    ' 6 - udpate Active
     sqlstr = "set dateformat dmy; exec sp_UpdateIncident "
     sqlstr = sqlstr&"  @number = "+num+", "
     sqlstr = sqlstr&"  @type = "+detType+", "
     sqlstr = sqlstr&"  @time_start_sd = '"+detStart+"', "
     sqlstr = sqlstr&"  @time_stop_sd = '"+detStop+"', "
     sqlstr = sqlstr&"  @discr_sd = '"+detTheme+"', "
     sqlstr = sqlstr&"  @status_sd = '"+detStatus+"', "
     sqlstr = sqlstr&"  @KP_sd = '"+detKP+"', "
     sqlstr = sqlstr&"  @priority_sd = '"+detPrior+"', "
     sqlstr = sqlstr&"  @action = "+actionNum

    'response.write sqlstr

     Rs.Open sqlstr, Conn
	 
	 'Log Action
	if actionNum = 1 then
		actionDiscr = "Incidents_admin.asp: Unknown action"
	elseif actionNum = 2 then
		actionDiscr = "Incidents_admin.asp: Unknown action"
	elseif actionNum = 3 then
		actionDiscr = "Incidents_admin.asp: Unknown action"
	elseif actionNum = 4 then
		actionDiscr = "Incidents_admin.asp: Unknown action"
	elseif actionNum = 5 then
		actionDiscr = "Incidents_admin.asp: insert/update incident("&num&") in Incidents_archive."
	elseif actionNum = 6 then
		actionDiscr = "Incidents_admin.asp: Unknown action"
	else
		actionDiscr = "Incidents_admin.asp: Unknown action"
	end if 
	Cmd.CommandText="sp_LogIncidents"
	Cmd.Parameters.Refresh
	Cmd.Parameters("@Action") = actionDiscr
	Cmd.Parameters("@Userlogin") = Auth_Name
	Cmd.Execute
end function

function  getIncidentByID()
    num = Request("num")
    Dim xmlHTTP
    set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlHTTP.Open "GET", api_url&num, False
    xmlHTTP.setRequestHeader "Content-Type","application/json"
    xmlHTTP.setRequestHeader "Authorization","Token token="+sdToken+", user_name="+sdLogin
    xmlHTTP.Send

    response.write xmlHTTP.responseText

'    response.Write "{""incidents"":[{""inc_id"":"&num&","&_
 '   """reg_date"":""2018-05-25T21:29:58.000Z"","&_
  '  """deadline"":""2018-05-28T03:29:58.000Z"","&_
  '  """cdl_name"":""Стандартный"","&_
  '  """inc_solution"":""Устранено на стороне банка-партнера"","&_
  '  """serv"":""Процессинг WAY4 (Уфа)"","&_
  '  """incserv"":null,""serv_id"":7789,"&_
  '  """inc_serv_id"":null,"&_
  '  """kp_name"":""КП Сопровождение пользователей Процессинга"","&_
  '  """inc_description"":""Нестабильность связи через межхостовые соединения с банком Казани"","&_
  '  """cl_date"":""2018-05-25T23:42:00.000Z"","&_
  '  """dop_info"":""Влияние: отказы в обслуживании карт по каналу с банком Казани\r\nВыявлено: 22:42 (MSK+2) 25.05.18\r\n\r\nОтветственный: банк Казани"","&_
  '  """categ"":""СБОЙ"","&_
  '  """fact_begin"":""2018-05-25T20:24:00.000Z"","&_
  '  """status"":""Устранен"","&_
  '  """inc_actualfinish"":""2018-05-25T22:30:58.000Z"","&_
  '  """prioritet"":""Стандартный""}]}"
    

         ' response.Write num
end function


if NOT IsEmpty(Request("todo")) then
	if Request("todo") = "deleteIncident" then
		deleteIncident() 
    elseif Request("todo") = "writeIncidentToDB" then
		writeIncidentToDB("5")
    elseif Request("todo") = "refreshActiveIncident" then
		writeIncidentToDB("6")
    elseif Request("todo") = "isOnlineIncident" then
		isOnlineIncident() 
    elseif Request("todo") = "getIncidentByID" then
		getIncidentByID() 
    end if 
	Response.End
end if	


%>
<!DOCTYPE HTML>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <!-- 1. Add these JavaScript inclusions in the head of your page -->
    <script src="js\jquery-3.2.1.min.js"></script>
    <script src="js\jquery-ui.min.js"></script>
    <script src="js\json2.js"></script>

    <!--<script src="js/popper.js"></script> 
    <script src="js/bootstrap-4.0.0/js/bootstrap.min.js"></script> -->
    <link rel="stylesheet" href="js/bootstrap-4.0.0/css/bootstrap.min.css">
    <link rel="stylesheet" href="js/font-awesome-4.7.0/css/font-awesome.min.css">
    <link type="text/css" href="js/jquery-ui.min.css" rel="stylesheet" />


    <style>
        a {
            text-decoration: none;
        }
        body {
            background: #fff;
            color: #000;
            font-family: Arial, Helvetica, sans-serif;
            line-height: 20px;
        }
        h4 {
            text-align: center;
        }

        /*.tr-hovered {
            background: #00ff90 !important;
            color: #000000 !important;
        }*/


        #header {
            width: 100%;
            background: #fff;
            color: #000;
            position: fixed;
            margin: 0;
            z-index:99;
        }

        #header .table-bordered {
            margin-bottom: 0;
        }

        #detailed-table {
            margin-top: 155px;
            margin-left: 0;
            margin-right: 0;
            padding: 0;
            width: 100%;
        }
        #detailed {
            width: 100%;
            padding: 0;
            margin: 0;
        }



        #header-right {
            display: inline-block;
        }

        #header-right {
            text-align: right;
            display: inline-block;
            padding: 10px 5px;
        }

        .editScreen {
            width: 100%;
            height: 100%;
            padding-top: 50px;
        }

        .editScreen .param_title {
            width: 30%;
            padding: 0 5px;
            display: inline-block;

        }

        .odd {
            background: #c2c2c2;
            color: #000000;
        }

        
        h1, h4 {
            margin: 5px;
        }

        .editScreen input, .editScreen select {
            margin: 10px 5px;
            width: 65%;
        }

        .editScreen button {
            margin: 5px 5px;
        }

        @media screen and (max-width: 1300px) {
             .editScreen {
                padding-top: 100px;
            }
        }

        @media screen and (max-width: 800px) {
             .editScreen {
                padding-top: 150px;
            }
        }

    </style>
    <script>
    //-----------------------------------------------------------------
    //-------main actions----------------------------------- 
    //deleteIncidentByNum refreshIncidentByNum  addIncidentByNum

     function writeIncidentToDB(incNum,detType,detLength,detStart,detStop,detTheme,detPrior,detKP,detStatus,isReload) {
        var r = Math.random();
        $.ajax({
            url: 'incidents_admin.asp',
            type: 'POST',
            data: { 
                todo: 'writeIncidentToDB',
                num: incNum,
                det_type: detType,
                det_length: detLength,
                det_start: detStart,
                det_stop: detStop,
                det_theme: converterhex(detTheme),
                det_prior: converterhex(detPrior),
                det_KP: converterhex(detKP),
                det_status: converterhex(detStatus),
                r:r
                    },
            error: function() {
                    console.error('Не удалось записать информацию в DB.');
                },
            success: function(result) {
                if (isReload===1) { location.reload(); }
            }
        });

    }

    function addIncidentByNum(incNum) {
        var incNum = prompt("Укажите номер НС");
        if (isNUmeric(incNum)) {
            var r = Math.random();

            $.ajax({
                url: 'incidents_admin.asp',
                type: 'POST',
                data: { 
                    todo: 'getIncidentByID',
                    num: incNum,
                    r:r
                    },                
                error: function() {
                        alert('Не удалось получить информацию от SD.');
                    },
                success: function(result) {
                    var tempRes=$.parseJSON( result );
                    if (tempRes.incidents.length==0) {
                        alert('Нет информации.');

                    } else {

                     var detType = '', detTypeDB = 0, 
                        detStart = '', detStartDB = '', fact_begin = new Date(tempRes.incidents[0].fact_begin),
                        detStop = '', detStopDB = '', cl_date = new Date(tempRes.incidents[0].cl_date),
                        detLength = 0,
                        detPrior = '',
                        detKP = '',
                        detTheme = '',
                        detStatus = '';

                    fact_begin.setTime(fact_begin.getTime() + (2*60*60*1000));
                    cl_date.setTime(cl_date.getTime() + (2*60*60*1000));

                    if ((tempRes.incidents[0].categ+'').toLowerCase()=='профилактические работы') {
                        detTypeDB = 2;
                    } else if (((tempRes.incidents[0].categ+'').toLowerCase()=='сбой')||((tempRes.incidents[0].categ+'').toLowerCase()=='сбой, не приводящий к остановке сервиса')) {
                        detTypeDB = 1;
                    }
                
                    var tmpDate = fact_begin;
                    detStartDB = padLeadingZero(tmpDate.getUTCDate(),2)+'.'+padLeadingZero((tmpDate.getUTCMonth()+1),2)+'.'+tmpDate.getUTCFullYear()+' '+padLeadingZero(tmpDate.getUTCHours(),2)+':'+padLeadingZero(tmpDate.getMinutes(),2);
                    tmpDate = cl_date;
                    detStopDB = padLeadingZero(tmpDate.getUTCDate(),2)+'.'+padLeadingZero((tmpDate.getUTCMonth()+1),2)+'.'+tmpDate.getUTCFullYear()+' '+padLeadingZero(tmpDate.getUTCHours(),2)+':'+padLeadingZero(tmpDate.getMinutes(),2);
                

                    detLength =  Math.round( (cl_date - fact_begin) /60/ 1000);
                    detPrior = tempRes.incidents[0].prioritet;
                    detKP = tempRes.incidents[0].kp_name;
                    detTheme = tempRes.incidents[0].inc_description.replace(/[\n\r\t\v\f\b]/g,' ');
                    detStatus = tempRes.incidents[0].status;

                    //console.log(incNum,detTypeDB,detLength,detStartDB,detStopDB,detTheme,detPrior,detKP,1);
                    writeIncidentToDB(incNum,detTypeDB,detLength,detStartDB,detStopDB,detTheme,detPrior,detKP,detStatus,1);
        
                    }

                }
            });



        } else {
            alert("Введенное значение не является числовым.");
        }

    }

    function deleteIncidentByNum(incNum) {
        if (confirm("Вы действитьльно хотите удалить НС №"+incNum+"?")) {
            var r = Math.random();
            $.ajax({
                url: 'incidents_admin.asp',
                type: 'POST',
                data: { 
                    todo: 'deleteIncident',
                    num: incNum,
                    r:r
                        },
                success: function(result) {
                    location.reload();
                }
            });
        }

    }

    function refreshIncidentByNum(incNum) {
/*
        1) Информация по конкретному сбою по его номеру

        GET http://sdportal/api/v1/incidents/Номер_сбоя

        2) Поиск сбоев. Определяется двумя параметрами - тип сбоя, состояние сбоя.

        GET http://sdportal/api/v1/incidents?СОСТОЯНИЕ=1&ТИП=1

        СОСТОЯНИЕ может быть:
        а) closed_today - закрытые сегодня сбои
        б) opened - открытые сбои

        ТИП может быть:
        а) sboi - категория СБОЙ
        б) prof - категория ПРОФИЛАКТИЧЕСКИЕ РАБОТЫ
        в) sboi_prof - категория СБОЙ + ПРОФИЛАКТИЧЕСКИЕ РАБОТЫ

        Примеры:
        GET http://sdportal/api/v1/incidents?closed_today=1&sboi_prof=1
        GET http://sdportal/api/v1/incidents?opened=1&sboi_prof=1
        GET http://sdportal/api/v1/incidents?opened=1&prof=1


        {
            "incidents": [
  
                {
                    "inc_id": 2782,                                      	# Номер сбоя
                    "reg_date": "2016-05-11T08:46:28.000Z",			# Время регистрации, московское время
                    "deadline": "2016-05-11T11:46:28.000Z",			# Крайний срок, московское время
                    "cdl_name": "Наивысший",						# Приоритет
                    "inc_solution": "dsgsdf",						# Решение
                    "serv": "IBSO-Retail",						# Сервис
                    "incserv": null,							# Связанный сервис
                    "serv_id": 7572,							# Номер сервиса
                    "inc_serv_id": null,							# Номер связанного сервиса
                    "kp_name": "КП Сеть передачи данных",
                    "inc_description": "тестовая_Высокий",				# Описание
                    "cl_date": "2016-10-12T15:22:00.000Z",				# Факт. время окончания работа, мск
                    "dop_info": "тестовая",						# Доп. информация
                    "categ": "ПРОФИЛАКТИЧЕСКИЕ РАБОТЫ",				# Категория
                    "fact_begin": "2016-10-10T16:16:00.000Z",			# Факт.время начала сбоя
                    "status": "В работе",						# Статус
                    "inc_actualfinish": "2016-10-12T15:22:34.000Z",		# Время окончания работ, мск
                    "prioritet": "Наивысший"						# Приоритет
                },
                {
                ...
                }
            ]
        }




        */

        var r = Math.random();
        $.ajax({
            url: 'incidents_admin.asp',
            type: 'POST',
            data: { 
                todo: 'getIncidentByID',
                num: incNum,
                r:r
                },                
            error: function() {
                    alert('Не удалось получить информацию от SD.');
                },
            success: function(result) {
                var tempRes=$.parseJSON(result);

                var detType = '', detTypeDB = 0, 
                    detStart = '', detStartDB = '', fact_begin = new Date(tempRes.incidents[0].fact_begin),
                    detStop = '', detStopDB = '', cl_date = new Date(tempRes.incidents[0].cl_date),
                    detLength = 0,
                    detPrior = '',
                    detKP = '',
                    detTheme = '',
                    detStatus = '';

                if ((tempRes.incidents[0].categ+'').toLowerCase()=='профилактические работы') {
                    detType = 'РАБОТЫ';
                    detTypeDB = 2;
                } else if (((tempRes.incidents[0].categ+'').toLowerCase()=='сбой')||((tempRes.incidents[0].categ+'').toLowerCase()=='сбой, не приводящий к остановке сервиса')) {
                    detType = 'АВАРИЯ';
                    detTypeDB = 1;
                }
                
                var tmpDate = fact_begin;
                detStart = "с "+padLeadingZero(tmpDate.getDate(),2)+'.'+padLeadingZero((tmpDate.getMonth()+1),2)+'.'+tmpDate.getFullYear()+' '+padLeadingZero((tmpDate.getUTCHours()+2),2)+':'+padLeadingZero(tmpDate.getMinutes(),2);
                detStartDB = padLeadingZero(tmpDate.getDate(),2)+'.'+padLeadingZero((tmpDate.getMonth()+1),2)+'.'+tmpDate.getFullYear()+' '+padLeadingZero((tmpDate.getUTCHours()+2),2)+':'+padLeadingZero(tmpDate.getMinutes(),2);
                tmpDate = cl_date;
                detStop = "по "+padLeadingZero(tmpDate.getDate(),2)+'.'+padLeadingZero((tmpDate.getMonth()+1),2)+'.'+tmpDate.getFullYear()+' '+padLeadingZero((tmpDate.getUTCHours()+2),2)+':'+padLeadingZero(tmpDate.getMinutes(),2);
                detStopDB = padLeadingZero(tmpDate.getDate(),2)+'.'+padLeadingZero((tmpDate.getMonth()+1),2)+'.'+tmpDate.getFullYear()+' '+padLeadingZero((tmpDate.getUTCHours()+2),2)+':'+padLeadingZero(tmpDate.getMinutes(),2);
                
                detLength =  Math.round( (cl_date - fact_begin) /60/ 1000);
                detPrior = tempRes.incidents[0].prioritet;
                detKP = tempRes.incidents[0].kp_name;
                detTheme = tempRes.incidents[0].inc_description;
                detStatus = tempRes.incidents[0].status;

                $('.n'+incNum+' .det_type').html(detType);
                $('.n'+incNum+' .det_length').html(detLength);
                $('.n'+incNum+' .det_start').html(detStart);
                $('.n'+incNum+' .det_stop').html(detStop);
                $('.n'+incNum+' .det_discr').html(detTheme);
                $('.n'+incNum+' .det_prior').html(detPrior);
                $('.n'+incNum+' .det_kp').html(detKP);
                
                //writeIncidentToDB(incNum,detTypeDB,detLength,detStartDB,detStopDB,detTheme,detKP,0);
                writeIncidentToDB(incNum,detTypeDB,detLength,detStartDB,detStopDB,detTheme,detPrior,detKP,detStatus,0);

            }
        });
    }

    function changeIsOnlineIncident(incNum, val) {
        //console.log(incNum);
        var r = Math.random();
        $.ajax({
            url: 'incidents_admin.asp',
            type: 'POST',
            data: { 
                todo: 'isOnlineIncident',
                num: incNum,
                isOnline: val,
                r:r
                    },
            success: function(result) {
                //location.reload();
            }
        });
    }



    
    //-------main actions-----------------------------------  
    //-----------------------------------------------------------------

    function padLeadingZero(num, size) {
        var s = "000000000" + num;
        return s.substr(s.length-size);
    }

    function isNUmeric(n) {
        return !isNaN(parseFloat(n))&&isFinite(n);
    }

    
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

    var sdToken = '', sdLogin = '', sdURL = 'test_sd.asp?num=';

    $(function() {
        //http://sdportal/api/v1/incidents/

        <%=initSDAPI %>

        $("input:checkbox").change(function() {
            var isOnline = 0;
            if (this.checked) {
                isOnline = 1; 
            }
            var num = $(this).attr("class");
            if (num.search(/isOnline_/i) != -1) {
                 num = num.slice(9);
                 changeIsOnlineIncident(num, isOnline);
            }
           
        });

        //Выделение при наведении парных строк таблицы
       $(".detail-table tr").hover(
            function() {
        // mouse enter
                var tmpX, tmpY;
                var tmpX = $(".detail-table tr").index(this);
                if (tmpX%2==0) {
                    tmpY = tmpX + 1;
                } else {
                    tmpY = tmpX;
                    tmpX = tmpX - 1;
                }
                
                $(".detail-table tr:eq("+tmpX+")").addClass("tr-hovered");
                $(".detail-table tr:eq("+tmpY+")").addClass("tr-hovered");
            },
            function() {
        // mouse leave
                $(".detail-table .tr-hovered").removeClass("tr-hovered");

            }
        
        );


    });

    </script>

</head>
<body>
    <div  class="row" id="header"  >
     <div class="col-md-10" ><h1>Экран администратора модуля аварий/профработ</h1></div>
     <div class="col-md-2" id="header-right" ><button id="buttonSave"  type="button" class="btn btn-success"  onClick="addIncidentByNum()"><i class="fa fa-plus"></i>&nbsp;Добавить</button></div>
  
        
       <table class='table table-bordered' >
        <thead>
        <tr><th rowspan="2" width="8%" >Категория НС</th><th rowspan="2" width="8%" >Номер НС</th><th rowspan="2"  width="10%" >Длительность </th><th  width="14%" >Начало факт.</th><th colspan="2"  width="40%" >Описание</th>
            <th rowspan="2"  width="10%" >Влияние на онлайн</th><th rowspan="2"  width="10%" >Функции</th>
        </tr>
        <tr><th  width="14%" >Устранено факт.</th><th  width="20%" >Приоритет</th><th width="20%" >КП</th></tr>
        </thead>
        </table>   
    </div>


   <!-- </div> -->


    <div class="row" id="detailed"><div class="col-md-12" id="detailed-table">
 
<%
     'response.Write "<h4>Закрытые аварии/профработы</h4><br>"
     sqlstr = "select  ia.* ,ithm.theme from  [dbo].[Incidents_archive] as ia left outer join [dbo].[Incidents_themes] ithm on ia.theme_id = ithm.ID order by time_start desc "  
     n=0
     Rs.Open sqlstr, Conn
        If not Rs.EOF then
            response.Write "<table class='table detail-table' >"
            do while (not Rs.EOF)
                if (n=0) then
                    isOdd = "odd"
                    n = 1
                else
                    isOdd = ""
                    n = 0
                end if 
                if (Rs.Fields("type")=2) then
                    detType = "Работы"
                else
                    detType = "Авария"
                end if

                detNumber = ""
                if not ISNULL(Rs.Fields("number")) then
                    detNumber = "№"&Rs.Fields("number")
                end if


                detStart = "с "&DateTimeFormat(Rs.Fields("time_start"), "dd.mm.yyyy hh:nn")
                if (Rs.Fields("type")=2) then
                    if not ISNULL(Rs.Fields("time_start_sd")) then
                        detStart = "с "&DateTimeFormat(Rs.Fields("time_start_sd"), "dd.mm.yyyy hh:nn")
                    end if
                end if         

                detStop = ""
                if not ISNULL(Rs.Fields("time_stop")) then
                     detStop = "по "&DateTimeFormat(Rs.Fields("time_stop"), "dd.mm.yyyy hh:nn")
                end if

                if not ISNULL(Rs.Fields("status_sd")) then
                    if (Rs.Fields("status_sd")="Устранен") then    'if status in SD is "Устранен"
                        detStart = "с "&DateTimeFormat(Rs.Fields("time_start_sd"), "dd.mm.yyyy hh:nn")
                        detStop = "по "&DateTimeFormat(Rs.Fields("time_stop_sd"), "dd.mm.yyyy hh:nn")
                    end if
                end if

                detLength = ""

                    if ((Rs.Fields("status")=2)or(Rs.Fields("status")=1)) then    'if Opened/Regisered
                        detLength = DateDiff("n",CDate(Rs.Fields("time_start")),Now)
                    end if
                    if (Rs.Fields("status")=4) then    'if Closed
                        detLength = DateDiff("n",CDate(Rs.Fields("time_start")),CDate(Rs.Fields("time_stop")))
                    end if
                    if not ISNULL(Rs.Fields("status_sd")) then
                        if (Rs.Fields("status_sd")="Устранен") then    'if status in SD is "Устранен"
                            detLength = DateDiff("n",CDate(Rs.Fields("time_start_sd")),CDate(Rs.Fields("time_stop_sd")))
                        end if
                    end if

                    if (detLength>0) then
                        if (detLength\60>0) then
                            detLength = detLength\60 & " ч "&(detLength-(detLength\60)*60)
                        end if
                        detLength = detLength&" мин"
                    end if


                detPrior = ""
                if not ISNULL(Rs.Fields("priority")) then
                     detPrior = Rs.Fields("priority")
                end if
                detKP = ""
                if not ISNULL(Rs.Fields("KP")) then
                     detKP = Rs.Fields("KP")
                end if


                detTheme = ""
                if not ISNULL(Rs.Fields("theme")) then
                     detTheme = Rs.Fields("theme")
                end if
                if not ISNULL(Rs.Fields("discr_sd")) then
                    detTheme = Rs.Fields("discr_sd")
                end if


                detOnline = ""
                if (Rs.Fields("online")=1) then
                     detOnline = "checked='checked'"
                end if
    
                response.Write "<tr class='n"&Rs.Fields("number")&" "&isOdd&"' ><td class='det_type' width='8%'  >"&detType&"</td><td class='det_number'  width='8%' >"&detNumber&"</td><td class='det_length'  width='10%' >"&detLength&"</td><td class='det_start'  width='14%' >"&detStart&"</td><td colspan='2' class='det_discr'  width='40%'  >"&detTheme&"</td>"
                response.Write "<td  width='10%' ><input type='checkbox' class='isOnline_"&Rs.Fields("number")&"'  "&detOnline&" /></td>"
                response.Write "<td  width='10%' ><a onClick='refreshIncidentByNum("&Rs.Fields("number")&")'><i class='fa fa-refresh'></i></a></td></tr>"
                response.Write "<tr class='n"&Rs.Fields("number")&" "&isOdd&"' ><td ></td><td  ></td><td  ></td><td class='det_stop'  width='14%'  >"&detStop&"</td><td class='det_prior'  width='20%' >"&detPrior&"</td><td class='det_kp'  width='20%'  >"&detKP&"</td>"
                response.Write "<td></td><td><a onClick='deleteIncidentByNum("&Rs.Fields("number")&")'><i class='fa fa-times'></i></a></td></tr>"

                Rs.MoveNext
            loop
            response.Write "</table >"
        end if
        Rs.Close
    
     %>
    </div></div>


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
