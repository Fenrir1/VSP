<%@ language="VBScript"%><!-- #include file="const.asp" --><!-- #include file="common.asp" -->
<%
' экран мониторинга аварий/профработ VSP

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

    '-----------------------------------------------------------
    '---init CONFIG SD API---------------------------------------
    initSDAPI = ""
    sqlstr = "set dateformat dmy; select str_val val, symbol_name from [dbo].[Incidents_config] where[symbol_name] in ( 'sdURL','sdToken','sdLogin') "
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        do while (not Rs.EOF)
            initSDAPI = initSDAPI&" "&Rs.Fields("symbol_name")&"='"&Rs.Fields("val")&"'; "
            Rs.MoveNext
        loop
    end if
    Rs.Close
    '---init CONFIG SD API---------------------------------------
    '-----------------------------------------------------------

    '-----------------------------------------------------------
    '---Get Active Incidents to refresh------------------------
    initIncidentsListToRefresh = ""
    sqlstr = "select  * from (select * from [dbo].[Incidents_active]	union select * from [dbo].[Incidents_archive] "
    sqlstr = sqlstr&" )  as ia where ([status]=2 or [status]=4) and isnull(status_sd,'')<>'Устранен'  order by [status]"
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        do while (not Rs.EOF)
'response.write Rs.Fields("number")&" "&Rs.Fields("status")&" "&Rs.Fields("status_sd")&"<br>"
                if (initIncidentsListToRefresh = "") then
                    initIncidentsListToRefresh = initIncidentsListToRefresh&"'"&Rs.Fields("number")&"'"
                else
                    initIncidentsListToRefresh = initIncidentsListToRefresh&",'"&Rs.Fields("number")&"'"
                end if

            Rs.MoveNext
        loop
    end if
    Rs.Close
'response.end
    '---Get Active Incidents to refresh------------------------
    '-----------------------------------------------------------


    '-----------------------------------------------------------
    '---Get Active Incidents and Closed with plot time left-----
    activeIncidentsCount = 0
    Dim activeIncidents(2,3)
    
    sqlstr = "select top 2 * from (select * from [dbo].[Incidents_active]	union select * from [dbo].[Incidents_archive] "
    sqlstr = sqlstr&" where  DATEDIFF(MINUTE,time_stop,getdate())<=(select int_val from Incidents_config where symbol_name = 'plot2_block4_end'))  as ia where [type]=1 order by [status]"
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        do while (not Rs.EOF)
            if (Rs.Fields("status")=4) then
                activeIncidents(activeIncidentsCount,0) = 4 ' close
                activeIncidents(activeIncidentsCount,1) = DateTimeFormat(Rs.Fields("time_start"), "yyyy,mm,dd,hh,nn")
                if isnull(Rs.Fields("time_stop")) then
                    activeIncidents(activeIncidentsCount,2) = DateTimeFormat(now, "yyyy,mm,dd,hh,nn")
                else
                    activeIncidents(activeIncidentsCount,2) = DateTimeFormat(Rs.Fields("time_stop"), "yyyy,mm,dd,hh,nn")
                end if
            else
                activeIncidents(activeIncidentsCount,0) = 1 ' active
                activeIncidents(activeIncidentsCount,1) = DateTimeFormat(Rs.Fields("time_start"), "yyyy,mm,dd,hh,nn")
                if isnull(Rs.Fields("time_stop")) then
                    activeIncidents(activeIncidentsCount,2) = DateTimeFormat(now, "yyyy,mm,dd,hh,nn")
                else
                    activeIncidents(activeIncidentsCount,2) = DateTimeFormat(Rs.Fields("time_stop"), "yyyy,mm,dd,hh,nn")
                end if

            end if


           ' if (Rs.Fields("status")=2)or((Rs.Fields("status")=4)and(Rs.Fields("status_sd")<>"Устранено"))  then
            '    if (initIncidentsListToRefresh = "") then
             '       initIncidentsListToRefresh = initIncidentsListToRefresh&"'"&Rs.Fields("number")&"'"
             '   else
             '       initIncidentsListToRefresh = initIncidentsListToRefresh&",'"&Rs.Fields("number")&"'"
             '   end if
            'end if

            activeIncidentsCount = activeIncidentsCount + 1
            Rs.MoveNext
        loop
    end if
    Rs.Close


    '-----------------------------------------------------------
    '---FOR PLOT0----------------------------------------------
    plot0CircleColor = "#6de84e"
    sqlstr = "select count(*) col from [dbo].[Incidents_active]	where  [type]=2 and [status]=2"
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        if (Rs.Fields("col")>0) then
               plot0CircleColor = "#e2e549"
        end if
    end if
    Rs.Close


    val_KPI = ""
    val_Total = "0"
    record_start = DateTimeFormat(now, "yyyy,mm-1,dd,hh,nn")
    record_stop = DateTimeFormat(now, "yyyy,mm-1,dd,hh,nn")
    now_start = DateTimeFormat(now, "yyyy,mm-1,dd,hh,nn")
    record_length =0

    if (activeIncidentsCount=0) then
        '-----------------------------------------------------------
        '---GET RECORD----------------------------------------------
        sqlstr = "   select top 1 istart.time_stop_sd record_start, iend.time_start_sd record_stop, datediff(MINUTE,istart.time_stop_sd,iend.time_start_sd) record_length "
        sqlstr = sqlstr&" from Incidents_archive istart left join Incidents_archive iend on istart.id<>iend.id where istart.time_stop_sd<iend.time_start_sd and istart.online=1 and iend.online=1 "
        sqlstr = sqlstr&" and istart.type=1 and iend.type=1  and not exists(select * from Incidents_archive  where ((time_start_sd>=istart.time_stop_sd and time_start_sd<iend.time_start_sd) "
        sqlstr = sqlstr&" or (time_stop_sd>istart.time_stop_sd and time_stop_sd<=iend.time_start_sd )) and [online]=1 ) order by datediff(MINUTE,istart.time_stop_sd,iend.time_start_sd) desc"

        'sqlstr = "   select top 1 istart.time_stop record_start, iend.time_start record_stop, datediff(MINUTE,istart.time_stop,iend.time_start) record_length "
	    'sqlstr = sqlstr&" from Incidents_archive istart left join Incidents_archive iend on istart.id<>iend.id where istart.time_stop<iend.time_start and istart.online=1 and iend.online=1 "
        'sqlstr = sqlstr&" and istart.type=1 and iend.type=1  and not exists(select * from Incidents_archive  where ((time_start>=istart.time_stop and time_start<iend.time_start) "
        'sqlstr = sqlstr&" or (time_stop>istart.time_stop and time_stop<=iend.time_start )) and [online]=1 ) order by datediff(MINUTE,istart.time_stop,iend.time_start) desc"
        Rs.Open sqlstr, Conn
        If not Rs.EOF then
            record_start = DateTimeFormat(Rs.Fields("record_start"), "yyyy,mm-1,dd,hh,nn")
            record_stop = DateTimeFormat(Rs.Fields("record_stop"), "yyyy,mm-1,dd,hh,nn")
            record_length = Rs.Fields("record_length")
            Rs.Close
            sqlstr = "select top 1 time_stop_sd now_start from Incidents_archive where [online]=1 and [type]=1 order by time_stop desc"
            Rs.Open sqlstr, Conn
            If not Rs.EOF then
                now_start = DateTimeFormat(Rs.Fields("now_start"), "yyyy,mm-1,dd,hh,nn")
            end if

        else
			if (Rs.State=1) then
				Rs.Close
			end if
             sqlstr = "set dateformat dmy; select cast(str_val as datetime) val, symbol_name from [dbo].[Incidents_config] where[symbol_name] in ( 'timer_close', 'timer_find', 'timer_now') "
             Rs.Open sqlstr, Conn
             If not Rs.EOF then
                tmpStart = now
                tmpEnd = now
                do while (not Rs.EOF)
                   if (Rs.Fields("symbol_name")="timer_find") then
                        record_start = DateTimeFormat(Rs.Fields("val"), "yyyy,mm-1,dd,hh,nn")
                        tmpStart = CDate(Rs.Fields("val"))
                    elseif (Rs.Fields("symbol_name")="timer_close") then
                        record_stop = DateTimeFormat(Rs.Fields("val"), "yyyy,mm-1,dd,hh,nn")
                        tmpEnd = CDate(Rs.Fields("val"))
                    elseif (Rs.Fields("symbol_name")="timer_now") then
                        now_start = DateTimeFormat(Rs.Fields("val"), "yyyy,mm-1,dd,hh,nn")
                   end if
                   Rs.MoveNext
                loop
                record_length = DateDiff("n",tmpStart,tmpEnd)
             end if

	    end if
        Rs.Close

        '-----------------------------------------------------------
        '---GET KPI and Val2----------------------------------------------
        sqlstr = "select str_val ,symbol_name from Incidents_config where symbol_name in ('avail_val1', 'avail_total')"
        Rs.Open sqlstr, Conn
        If not Rs.EOF then
            do while (not Rs.EOF)
                if (Rs.Fields("symbol_name")="avail_val1") then
                    val_KPI = ""
                    if (Rs.Fields("str_val")<>"") then
                        val_KPI = Rs.Fields("str_val")
                    end if
                elseif (Rs.Fields("symbol_name")="avail_total") then
                    val_Total = "0"
                    if (Rs.Fields("str_val")<>"") then
                        val_Total = Rs.Fields("str_val")
                    end if
                end if
                Rs.MoveNext
            loop
        end if
        Rs.Close

    end if
    '-----------------------------------------------------------
    '---FOR PLOT0----------------------------------------------

    '-------------------------------------------------------------
    '----------PLOTS CONFIG---------------------------------------

    init_plot1 = ""
    init_plot2 = ""

    sqlstr = "SELECT *  FROM Incidents_config where group_name in ('График «Авария обнаружена»','График «Авария устранена»') order by group_order, [id]"
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        do while (not Rs.EOF)
            if (Rs.Fields("group_name")="График «Авария обнаружена»") then
            
                init_plot1 = init_plot1 & " this."&Rs.Fields("symbol_name")&" = "

                if (Rs.Fields("isNumber")=1) then
                    init_plot1 = init_plot1 & Rs.Fields("int_val") & ";"
                else
                    init_plot1 = init_plot1 & "'" & Rs.Fields("str_val") & "';"
                end if

            elseif (Rs.Fields("group_name")="График «Авария устранена»") then
                init_plot2 = init_plot2 & " this."&Rs.Fields("symbol_name")&" = "

                if (Rs.Fields("isNumber")=1) then
                    init_plot2 = init_plot2 & Rs.Fields("int_val") & ";"
                else
                    init_plot2 = init_plot2 & "'" & Rs.Fields("str_val") & "';"
                end if

            end if

            Rs.MoveNext
        loop

    end if
    Rs.Close



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

function getThemes() 
    list = ""
    sqlstr = "SELECT [ID],theme  FROM Incidents_themes"
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        do while (not Rs.EOF)
            list = list & "<option value='"&Rs.Fields("ID")&"' >"&Rs.Fields("theme")&"</option>"
            Rs.MoveNext
        loop
    end if
    Rs.Close
    getThemes = list
end function

function  renameTheme()
    newName = URLDecode(Request("newName"))
    themeID=Request("id")
    if isnumeric(themeID) then
        sqlstr = "if exists(select  * from Incidents_themes where [ID]="&themeID&" ) "
        sqlstr = sqlstr&" update Incidents_themes set theme='"&newName&"' where [ID]="&themeID&" "
        Rs.Open sqlstr, Conn
    end if
end function

function  addTheme()
    newName = URLDecode(Request("newName"))
    sqlstr = " insert Incidents_themes (theme) values ('"&newName&"') "
    Rs.Open sqlstr, Conn
end function

function  deleteTheme()
    themeID=Request("id")
    if isnumeric(themeID) then
        sqlstr = " delete Incidents_themes  where [ID]="&themeID&" "
        Rs.Open sqlstr, Conn
    end if
end function

function UpdateIncident(action_to_do)
    ' 1 - activate
    ' 2 - register
    ' 3 - udpate
    ' 4 - close
    themeID=Request("theme")
    position=Request("position")
    type_inc=Request("type")
    if isnumeric(themeID) then
        Cmd.CommandText="sp_UpdateIncident"
        Cmd.Parameters.Refresh
        Cmd.Parameters("@theme_id") = themeID
        Cmd.Parameters("@position") = position
        Cmd.Parameters("@type") = type_inc
        Cmd.Parameters("@action") = action_to_do
        if (action_to_do=3) then
            Cmd.Parameters("@online") = Request("online")
            Cmd.Parameters("@number") = Request("SDnumber")
        end if
        Cmd.Execute
    end if
end function

if NOT IsEmpty(Request("todo")) then
	if Request("todo") = "getThemes" then
		Response.Write getThemes() 
    elseif Request("todo") = "renameTheme" then
		renameTheme() 
    elseif  Request("todo") = "deleteTheme" then
		deleteTheme() 
    elseif  Request("todo") = "addTheme" then
		addTheme()
    elseif  Request("todo") = "findErr" then
		UpdateIncident(1)
    elseif  Request("todo") = "registerErr" then
		UpdateIncident(2)
    elseif  Request("todo") = "updateErr" then
		UpdateIncident(3)
    elseif  Request("todo") = "closeErr" then
		UpdateIncident(4)
	end if 
	Response.End
end if	
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta http-equiv='refresh' content='60; url=incidents_monitor.asp'>

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
            color: inherit;
        }
        body {
            background: #000000;
            color: #c2c2c2;
            font-family: Arial, Helvetica, sans-serif;
            font-size: 20pt;
        }
        h4 {
            font-size: 20pt;
        }

        #header {
            width: 100%;
        }
        #header-right {
            width: 70%;
            display: inline-block;
        }

        #header-right {
            width: 28%;
            text-align: right;
            display: inline-block;
            padding: 10px 0;
        }

        #plot0 {
            width: 100%;
            display: none;
            margin-bottom: 20px;
        }
        #plot1 {
            width: 100%;
            display: none;
            vertical-align: top;
            margin-bottom: 20px;
        }
        #plot2 {
            width: 100%;
            display: none;
            vertical-align: top;
            margin-bottom: 20px;
        }
        #detail-table {
            width: 100%;
            /* margin: 0;*/
        }

        #plot1 > div, #plot2 > div {
            display: inline-block;
        }

        #plot1-circle, #plot2-circle {
            top: 0;
            height: auto;
        }


        #plot0 #plot0-table {
            /*min-width: 150px;*/
            height: 150px;
            display: inline-block;
        }
        #plot0-table table td {
            /*min-width: 150px;*/
            font-size: 22pt;
            color: #ffffff;
            padding: 10px 30px 0 0;

        }

        .circle {
            padding-left: 50px;
            padding-top: 30px;
        }

        #plot0 #plot0-plot {
            /*width: 450px;*/
            height: 150px;
            display: inline-block;
        }
        /*#plot0-circle, #plot1-circle {
            padding-left: 50px;
        }*/
        #plot0 #plot0-tooltip {
            display: none;
            position: fixed;
            overflow: hidden;
            background: rgba(182, 183, 186, 0.7);
            color: #000000;
            border: 1px solid #ffffff;
            border-radius: 3px;
            padding: 5px;
        }

        .block-adm {
            display: block;
            padding: 5px 10px;
            margin: 0 10px;
            border: 1px solid #99FFFF;
        }
        .block {
            display: inline-block;
            padding: 5px 10px;
            margin: 0 10px;
            border: 1px solid #99FFFF;
            line-height: 40px;
            height: 82px;
            overflow: hidden;
        }
        .block-selected {
           /* display: inline-block;
            padding: 5px 10px;
            margin: 0 10px;*/
            border: 2px solid rgb(239, 60, 57);  /*#ef3c39;*/
            box-shadow: 0 0 10px 5px rgba(239, 60, 57, 0.5);

        }

       #detail-table h4 {
           text-align: center;
           padding-bottom: 20px;
       }

/*"#66FFFF","#99FFFF"  ef3c39*/

        

        .editScreen {
            width: 100%;
            height: 100%;
        }
        .editErr {
            width: 49%;
            height: 100%;
            border: 2px solid #c2c2c2;
            background: #d7d7d7;
            color:#000000;
            display: inline-block;
        }
        .editPrf {
            width: 49%;
            height: 100%;
            border: 2px solid #c2c2c2;
            background: #c2c2c2;
            color:#000000;
            display: inline-block;
        }
        .editErr1, .editErr2, .editPrf1, .editPrf2 {
            border: 1px solid #ffffff;
            border-radius: .25rem;
            margin: 10px 5px;
            font-weight: 600;
        }
        h1, h4 {
            margin: 5px;
        }

        .editScreen input, .editScreen select {
            margin: 10px 5px;
            width: 30%;
        }

        .editScreen button {
            margin: 5px 5px;
        }

        .ui-dialog {
            background: #c2c2c2;

        }
        .ui-dialog select {
            width: 350px;
        }
        .ui-dialog .ui-dialog-titlebar {
            background: #c2c2c2;
            color: #ffffff;

        }
        .ui-dialog .ui-dialog-buttonpane {
            background: #c2c2c2;
            border-top: 2px solid #ffffff;

        }

        #editThemesDlg  input, #editThemesDlg  select  {
            margin: 5px 10px;
        }

        .checked::after {
            font-family: FontAwesome;
            content: "  \f00c";
        }
    </style>
    <script>

    //-------Config SD API-------------------------------------
    var sdURL='', sdToken='', sdLogin='';
    var incToRefresh = [<%=initIncidentsListToRefresh %>];

//-------------------------------------------------------
//------START: Request SD API----------------------------
function refreshIncidents() {
        console.log('refresh incidents');
        if (incToRefresh.length>0) {
            //console.log(incToRefresh);
            //incToRefresh.forEach( refreshIncidentByNum(item, i, arr));
            for (var i = 0; i < incToRefresh.length; ++i) {
                refreshIncidentByNum(incToRefresh[i]);
            }
        } else {
            console.log('no active incidents');
        }

        }

function refreshIncidentByNum(incNum) {
//console.log(incNum);
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
				if (tempRes.incidents.length>0) {
				
//console.log(result);				

					var detType = '', detTypeDB = 0, 
						detStart = '', detStartDB = '', fact_begin = new Date(tempRes.incidents[0].fact_begin),
						detStop = '', detStopDB = '', cl_date = new Date(tempRes.incidents[0].cl_date),
						detLength = 0,
						detPrior = '',
						detKP = '',
						detStatus = '',
						detTheme = '';

                    fact_begin.setTime(fact_begin.getTime() + (2*60*60*1000));
                    cl_date.setTime(cl_date.getTime() + (2*60*60*1000));

					if ((tempRes.incidents[0].categ+'').toLowerCase()=='профилактические работы') {
						detType = 'РАБОТЫ';
						detTypeDB = 2;
					} else if (((tempRes.incidents[0].categ+'').toLowerCase()=='сбой')||((tempRes.incidents[0].categ+'').toLowerCase()=='сбой, не приводящий к остановке сервиса')) {
						detType = 'АВАРИЯ';
						detTypeDB = 1;
					}
					
					var tmpDate = fact_begin;
					detStart = "с "+padLeadingZero(tmpDate.getUTCDate(),2)+'.'+padLeadingZero((tmpDate.getUTCMonth()+1),2)+'.'+tmpDate.getUTCFullYear()+' '+padLeadingZero(tmpDate.getUTCHours(),2)+':'+padLeadingZero(tmpDate.getMinutes(),2);
					detStartDB = padLeadingZero(tmpDate.getUTCDate(),2)+'.'+padLeadingZero((tmpDate.getUTCMonth()+1),2)+'.'+tmpDate.getUTCFullYear()+' '+padLeadingZero(tmpDate.getUTCHours(),2)+':'+padLeadingZero(tmpDate.getMinutes(),2);
					tmpDate = cl_date;
					detStop = "по "+padLeadingZero(tmpDate.getUTCDate(),2)+'.'+padLeadingZero((tmpDate.getUTCMonth()+1),2)+'.'+tmpDate.getUTCFullYear()+' '+padLeadingZero(tmpDate.getUTCHours(),2)+':'+padLeadingZero(tmpDate.getMinutes(),2);
					detStopDB = padLeadingZero(tmpDate.getUTCDate(),2)+'.'+padLeadingZero((tmpDate.getUTCMonth()+1),2)+'.'+tmpDate.getUTCFullYear()+' '+padLeadingZero(tmpDate.getUTCHours(),2)+':'+padLeadingZero(tmpDate.getMinutes(),2);
					   
					detLength =  Math.round( (cl_date - fact_begin) /60/ 1000);
					detPrior = tempRes.incidents[0].prioritet;
					detKP = tempRes.incidents[0].kp_name;
					detTheme = tempRes.incidents[0].inc_description.replace(/[\n\r\t\v\f\b]/g,' ');
					detStatus = tempRes.incidents[0].status;

				   /* $('.n'+incNum+' .det_type').html(detType);
					$('.n'+incNum+' .det_length').html(detLength);
					$('.n'+incNum+' .det_start').html(detStart);
					$('.n'+incNum+' .det_stop').html(detStop);
					$('.n'+incNum+' .det_discr').html(detTheme);
					$('.n'+incNum+' .det_prior').html(detPrior);
					$('.n'+incNum+' .det_kp').html(detKP);*/
				
					writeIncidentToDB(incNum,detTypeDB,detLength,detStartDB,detStopDB,detTheme,detPrior,detKP,detStatus,0);
				}
            }
        });
    }

    function writeIncidentToDB(incNum,detType,detLength,detStart,detStop,detTheme,detPrior,detKP,detStatus,isReload) {

        var r = Math.random();
        $.ajax({
            url: 'incidents_admin.asp',
            type: 'POST',
            data: { 
                todo: 'refreshActiveIncident',
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

//------END: Request SD API------------------------------
//-------------------------------------------------------


//-------------------------------------------------------
//------START: Plot 2------------------------------------
function initPlot2(plotId, circleId, statTime, endTime) {
            this.circleId = circleId;
            this.circleColor = '#e2e549';
            this.circleText = 'TIME';
            var tmpDate = (statTime+'').split(',');
            tmpDate.map(function(name) { return name.length;});
            this.curTimeStart  = new Date(tmpDate[0],tmpDate[1]-1,tmpDate[2],tmpDate[3],tmpDate[4]); //new Date(2017,10,10,11,15);
            tmpDate = (endTime+'').split(',');
            tmpDate.map(function(name) { return name.length;});
            this.curTimeEnd = new Date(tmpDate[0],tmpDate[1]-1,tmpDate[2],tmpDate[3],tmpDate[4]);
            this.curTimeLength = 0;
            this.curTimerLength = 23;

            this.plotId = plotId;
            this.plot2_name = 'Авария устранена';
            this.plot2_block1 = 'Звонок Ковшову Д.В.';
            this.plot2_block2 = 'НС в SD';
            this.plot2_block3 = 'Уведомление';
            this.plot2_block4 = 'Отчет';
            this.plot2_block1_adm = 'Уточнение причин\масштаба.Звонок Ковшову Д.В.+79272343537';
            this.plot2_block2_adm = 'Закрытие НС в SD';
            this.plot2_block3_adm = 'Публикация; рассылка (боевая+ДСБП); смс';
            this.plot2_block4_adm = 'Проверка последствий, заполнение отчета';
            this.plot2_block1_start = 0;
            this.plot2_block1_end = 5;
            this.plot2_block2_start = 5;
            this.plot2_block2_end = 8;
            this.plot2_block3_start = 8;
            this.plot2_block3_end = 18;
            this.plot2_block4_start = 18;
            this.plot2_block4_end = 99;
            this.selected_block = 0;
            this.plot_max_length = 1400;

             <%=init_plot2 %>

            var self = this;

            this.renderCircle = function() {
                var canvas_c = document.getElementById(this.circleId),
                    ctx_c = canvas_c.getContext("2d");  

                this.curTimeEnd 
                var tmpDate = new Date();
                var hourDiff = tmpDate - this.curTimeEnd; //in ms
                this.curTimeLength = Math.floor(hourDiff /60/ 1000); //in min  

                hourDiff = this.curTimeEnd - this.curTimeStart; //in ms
                this.curTimerLength = Math.floor(hourDiff /60/ 1000); //in min 

                this.circleText = timeConvertTimer(this.curTimerLength);                

                var centerX = canvas_c.width / 2;
                var centerY = canvas_c.height / 2;
                var radius = 90;

                ctx_c.beginPath();
                ctx_c.arc(centerX, centerY, radius, 0, 2 * Math.PI, false);
                ctx_c.fillStyle = this.circleColor;
                ctx_c.fill();


                ctx_c.font = '32pt Calibri';
                ctx_c.fillStyle = 'black';
                ctx_c.fillText(this.circleText, centerX-50, centerY+15);
                              
            };

            this.renderPlot = function() {
                $('#'+self.plotId).show();
                var tmpTxt = '';
                $('#'+self.plotId).html('');
                $('#'+self.plotId).append('<h4>'+self.plot2_name+'</h4>');

                if ((self.curTimeLength>=self.plot2_block1_start)&&(self.curTimeLength<self.plot2_block1_end)) {
                    self.selected_block = 1;
                }
                if ((self.curTimeLength>=self.plot2_block2_start)&&(self.curTimeLength<self.plot2_block2_end)) {
                    self.selected_block = 2;
                }
                if ((self.curTimeLength>=self.plot2_block3_start)&&(self.curTimeLength<self.plot2_block3_end)) {
                    self.selected_block = 3;
                }
                if ((self.curTimeLength>=self.plot2_block4_start)&&(self.curTimeLength<self.plot2_block4_end)) {
                    self.selected_block = 4;
                }
                if (self.curTimeLength>=self.plot2_block4_end) {
                    self.selected_block = 4;
                }

                if (self.selected_block>0) {
                    tmpTxt = self['plot2_block'+self.selected_block+'_adm'];
                }

                $('#'+self.plotId).append('<div class="block-adm" >'+tmpTxt+'</div>');

                var tmpWidth = Math.floor($('#'+self.plotId+' .block-adm').width()/4)-20;


                $('#'+self.plotId).append('<div><canvas height="100" width="'+(self.plot_max_length+100)+'" id="'+self.plotId+'-canvas" ></canvas></div>');
                var canvas = document.getElementById(self.plotId+'-canvas'),
                    ctx = canvas.getContext("2d");

                var startX = 40;
                
                ctx.beginPath();
                ctx.fillStyle = '#99FFFF';
                ctx.fillRect(4,44,4,4);
                ctx.fillRect((startX/3)+4,44,4,4);
                ctx.fillRect(2*(startX/3)+4,44,4,4);

                ctx.moveTo(startX,45);
                ctx.lineTo(self.plot_max_length,45);
                ctx.strokeStyle = '#99FFFF';
                ctx.stroke();

                ctx.font = '20pt Calibri';
                var tmpX = 0, prevX=0, maxX = self.plot2_block4_end;
                var stepX = Math.floor((self.plot_max_length-startX)/maxX);

                for (var i=1; i<=4; i++) {
                    tmpX = self['plot2_block'+i+'_start'];
                    if ((i==1)||(tmpX!=prevX)) {
                        ctx.moveTo(startX+tmpX*stepX,41);
                        ctx.lineTo(startX+tmpX*stepX,49);
                        ctx.strokeStyle = '#99FFFF';
                        ctx.stroke(); 
                        ctx.fillText(tmpX, startX-10+tmpX*stepX, 70);
                    }
                    tmpX = self['plot2_block'+i+'_end'];
                    ctx.moveTo(startX+tmpX*stepX,41);
                    ctx.lineTo(startX+tmpX*stepX,49);
                    ctx.strokeStyle = '#99FFFF';
                    ctx.stroke(); 
                    ctx.fillText(tmpX, startX-10+tmpX*stepX, 70);
                    prevX = tmpX;
                }

                ctx.beginPath();
                if (self.curTimeLength > self.plot2_block4_end) {
                    tmpX = self.plot_max_length;                    
                } else {
                    tmpX = startX+self.curTimeLength*stepX;    
                }
                
                ctx.moveTo(startX,45);
                ctx.lineWidth=4;
                ctx.lineTo(tmpX,45);
                ctx.strokeStyle = '#66FFFF';
                ctx.stroke();   

                ctx.beginPath();
                ctx.lineWidth=1;
                ctx.fillStyle = '#ef3c39';
                ctx.strokeStyle = '#ef3c39';
                var x = tmpX; // x coordinate
                var y = 20; // y coordinate
                var radius = 10; // Arc radius
                var startAngle = 0; // Starting point on circle
                var endAngle = Math.PI; // End point on circle   
                ctx.arc(x, y, radius, startAngle, endAngle, true); 
                ctx.lineTo(tmpX,45);
                ctx.lineTo(tmpX+radius,20);
                ctx.fill();  //stroke();         


                var tmpClass = 'block';
                for (var i=1; i<=4; i++) {
                    tmpTxt = self['plot2_block'+i];
                    if (self.selected_block==i) {
                        tmpClass = 'block block-selected';
                    } else {
                        tmpClass = 'block';
                    }
                    //console.log(tmpTxt);
                    $('#'+self.plotId).append('<div style="width: '+tmpWidth+'px;" class="'+tmpClass+'" >'+tmpTxt+'</div>');
                }
                

            };

        }
//------END: Plot 2--------------------------------------
//-------------------------------------------------------

//-------------------------------------------------------
//------START: Plot 1------------------------------------
        function initPlot1(plotId, circleId, statTime, endTime) {
            this.circleId = circleId;
            this.circleColor = '#ef3c39';
            this.circleText = 'TIME';
            //this.curTimeStart = new Date(2017,10,10,11,15);
            //this.curTimeEnd = new Date();
            var tmpDate = (statTime+'').split(',');
            tmpDate.map(function(name) { return name.length;});
            this.curTimeStart = new Date(tmpDate[0],tmpDate[1]-1,tmpDate[2],tmpDate[3],tmpDate[4]); //new Date(2017,10,10,11,15);
            tmpDate = (endTime+'').split(',');
            tmpDate.map(function(name) { return name.length;});
            this.curTimeEnd = new Date(tmpDate[0],tmpDate[1]-1,tmpDate[2],tmpDate[3],tmpDate[4]);
            this.curTimeLength = 0;

            this.plotId = plotId;
            this.plot1_name = 'Авария обнаружена';
            this.plot1_block1 = 'Анализ аварии';
            this.plot1_block2 = 'Звонок ДСП';
            this.plot1_block3 = 'Звонок Ковшову Д.В.; НС в SD';
            this.plot1_block4 = 'Уведомления';
            this.plot1_block5 = '';
            this.plot1_block1_adm = 'Классификация аварии: проявление, масштаб, критичность';
            this.plot1_block2_adm = 'Дежурный ДСП +7917409355 или СИТ\ИБ';
            this.plot1_block3_adm = 'Открытие НС в SD; Ковшов Д.В.+79272343537; Сохранение НС';
            this.plot1_block4_adm = 'Публикация; рассылка (боевая+ДСБП); смс';
            this.plot1_block5_adm = 'Устранение аварии';
            this.plot1_block1_start = 0;
            this.plot1_block1_end = 5;
            this.plot1_block2_start = 5;
            this.plot1_block2_end = 12;
            this.plot1_block3_start = 12;
            this.plot1_block3_end = 20;
            this.plot1_block4_start = 20;
            this.plot1_block4_end = 30;
            this.plot1_block5_start = 30;
            this.plot1_block5_end = 0;
            this.selected_block = 0;
            this.plot_max_length = 1400;

        <%=init_plot1 %>

            var self = this;

            this.renderCircle = function() {
                var canvas_c = document.getElementById(this.circleId),
                    ctx_c = canvas_c.getContext("2d");  

                this.curTimeEnd = new Date();

                var hourDiff = this.curTimeEnd - this.curTimeStart; //in ms
                this.curTimeLength = Math.floor(hourDiff /60/ 1000); //in min   
                this.circleText = timeConvertTimer(this.curTimeLength);                

                var centerX = canvas_c.width / 2;
                var centerY = canvas_c.height / 2;
                var radius = 90;

                ctx_c.beginPath();
                ctx_c.arc(centerX, centerY, radius, 0, 2 * Math.PI, false);
                ctx_c.fillStyle = this.circleColor;
                ctx_c.fill();

                ctx_c.font = '32pt Calibri';
                ctx_c.fillStyle = 'white';
                ctx_c.fillText(this.circleText, centerX-50, centerY+15);
                              
            };

            this.renderPlot = function() {
                var tmpTxt = '';
                $('#'+self.plotId).html('');
                $('#'+self.plotId).append('<h4>'+self.plot1_name+'</h4>');

                if ((self.curTimeLength>=self.plot1_block1_start)&&(self.curTimeLength<self.plot1_block1_end)) {
                    self.selected_block = 1;
                }
                if ((self.curTimeLength>=self.plot1_block2_start)&&(self.curTimeLength<self.plot1_block2_end)) {
                    self.selected_block = 2;
                }
                if ((self.curTimeLength>=self.plot1_block3_start)&&(self.curTimeLength<self.plot1_block3_end)) {
                    self.selected_block = 3;
                }
                if ((self.curTimeLength>=self.plot1_block4_start)&&(self.curTimeLength<self.plot1_block4_end)) {
                    self.selected_block = 4;
                }
                if (self.curTimeLength>=self.plot1_block5_start) {
                    self.selected_block = 5;
                }

                if (self.selected_block>0) {
                    tmpTxt = self['plot1_block'+self.selected_block+'_adm'];
                }

                $('#'+self.plotId).append('<div class="block-adm" >'+tmpTxt+'</div>');

                var tmpWidth = Math.floor($('#'+self.plotId+' .block-adm').width()/4)-20;

                $('#'+self.plotId).append('<div><canvas height="100" width="'+(self.plot_max_length+100)+'" id="'+self.plotId+'-canvas" ></canvas></div>');
                var canvas = document.getElementById(self.plotId+'-canvas'),
                    ctx = canvas.getContext("2d");

                var startX = 10;
                
                ctx.beginPath();
                ctx.moveTo(startX,45);
                ctx.lineTo(self.plot_max_length,45);
                ctx.strokeStyle = '#99FFFF';
                ctx.stroke();

                ctx.font = '20pt Calibri';
                ctx.fillStyle = '#99FFFF';
                ctx.fillRect(self.plot_max_length+4,44,4,4);
                ctx.fillRect(self.plot_max_length+(40/3)+4,44,4,4);
                ctx.fillRect(self.plot_max_length+2*(40/3)+4,44,4,4);

                var tmpX = 0, prevX=0, maxX = self.plot1_block5_start;
                var stepX = Math.floor((self.plot_max_length-startX)/maxX);

                for (var i=1; i<=4; i++) {
                    tmpX = self['plot1_block'+i+'_start'];
                    if ((i==1)||(tmpX!=prevX)) {
                        ctx.moveTo(startX+tmpX*stepX,41);
                        ctx.lineTo(startX+tmpX*stepX,49);
                        ctx.strokeStyle = '#99FFFF';
                        ctx.stroke(); 
                        ctx.fillText(tmpX, startX-10+tmpX*stepX, 70);
                    }
                    tmpX = self['plot1_block'+i+'_end'];
                    ctx.moveTo(startX+tmpX*stepX,41);
                    ctx.lineTo(startX+tmpX*stepX,49);
                    ctx.strokeStyle = '#99FFFF';
                    ctx.stroke(); 
                    ctx.fillText(tmpX, startX-10+tmpX*stepX, 70);
                    prevX = tmpX;
                }

                tmpX = self['plot1_block5_start'];
                if (tmpX!=prevX) {
                        ctx.moveTo(startX+tmpX*stepX,41);
                        ctx.lineTo(startX+tmpX*stepX,49);
                        ctx.strokeStyle = '#66FFFF';
                        ctx.stroke(); 
                        ctx.fillText(tmpX, startX-10+tmpX*stepX, 70);
                }

                ctx.beginPath();
                if (self.selected_block == 5) {
                    tmpX = self.plot_max_length;                    
                } else {
                    tmpX = startX+self.curTimeLength*stepX;    
                }
                
                ctx.moveTo(startX,45);
                ctx.lineWidth=4;
                ctx.lineTo(tmpX,45);
                ctx.strokeStyle = '#66FFFF';
                ctx.stroke();   

                ctx.beginPath();
                ctx.lineWidth=1;
                ctx.fillStyle = '#ef3c39';
                ctx.strokeStyle = '#ef3c39';
                var x = tmpX; // x coordinate
                var y = 20; // y coordinate
                var radius = 10; // Arc radius
                var startAngle = 0; // Starting point on circle
                var endAngle = Math.PI; // End point on circle   
                ctx.arc(x, y, radius, startAngle, endAngle, true); 
                ctx.lineTo(tmpX,45);
                ctx.lineTo(tmpX+radius,20);
                ctx.fill();  //stroke();         


                var tmpClass = 'block';
                for (var i=1; i<=4; i++) {
                    tmpTxt = self['plot1_block'+i];
                    if (self.selected_block==i) {
                        tmpClass = 'block block-selected';
                    } else {
                        tmpClass = 'block';
                    }
                    //console.log(tmpTxt);
                    $('#'+self.plotId).append('<div style="width: '+tmpWidth+'px;" class="'+tmpClass+'" >'+tmpTxt+'</div>');
                }
                

            };

        }
//------START: Plot 1------------------------------------
//-------------------------------------------------------


//-------------------------------------------------------
//------START: Plot 0------------------------------------
        //var plotObject0 = 
        function initPlot0(canvasId, tableId) {
            this.recordTimeStart = new Date(<%=record_start %>);
            this.recordTimeEnd = new Date(<%=record_stop %>);
            this.curTimeStart = new Date(<%=now_start %>);
            this.curTimeEnd = new Date();
            this.kpi = '<%=val_KPI %>';
            this.itog = <%=val_Total %>;
            this.val2 = 0;
            this.recordTimeLength = 230;
            this.recordTimeText = '';
            
            this.curTimeLength = 133;
            this.curTimeText = '';
            //var tmpWidth = $('#'+canvasId).width();
            this.maxLength = 850;
            this.maxHeight = 160;
            this.barHeight = 40;
            this.bars = [];
            this.plot_colors = [["#33FF66","#33FF99"],["#66FFFF","#99FFFF"]];
            this.plot_hover = false;
            this.hover_id = -1;
            this.canvasId = canvasId;
            this.tableId = tableId;
            this.circleColor = '<%=plot0CircleColor %>';
            this.circleText = 'TIME';

            this.refresh = function() {
                this.curTimeEnd = new Date();

                var hourDiff = this.recordTimeEnd - this.recordTimeStart; //in ms
                this.recordTimeLength = Math.round(hourDiff / 60 / 1000); //in minutes
                hourDiff = this.curTimeEnd - this.curTimeStart; //in ms
                this.curTimeLength = Math.round(hourDiff / 60 / 1000); //in minutes

                var mVal = Math.max(this.recordTimeLength,this.curTimeLength);
                var x1 = ((this.recordTimeLength*60/mVal)*this.maxLength)/100;
                var x2 = ((this.curTimeLength*60/mVal)*this.maxLength)/100;

                if (x1<2) { x1 = 2;}
                if (x2<2) { x2 = 2;}
                this.bars.push( {x: 5, y: 20, w: x1+5, h: this.barHeight} );
                this.bars.push( {x: 5, y: 20+this.barHeight+10, w: x2+5, h: this.barHeight} );

                this.curTimeText = timeConvert(this.curTimeLength);
                this.recordTimeText = timeConvert(this.recordTimeLength);
            };

            var canvas = document.getElementById(this.canvasId),
                    ctx = canvas.getContext("2d");
            var self = this;

            this.renderPlot = function() {
        /*

        this.maxHeight = 160;
            this.barHeight = 40;
        
        */
                ctx.beginPath();
                ctx.moveTo(1,1);
                ctx.lineTo(1,self.maxHeight-10);
                ctx.lineTo(self.maxLength-10, self.maxHeight-10);
                ctx.strokeStyle = '#c2c2c2';
                ctx.stroke();
                ctx.moveTo(self.maxLength-20,self.maxHeight-15);
                ctx.lineTo(self.maxLength-10,self.maxHeight-10);
                ctx.lineTo(self.maxLength-20,self.maxHeight-5);
                ctx.strokeStyle = '#c2c2c2';
                ctx.stroke();

                for(_i = 0; _b = this.bars[_i]; _i ++) {
                    ctx.fillStyle = (this.plot_hover && this.hover_id === _i) ? this.plot_colors[_i][1] : this.plot_colors[_i][0];
                    ctx.fillRect(_b.x, _b.y, _b.w, _b.h);
                }

                ctx.font = '22pt Calibri';
                ctx.fillStyle = 'white';
                ctx.fillText(this.recordTimeText, self.maxLength/2+150, 20+this.barHeight-10);
                ctx.fillText(this.curTimeText, self.maxLength/2+150, 20+2*this.barHeight);

                ctx.fillStyle = '#818182';
                ctx.fillText('Предыдущий рекорд', 10, 20+this.barHeight-10);
                ctx.fillText('Сейчас', 10, 20+2*this.barHeight);

            };

            this.renderTable = function() {
                var d1 = new Date();
                var d2 = new Date(d1.getFullYear(),0,1);
                this.val2 =Math.round(1000*( 100-(this.itog*100)/((d1-d2)/1000/60) ))/1000;


       // Math.round( ((this.itog*100)/(((d1-d2)/1000/60)*24*60))*100)/100;
                var tmpTxt = '<h4>Доступность</h4><table><tr><td>KPI,%</td><td>'+this.kpi+'</td></tr><tr><td>Сейчас</td><td>'+this.val2+'</td></tr></table>';

                $('#'+this.tableId).html(tmpTxt);
            };

            this.renderCircle = function() {
                var canvas_c = document.getElementById('plot0-circle-canvas'),
                    ctx_c = canvas_c.getContext("2d");  

                var centerX = canvas_c.width / 2;
                var centerY = canvas_c.height / 2;
                var radius = 90;

                ctx_c.beginPath();
                ctx_c.arc(centerX, centerY, radius, 0, 2 * Math.PI, false);
                ctx_c.fillStyle = this.circleColor;
                ctx_c.fill();
                              
            }

            canvas.onmousemove = function(e) {
                // Get the current mouse position
                var r = canvas.getBoundingClientRect(),
                    x = e.clientX - r.left, y = e.clientY - r.top, tmpTxt ='';
                self.plot_hover = false;
                self.hover_id = -1;
                $('#plot0-tooltip').hide();
            
                ctx.clearRect(0, 0, canvas.width, canvas.height);
            
                for(var i = self.bars.length - 1, b; b = self.bars[i]; i--) {
                    if(x >= b.x && x <= b.x + b.w &&
                       y >= b.y && y <= b.y + b.h) {
                        // The mouse honestly hits the rect
                        self.plot_hover = true;
                        self.hover_id = i;
                        $('#plot0-tooltip').css('top', (e.clientY ) + 'px');
                        $('#plot0-tooltip').css('left', (e.clientX + 20) + 'px');
                        if (i==0) {
                            tmpTxt = 
                            self.recordTimeStart.getDate()+'.'+
                            (self.recordTimeStart.getMonth()+1)+'.'+
                            self.recordTimeStart.getFullYear()+' '+
                            padLeadingZero(self.recordTimeStart.getHours(),2)+':'+
                            padLeadingZero(self.recordTimeStart.getMinutes(),2)+' - '+
                            self.recordTimeEnd.getDate()+'.'+
                            (self.recordTimeEnd.getMonth()+1)+'.'+
                            self.recordTimeEnd.getFullYear()+' '+
                            padLeadingZero(self.recordTimeEnd.getHours(),2)+':'+
                            padLeadingZero(self.recordTimeEnd.getMinutes(),2);
                        } else {
                            tmpTxt = 
                            self.curTimeStart.getDate()+'.'+
                            (self.curTimeStart.getMonth()+1)+'.'+
                            self.curTimeStart.getFullYear()+' '+
                            padLeadingZero(self.curTimeStart.getHours(),2)+':'+
                            padLeadingZero(self.curTimeStart.getMinutes(),2)+' - '+
                            self.curTimeEnd.getDate()+'.'+
                            (self.curTimeEnd.getMonth()+1)+'.'+
                            self.curTimeEnd.getFullYear()+' '+
                            padLeadingZero(self.curTimeEnd.getHours(),2)+':'+
                            padLeadingZero(self.curTimeEnd.getMinutes(),2);
                        }
                        $('#plot0-tooltip').html(tmpTxt);
                        $('#plot0-tooltip').show();

                        break;
                    }
                }
                // Draw the rectangles by Z (ASC)
                self.renderPlot();

                /*var message = 'Mouse position: ' + x + ',' + y + ' hover'+self.plot_hover+'id'+self.hover_id;
                ctx.font = '10pt Calibri';
                ctx.fillStyle = 'white';
                ctx.fillText(message, 10, 15);*/

            };

        };

//------END: Plot 0--------------------------------------
//-------------------------------------------------------

    function padLeadingZero(num, size) {
        var s = "000000000" + num;
        return s.substr(s.length-size);
    }

    function timeConvertTimer(time) { 
        var s = 0, h=0, m=0, restTime = time*1, res = '', tmpTxt = '';
        h = Math.floor(restTime / 60);
        if (h>=0) {
            restTime = restTime-60*h;
            res = res+padLeadingZero(h,2)+':';
        }
        m = Math.floor(restTime);
        if (m>=0) {
            res = res+padLeadingZero(m,2);
        }
        return res;

    }

    function timeConvert(time) { 
        //console.log(time);
        var d = 0, h=0, m=0, restTime = time*1, res = '', tmpTxt = '';
        d = Math.floor(restTime / (60*24)) ;
        if (d>=1) {
            restTime = restTime-60*24*d;
            if (d==1) {
                tmpTxt = ' день ';
            } else {
                if (d>=2&&d<=4) {
                    tmpTxt = ' дня ';
                } else {
                    tmpTxt = ' дней ';
                }

            }
            res = res+d+tmpTxt

        } 
        //console.log(d,restTime);

        h = Math.floor(restTime / 60);
        if (h>=1) {
            restTime = restTime-60*h;
            res = res+h+' ч '
        }
        //console.log(h,restTime);

        m = Math.floor(restTime);
        if (m>0) {
            res = res+m+' мин'
        }
        //console.log(m,restTime);
        return res;

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



    $(function() {

        <%=initSDAPI %>

        refreshIncidents();

        var tmpMessage = '';
        var tmpDate = new Date();
        tmpMessage = tmpMessage+padLeadingZero(tmpDate.getDate(),2)+'.'+padLeadingZero((tmpDate.getMonth()+1),2)+'.'+tmpDate.getFullYear()+' '+padLeadingZero(tmpDate.getHours(),2)+':'+padLeadingZero(tmpDate.getMinutes(),2);

        $('#curTime').html(tmpMessage);

        var plotObject0 = new initPlot0('plot0-plot-canvas','plot0-table');
        plotObject0.refresh();
        plotObject0.renderPlot();
        plotObject0.renderTable();
        plotObject0.renderCircle();

        /*var plotObject1 = new initPlot1('plot1-plot','plot1-circle-canvas');
        plotObject1.renderCircle();
        plotObject1.renderPlot();

        var plotObject2 = new initPlot2('plot2-plot','plot2-circle-canvas');
        plotObject2.renderCircle();
        plotObject2.renderPlot();*/

<%
    '-----------------------------------------------------------
    '---INIT Active Incidents and Closed with plot time left-----
    if (activeIncidentsCount>0) then
        For i = 0 to activeIncidentsCount-1
            if (activeIncidents(i,0)=4) then
                response.write " $('#plot"&(i+1)&"').css('display','flex');"
                response.write " var plotObject"&(i+1)&" = new initPlot2('plot"&(i+1)&"-plot','plot"&(i+1)&"-circle-canvas','"&activeIncidents(i,1)&"','"&activeIncidents(i,2)&"');"
                response.write " plotObject"&(i+1)&".renderCircle();"
                response.write " plotObject"&(i+1)&".renderPlot();"
            else
                response.write " $('#plot"&(i+1)&"').css('display','flex');"
                response.write " var plotObject"&(i+1)&" = new initPlot1('plot"&(i+1)&"-plot','plot"&(i+1)&"-circle-canvas','"&activeIncidents(i,1)&"','"&activeIncidents(i,2)&"');"
                response.write " plotObject"&(i+1)&".renderCircle();"
                response.write " plotObject"&(i+1)&".renderPlot();"
            end if

        Next
    else
        response.write " $('#plot0').css('display','flex');"
        response.write " var plotObject0 = new initPlot0('plot0-plot-canvas','plot0-table');"
        response.write " plotObject0.refresh();"
        response.write " plotObject0.renderPlot();"
        response.write " plotObject0.renderTable();"
        response.write " plotObject0.renderCircle();"

    end if


        %>

    });        
    </script>
</head>
<body>
    <div class="row" id="header" >
        <div class="col-md-8" id="header-left" ><h1>Экран мониторинга аварий/профработ</h1></div>
        <div class="col-md-4" id="header-right" ><span id='curTime'>TIME</span></div>
    </div> 
    <div class="row" id="plot0" >
            <div class="col-md-3 circle" id="plot0-circle" >
                <canvas width="180" height="180" id="plot0-circle-canvas" ></canvas>
            </div>
            <div id="plot0-table" class="col-md-3" >
                    
            </div>
            <div id="plot0-plot" class="col-md-6" >
                <h4>Без аварий</h4>
                <canvas width="850" height="160" id="plot0-plot-canvas" ></canvas>
                <div id="plot0-tooltip">Tooltip text...</div>
            </div>

    </div>   
    <div class="row" id="plot1" >
        <div class="col-md-2 circle" id="plot1-circle" >
            <canvas width="180" height="180" id="plot1-circle-canvas" ></canvas>
        </div>
        <div class="col-md-9" id="plot1-plot" >
                       
        </div>
    </div> 
    <div class="row" id="plot2" >
        <div class="col-md-2 circle" id="plot2-circle" >
            <canvas width="180" height="180" id="plot2-circle-canvas" ></canvas>
         </div>
        <div class="col-md-9" id="plot2-plot" >
            
        </div>
    </div>
    <div class="row" id="detail-table" ><div class="col-md-12">
<%
    '----------------------------------------------------------------
    '-------start: Detailed table------------------------------------

       if (activeIncidentsCount=0) then
            response.Write "<h4><a href='incidents_admin.asp' >АВАРИИ / ПРОФРАБОТЫ</a></h4>"
            sqlstr = "select top 2 ia.* ,ithm.theme from (select * from [dbo].[Incidents_active]	union select * from [dbo].[Incidents_archive] ) as ia left outer join [dbo].[Incidents_themes] ithm on ia.theme_id = ithm.ID order by [status],time_start desc "
        else
            'response.Write "<h4><a href='incidents_admin.asp' >Открытые аварии/профработы</a></h4>"
			response.Write "<h4><a href='incidents_admin.asp' >АВАРИИ / ПРОФРАБОТЫ</a></h4>"
					   
            sqlstr = "select top 2 ia.* ,ithm.theme from (select * from [dbo].[Incidents_active]	union select * from [dbo].[Incidents_archive] where DATEDIFF(MINUTE,time_stop,getdate())<=(select int_val from Incidents_config where symbol_name = 'plot2_block4_end') ) as ia left outer join [dbo].[Incidents_themes] ithm on ia.theme_id = ithm.ID order by [type],[status],time_start desc "
        end if
        ' select top 2 * from (select * from [dbo].[Incidents_active]	union select * from [dbo].[Incidents_archive] ) as ia order by [status],time_start desc

       ' sqlstr = "select top 2 ia.* ,ithm.theme from (select * from [dbo].[Incidents_active]	union select * from [dbo].[Incidents_archive] ) as ia left outer join [dbo].[Incidents_themes] ithm on ia.theme_id = ithm.ID order by [status],time_start desc "
        Rs.Open sqlstr, Conn
        If not Rs.EOF then
            response.Write "<table class='table detail-table' >"
            do while (not Rs.EOF)
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
                    if (Rs.Fields("status_sd")="Устранен") then    'if status in SD is "Устранено"
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
                        if (Rs.Fields("status_sd")="Устранен") then    'if status in SD is "Устранено"
                            detLength = DateDiff("n",CDate(Rs.Fields("time_start_sd")),CDate(Rs.Fields("time_stop_sd")))
                        end if
                    end if

                    if (detLength>0) then
                        if (detLength\60>0) then
                            detLength = detLength\60 & " ч "&(detLength-(detLength\60)*60)
                        end if
                        detLength = detLength&" мин"
                    end if


                'if not ISNULL(Rs.Fields("length")) then
                 '    detLength = Rs.Fields("length")
                'end if
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

                response.Write "<tr class='n"&Rs.Fields("number")&"' ><td class='det_type' >"&detType&"</td><td class='det_number' >"&detNumber&"</td><td class='det_length' >"&detLength&"</td><td class='det_start' >"&detStart&"</td><td colspan='2' class='det_discr' >"&detTheme&"</td></tr>"
                response.Write "<tr class='n"&Rs.Fields("number")&"' ><td ></td><td  ></td><td  ></td><td class='det_stop' >"&detStop&"</td><td class='det_prior' >"&detPrior&"</td><td class='det_kp' >"&detKP&"</td></tr>"

                Rs.MoveNext
            loop
            response.Write "</table >"
        end if
        Rs.Close

       
  
    '-------end: Detailed table------------------------------------  
    '----------------------------------------------------------------
    
     %>
    </div>
    </div>

    <%
        if (activeIncidentsCount<=1) then %>
    <script>

        $('#detail-table').css('margin-top','200px');
        $('#detail-table tr').css('height','64px');

    </script>

    <%    end if
         %>
    
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