<%@ language="VBScript"%><!-- #include file="const.asp" --><!-- #include file="common.asp" -->
<%
' экран оператора аварий/профработ VSP

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

    '------------------------------------------------------
    '---Get Incidents Themes----------------------------------
    list_themes = getThemes() 

    '------------------------------------------------------
    '---Get Active Incidents----------------------------------
    init_incidents = ""
    sqlstr = "SELECT * FROM Incidents_active"
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        do while (not Rs.EOF)
            if (Rs.Fields("status")=1) then
                init_incidents = init_incidents & " IncidentsList["&(Rs.Fields("type")-1)&"]["&(Rs.Fields("position")-1)&"].activate(); "
            elseif  (Rs.Fields("status")=2) then
                init_incidents = init_incidents & " IncidentsList["&(Rs.Fields("type")-1)&"]["&(Rs.Fields("position")-1)&"].activate(); "
                init_incidents = init_incidents & " IncidentsList["&(Rs.Fields("type")-1)&"]["&(Rs.Fields("position")-1)&"].register(); "
            end if
            init_incidents = init_incidents & " IncidentsList["&(Rs.Fields("type")-1)&"]["&(Rs.Fields("position")-1)&"].theme_id="&Rs.Fields("theme_id")&"; "
            init_incidents = init_incidents & " IncidentsList["&(Rs.Fields("type")-1)&"]["&(Rs.Fields("position")-1)&"].SDnumber='"&Rs.Fields("number")&"'; "
            init_incidents = init_incidents & " IncidentsList["&(Rs.Fields("type")-1)&"]["&(Rs.Fields("position")-1)&"].online='"&Rs.Fields("online")&"'*1; "
            init_incidents = init_incidents & " IncidentsList["&(Rs.Fields("type")-1)&"]["&(Rs.Fields("position")-1)&"].discr_sd='"&Rs.Fields("discr_sd")&"'; "
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
	
    'Log Action
	actionDiscr = "Incidents_operator.asp: update theme("&themeID&")."
	Cmd.CommandText="sp_LogIncidents"
	Cmd.Parameters.Refresh
	Cmd.Parameters("@Action") = actionDiscr
	Cmd.Parameters("@Userlogin") = Auth_Name
	Cmd.Execute
end function

function  addTheme()
    newName = URLDecode(Request("newName"))
    sqlstr = " insert Incidents_themes (theme) values ('"&newName&"') "
    Rs.Open sqlstr, Conn
	
	'Log Action
	actionDiscr = "Incidents_operator.asp: add new theme("&newName&")."
	Cmd.CommandText="sp_LogIncidents"
	Cmd.Parameters.Refresh
	Cmd.Parameters("@Action") = actionDiscr
	Cmd.Parameters("@Userlogin") = Auth_Name
	Cmd.Execute
end function

function  deleteTheme()
    themeID=Request("id")
    if isnumeric(themeID) then
        sqlstr = " delete Incidents_themes  where [ID]="&themeID&" "
        Rs.Open sqlstr, Conn
    end if
	
	'Log Action
	actionDiscr = "Incidents_operator.asp: delete theme(theme_id="&themeID&")."
	Cmd.CommandText="sp_LogIncidents"
	Cmd.Parameters.Refresh
	Cmd.Parameters("@Action") = actionDiscr
	Cmd.Parameters("@Userlogin") = Auth_Name
	Cmd.Execute
end function

function UpdateIncident(action_to_do)
    ' 1 - activate
    ' 2 - register
    ' 3 - udpate
    ' 4 - close
    themeID=Request("theme")
    position=Request("position")
    type_inc=Request("type")

    'response.write themeID&"<br>"
    'response.write position&"<br>"
    'response.write type_inc&"<br>"
    'response.write action_to_do&"<br>"
    'response.write Request("SDnumber")&"<br>"


    if isnumeric(themeID) then
        Cmd.CommandText="sp_UpdateIncident"
        Cmd.Parameters.Refresh
        Cmd.Parameters("@theme_id") = themeID
        Cmd.Parameters("@position") = position
        Cmd.Parameters("@type") = type_inc
        Cmd.Parameters("@action") = action_to_do
        Cmd.Parameters("@number") = Request("SDnumber")
        if (action_to_do=3) then
            Cmd.Parameters("@online") = Request("online")
        end if
        Cmd.Execute
		
		
		'Log Action
		if action_to_do = 1 then
			actionDiscr = "Incidents_operator.asp: found incident(theme_id="&themeID&",position="&position&"). Insert into Incidents_active."
		elseif action_to_do = 2 then
			actionDiscr = "Incidents_operator.asp: register active incident(theme_id="&themeID&",position="&position&",number="&Request("SDnumber")&")."
		elseif action_to_do = 3 then
			actionDiscr = "Incidents_operator.asp: update active incident(theme_id="&themeID&",position="&position&",number="&Request("SDnumber")&",online="&Request("online")&")."
		elseif action_to_do = 4 then
			actionDiscr = "Incidents_operator.asp: close active incident(theme_id="&themeID&",position="&position&"). Move into Incidents_archive."
		else
			actionDiscr = "Incidents_operator.asp: Unknown action"
		end if 
		Cmd.CommandText="sp_LogIncidents"
        Cmd.Parameters.Refresh
        Cmd.Parameters("@Action") = actionDiscr
        Cmd.Parameters("@Userlogin") = Auth_Name
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
            background: #000000;
            color: #c2c2c2;
            font-family: Arial, Helvetica, sans-serif;
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
            float: left;
            margin-left: 10px;
        }
        .editPrf {
            width: 49%;
            height: 100%;
            border: 2px solid #c2c2c2;
            background: #c2c2c2;
            color:#000000;
            display: inline-block;
            margin-left: 10px;
        }
        .editErr1, .editErr2, .editPrf1, .editPrf2 {
            border: 1px solid #ffffff;
            border-radius: .25rem;
            margin: 10px 5px;
            font-weight: 600;
            height: 320px;
          
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

//----------------------------------------------------------------
//------START: Request SD API----------------------------
   function refreshIncidentByNum(incNum,num,type_id) {
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
   //console.log('refreshIncidentByNum: '+incNum+' '+num+' '+type_id);
   //console.log('Refresh:'+incNum); 
                if (tempRes.incidents.length==0) {
                        alert('Нет информации.');

                } else {

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

                /*$('.n'+incNum+' .det_type').html(detType);
                $('.n'+incNum+' .det_length').html(detLength);
                $('.n'+incNum+' .det_start').html(detStart);
                $('.n'+incNum+' .det_stop').html(detStop);
                $('.n'+incNum+' .det_discr').html(detTheme);
                $('.n'+incNum+' .det_prior').html(detPrior);
                $('.n'+incNum+' .det_kp').html(detKP);*/
                IncidentsList[type_id-1][num-1].discr_sd = detTheme;
                var incidentType = 'Prf';
                if (type_id==1) {
                    incidentType = 'Err';
                } 
                 $('.edit'+incidentType+num+' .text-info').html(IncidentsList[type_id-1][num-1].discr_sd);
                    
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


    //-----------------------------------------------------------------
    //-------Type: Incidents------------------------------ 
    var  IncidentsList = [];
    var  IncidentsTypes = ['Авария', 'Профработы'];

    function Incident() {
        this.active = false;
        this.registered = false;
        this.theme_id = 0;  
        this.discr_sd = '';
        this.SDnumber = '';
        this.online = 0; 
        this.activate = function() { this.active = true; };
        this.register = function() { this.registered = true; };
        this.close = function() {
            this.active = false;
            this.registered = false;
            this.theme_id = 0;  
            this.SDnumber = '';
            this.online = 0; 
        };
    };

    function InitIncidentsArr() {
        var  ErrList = [];   
        var  PrfList = []; 
        for(var i = 0; i<=1; i++) {
            var newErr = new Incident();
            ErrList.push(newErr);
            var newPrf = new Incident();
            PrfList.push(newPrf);
        }
        IncidentsList.push(ErrList);
        IncidentsList.push(PrfList);
    }

    function InitScreen() {
        <%=init_incidents %>

        var incidentType = 'Err';
        var num =0;

        for(var i = 0; i<=1; i++) {
            for(var j = 0; j<=1; j++) {
                if (i==0) {
                    incidentType = 'Err';
                } else {
                    incidentType = 'Prf';
                }
                num = j+1;
                $('#select'+incidentType+'Name'+num).val(IncidentsList[i][j].theme_id);
                $('#'+incidentType+'Number'+num).val(IncidentsList[i][j].SDnumber);
                $('#select'+incidentType+'Online'+num).val(IncidentsList[i][j].online);
                checkIncidentsActivity(num, i+1);
                
            }
        }
    }
    //-------Type: Indicator------------------------------ 
    //-----------------------------------------------------------------

    //-----------------------------------------------------------------
    //----START: Errors edit-------------------------------------------
    function findErr(num, type_id) {
        var incidentType = 'Err';
        if (type_id==2) {
           incidentType = 'Prf'; 
        } 
        var action_to_do = 'findErr';
        var theme_id = $('#select'+incidentType+'Name'+num).val()*1;
        if (theme_id>0) {
            var r = Math.random();
            $.ajax({
                url: 'incidents_operator.asp',
                type: 'POST',
                data: { 
                    todo: action_to_do,
                    theme: theme_id,
                    position: num,
                    type: type_id,
                    r:r
                        },
                success: function(result) {
                    IncidentsList[type_id-1][num-1].activate();
                    alert(IncidentsTypes[type_id-1]+' обнаружена');
                    checkIncidentsActivity(num, type_id);
                }
            });
            
        }
    }

    function regErr(num, type_id) {
        var incidentType = 'Err';
        if (type_id==2) {
           incidentType = 'Prf'; 
        } 
        var sd_number =  $('#'+incidentType+'Number'+num).val()||IncidentsList[type_id-1][num-1].SDnumber;
        var r = Math.random();
        //Register only if SD responded
         $.ajax({
            url: 'incidents_admin.asp',
            type: 'POST',
            data: { 
                todo: 'getIncidentByID',
                num: sd_number,
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
                        $.ajax({
                            url: 'incidents_operator.asp',
                            type: 'POST',
                            data: { 
                                todo:'registerErr',
                                position: num,
                                SDnumber: sd_number,
                                type: type_id,
                                r:r
                                    },
                            success: function(result) {
                                IncidentsList[type_id-1][num-1].register();
                                alert(IncidentsTypes[type_id-1]+' зарегистрирована');
                                refreshIncidentByNum(sd_number,num,type_id);
                                checkIncidentsActivity(num, type_id);       
                    
                            }
                        });
                    }

                }
            });

    }

    function closeErr(num, type_id) {
        var incidentType = 'Err';
        if (type_id==2) {
           incidentType = 'Prf'; 
        } 
        var theme_id = $('#select'+incidentType+'Name'+num).val()*1;
        if (theme_id>0) {
            var r = Math.random();
            $.ajax({
                url: 'incidents_operator.asp',
                type: 'POST',
                data: { 
                    todo:'closeErr',
                    theme: theme_id,
                    position: num,
                    type: type_id,
                    r:r
                        },
                success: function(result) {
                    IncidentsList[type_id-1][num-1].close();
                    alert(IncidentsTypes[type_id-1]+' закрыта');
                    $('#select'+incidentType+'Name'+num).val(0);
                    $('#'+incidentType+'Number'+num).val('');
                    checkIncidentsActivity(num, type_id);

                    //refreshIncidentByNum(num);
                }
            });
            
        }
    }

    function updateIncident(num, type_id) {
        var incidentType = 'Err';
        if (type_id==2) {
           incidentType = 'Prf'; 
        } 
        var action_to_do = 'findErr';
        var theme_id = $('#select'+incidentType+'Name'+num).val()*1;
        var sd_number =  $('#'+incidentType+'Number'+num).val()||IncidentsList[type_id-1][num-1].SDnumber;
        var on_line = $('#select'+incidentType+'Online'+num).val()*1;
        console.log(on_line);
        if (theme_id>0) {
            var r = Math.random();
            $.ajax({
                url: 'incidents_operator.asp',
                type: 'POST',
                data: { 
                    todo:'updateErr',
                    theme: theme_id,
                    SDnumber: sd_number,
                    online: on_line,
                    position: num,
                    type: type_id,
                    r:r
                        },
                success: function(result) {
                    console.log(type_id-1,num-1,'обнoвлена');
                    
                    //refreshIncidentByNum(num);
                }
            });
            
        }
    }

    //----END: Errors edit-------------------------------------------
    //-----------------------------------------------------------------


    //-----------------------------------------------------------------
    //----START: Themes edit-------------------------------------------


    function editThemes() {
        $("#editThemesDlg").dialog("open");
    }

    function refreshThemesList() {
        r = Math.random();
        $.ajax({
            url: 'incidents_operator.asp',
            type: 'POST',
            data: { 
                todo:'getThemes',
                r:r
                 },
            success: function(result) {
                $("#editThemesSelect").html(result);
                $("#editTextTheme").val("");
            }
        });
    }
    function renameTheme() {
        var newText = $("#editTextTheme").val();
        if (newText!="") {
            if (newText==$("#editThemesSelect option:selected").text()) {
                alert("Текст описания не был изменён.");
            } else {
                var r = Math.random();
                var themeId = $("#editThemesSelect").val();
                $.ajax({
                    url: 'incidents_operator.asp',
                    type: 'POST',
                    data: { 
                        todo:'renameTheme',
                        id: themeId,
                        newName:  converterhex(newText),
                        r:r
                         },
                    success: function(result) {
                        refreshThemesList();
                        alert("Тема была изменена.");
                    }
                });
             }
        } else {
            alert("Текст описания не должен быть пустым.");
        }
    }
    function deleteTheme() {
        var r = Math.random();
        var themeId = $("#editThemesSelect").val();
        if  (($("#selectErrName1").val()!=themeId)&&($("#selectErrName2").val()!=themeId)&&($("#selectPrfName1").val()!=themeId)&&($("#selectPrfName2").val()!=themeId)) {
            $.ajax({
                url: 'incidents_operator.asp',
                type: 'POST',
                data: { 
                    todo:'deleteTheme',
                    id: themeId,
                    r:r
                        },
                success: function(result) {
                    refreshThemesList();
                    alert("Тема была удалена.");
                }
            });
        } else {
            alert("Эта тема сейчас используется.");
        }
    }
    function addTheme() {
        var newText = $("#newTextTheme").val();
        if (newText!="") {
                var isExists = false;
                $( "#editThemesSelect option" ).each(function( index ) {
                  if ($( this ).text()==newText) { isExists = true;  }
                });
            if (!isExists) {
                var r = Math.random();
                $.ajax({
                    url: 'incidents_operator.asp',
                    type: 'POST',
                    data: { 
                        todo:'addTheme',
                        newName:  converterhex(newText),
                        r:r
                         },
                    success: function(result) {
                        refreshThemesList();
                        alert("Тема была создана.");

                    }
                });
            } else {
                alert("Такая тема уже существует.");
            }
            
        } else {
            alert("Текст описания не должен быть пустым.");
        }
    }
    //----END: Themes edit-------------------------------------------
    //-----------------------------------------------------------------

    function checkErrBtn(num, type_id) {
        var incidentType = 'Err';
        if (type_id==2) {
           incidentType = 'Prf'; 
        } 

        if (type_id==2) {
            if (IncidentsList[type_id-1][num-1].registered) {
                updateIncident(num, type_id);
            } 
        } else {
            if (IncidentsList[type_id-1][num-1].active) {
                updateIncident(num, type_id);
            } 
        }

        checkIncidentsActivity(num, type_id);

    }

    function checkIncidentsActivity(num, type_id) {
        var incidentType = 'Err',
            textInf = 'Авария ';
        if (type_id==2) {
           incidentType = 'Prf'; 
           textInf = 'Профработы ';
        } 
        /*
            select'+incidentType+'Name'+num
            button'+incidentType+'Find'+num
            '+incidentType+'Number'+num
            button'+incidentType+'Reg'+num
            select'+incidentType+'Online'+num
            button'+incidentType+'Close'+num
        */
        if (IncidentsList[type_id-1][num-1].discr_sd!='') {
            $('.edit'+incidentType+num+' .text-info').html(IncidentsList[type_id-1][num-1].discr_sd);
        } else {
            $('.edit'+incidentType+num+' .text-info').html(textInf+num);
        }
        if (IncidentsList[type_id-1][num-1].active) {
            $('#button'+incidentType+'Find'+num).addClass('checked');
            $('#button'+incidentType+'Find'+num).prop('disabled', true); 

            $('#select'+incidentType+'Online'+num).prop('disabled', false);
            $('#'+incidentType+'Number'+num).prop('disabled', false);
             if ($('#'+incidentType+'Number'+num).val()!='') {
                    $('#button'+incidentType+'Reg'+num).prop('disabled', false);
             } else {

                $('#button'+incidentType+'Reg'+num).prop('disabled', true);
            }
        } else {
            if (type_id==1) {
                $('#button'+incidentType+'Find'+num).removeClass('checked');
                if ($('#select'+incidentType+'Name'+num).val()*1>0) {   
                    $('#button'+incidentType+'Find'+num).prop('disabled', false);
                } else {
                    $('#button'+incidentType+'Find'+num).prop('disabled', true);
                    $('#'+incidentType+'Number'+num).prop('disabled', true);
                }
                $('#'+incidentType+'Number'+num).prop('disabled', true);
                $('#button'+incidentType+'Reg'+num).prop('disabled', true);  
                $('#select'+incidentType+'Online'+num).prop('disabled', true); 
            } else {
                $('#button'+incidentType+'Reg'+num).prop('disabled', true);  
                $('#select'+incidentType+'Online'+num).prop('disabled', true); 
            }
        }
        if (IncidentsList[type_id-1][num-1].registered) {
            
            $('#button'+incidentType+'Find'+num).prop('disabled', true); 
            $('#'+incidentType+'Number'+num).prop('disabled', false);
            $('#button'+incidentType+'Close'+num).prop('disabled', false);
            $('#button'+incidentType+'Reg'+num).addClass('checked');  
            $('#button'+incidentType+'Reg'+num).prop('disabled', true);  
        } else {
            if (type_id==2) {
                if ($('#'+incidentType+'Number'+num).val()!='') {
                        $('#button'+incidentType+'Reg'+num).prop('disabled', false);
                 } else {

                    $('#button'+incidentType+'Reg'+num).prop('disabled', true);
                }
            }
            $('#button'+incidentType+'Reg'+num).removeClass('checked');
            $('#button'+incidentType+'Close'+num).prop('disabled', true);  
        }


    }
        
     function selectTheme() {
        $("#editTextTheme").val( $("#editThemesSelect option:selected").text() );
     }


    function padLeadingZero(num, size) {
        var s = "000000000" + num;
        return s.substr(s.length-size);
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

        InitIncidentsArr();
        InitScreen();

        /*checkErrBtn(1,1);
        checkErrBtn(2,1); 
        checkErrBtn(1,2);
        checkErrBtn(2,2);*/

        $("#selectErrName1").on( "change", function( event, ui ) {  checkErrBtn(1,1);  } );
        $("#selectErrName2").on( "change", function( event, ui ) {  checkErrBtn(2,1);  } );
        $("#selectPrfName1").on( "change", function( event, ui ) {  checkErrBtn(1,2);  } );
        $("#selectPrfName2").on( "change", function( event, ui ) {  checkErrBtn(2,2);  } );

        $("#ErrNumber1").on( "change", function( event, ui ) {  checkErrBtn(1,1);  } );
        $("#ErrNumber2").on( "change", function( event, ui ) {  checkErrBtn(2,1);  } );
        $("#PrfNumber1").on( "change", function( event, ui ) {  checkErrBtn(1,2);  } );
        $("#PrfNumber2").on( "change", function( event, ui ) {  checkErrBtn(2,2);  } );

        $("#selectErrOnline1").on( "change", function( event, ui ) {  checkErrBtn(1,1);  } );
        $("#selectErrOnline2").on( "change", function( event, ui ) {  checkErrBtn(2,1);  } );
        $("#selectPrfOnline1").on( "change", function( event, ui ) {  checkErrBtn(1,2);  } );
        $("#selectPrfOnline2").on( "change", function( event, ui ) {  checkErrBtn(2,2);  } );

        $("#editThemesSelect").on( "change", function( event, ui ) {  selectTheme();  } );

        
        $('[data-toggle="tooltip"]').tooltip();

        $("#editThemesDlg").dialog({ title: "Редактирование тем аварий", width: 420, modal: true, autoOpen: false });
        $("#editThemesDlg").dialog("option",
            "buttons", {
                "Закрыть": function () { $(this).dialog("close"); location.reload(); }
            });
                
    });

    </script>

</head>
<body>
    <div class="row" id="header" >
     <div class="col-md-8" id="header-left" ><h1>Экран оператора аварий/профработ</h1></div>
     <div class="col-md-4" id="header-right" ><a data-toggle="tooltip" title="Редактировать темы аварий" onclick="editThemes()" ><i class="fa fa-cog fa-2x"></i></a></div>

    </div>
    <!--<button onClick="selectErrScreen()">Аварии</button>
    <button onClick="selectProfScreen()">Профилактики</button> -->

    <div class="editScreen">
        <div class="editErr">
            <div class="editErr1" > <!-- editErr1 bd-callout bd-callout-warning-->
                <h4 class="text-info">Авария 1</h4>
                <div >
                    <span >Описание аварии</span>
                    <select  id="selectErrName1" ><option value="0">unknown</option><%=list_themes %></select>
                </div>
                <div>
                    <button id="buttonErrFind1"  type="button" class="btn btn-secondary"  onClick="findErr(1,1)">Авария обнаружена</button>
                </div>
                <div >
                    <span>Номер НС</span>
                    <input type="text" id="ErrNumber1" value="" />
                </div>
                <div >
                    <button id="buttonErrReg1"  type="button" class="btn btn-secondary"  onClick="regErr(1,1)">Авария зарегистрирована</button>
                </div>
                <div >
                    <span>Влияние на он-лайн сервис</span>
                    <select id="selectErrOnline1" ><option value="0">Нет</option><option value="1">Да</option></select>
                </div>
                <div >
                    <button id="buttonErrClose1"  type="button" class="btn btn-success" onClick="closeErr(1,1)">Авария устранена</button>
                </div>
            </div>
            <div class="editErr2">
                <h4 class="text-info">Авария 2</h4>
                <div>
                    <span>Описание аварии</span>
                    <select id="selectErrName2" ><option value="0">unknown</option><%=list_themes %></select>
                </div>
                <div>
                    <button id="buttonErrFind2" type="button" class="btn btn-secondary"   onClick="findErr(2,1)">Авария обнаружена</button>
                </div>
                <div>
                    <span>Номер НС</span>
                    <input type="text" id="ErrNumber2" value="" />
                </div>
                <div>
                    <button id="buttonErrReg2"  type="button" class="btn btn-secondary" onClick="regErr(2,1)">Авария зарегистрирована</button>
                </div>
                <div>
                    <span>Влияние на он-лайн сервис</span>
                    <select id="selectErrOnline2" ><option value="0">Нет</option><option value="1">Да</option></select>
                </div>
                <div>
                    <button id="buttonErrClose2" type="button" class="btn btn-success" onClick="closeErr(2,1)">Авария устранена</button>
                </div>
            </div>
        </div>
<!--   Start: PROFILAXIS     -->
        <div class="editPrf">
            <div class="editPrf1">
                <h4 class="text-info" >Профработы 1</h4>
                <!--  <div>
                    <span>Описание профилактики</span>
                    <select id="selectPrfName1" ><option value="0">unknown</option><%=list_themes %></select>
                </div>
                <div>
                    <button id="buttonPrfFind1"   type="button" class="btn btn-secondary"  onClick="findErr(1,2)">Профилактика обнаружена</button>
                </div>-->
                <div>
                    <span>Номер НС</span>
                    <input type="text" id="PrfNumber1" value="" />
                </div>
                <div>
                    <button id="buttonPrfReg1"    type="button" class="btn btn-secondary" onClick="regErr(1,2)">Работы зарегистрированы</button>
                </div>
           <!--       <div>
                    <span>Влияние на он-лайн сервис</span>
                    <select id="selectPrfOnline1" ><option value="0">Нет</option><option value="1">Да</option></select>
                </div>
                <div>
                    <button id="buttonPrfClose1"  type="button" class="btn btn-success"  onClick="closeErr(1,2)">Профилактика устранена</button>
                </div>  -->
            </div>
            <div class="editPrf2">
                <h4 class="text-info" >Профработы 2</h4>
                <!-- <div>
                    <span>Описание профилактики</span>
                    <select id="selectPrfName2" ><option value="0">unknown</option><%=list_themes %></select>
                </div>
                <div>
                    <button id="buttonPrfFind2"   type="button" class="btn btn-secondary"  onClick="findErr(2,2)">Профилактика обнаружена</button>
                </div>-->
                <div>
                    <span>Номер НС</span>
                    <input type="text" id="PrfNumber2" value="" />
                </div>
                <div>
                    <button id="buttonPrfReg2"   type="button" class="btn btn-secondary"  onClick="regErr(2,2)">Работы зарегистрированы</button>
                </div>
                <!--  <div>
                    <span>Влияние на он-лайн сервис</span>
                    <select id="selectPrfOnline2" ><option value="0">Нет</option><option value="1">Да</option></select>
                </div>
                <div>
                    <button id="buttonPrfClose2"  type="button" class="btn btn-success"  onClick="closeErr(2,2)">Профилактика устранена</button>
                </div>  -->
            </div>
        </div>

    </div>

    <div id="editThemesDlg">
        <select id="editThemesSelect" ><%=list_themes %></select>
        <br/>
        <input id="editTextTheme" typ="text" value="" placeholder="Изменить описание" />
        <a data-toggle="tooltip" title="Редактировать тему" onclick="renameTheme()">
            <i class="fa fa-pencil" aria-hidden="true"></i>
        </a>
        <a data-toggle="tooltip" title="Удалить тему" onclick="deleteTheme()">
            <i class="fa fa-trash" aria-hidden="true"></i>
        </a>
        <br />
        <input id="newTextTheme" typ="text" value="" placeholder="Новая тема" />
        <a data-toggle="tooltip" title="Создать тему" onclick="addTheme()">
            <i class="fa fa-plus" aria-hidden="true"></i>
        </a>

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
<!-- Разработка: Берников И.П. -->
<!-- Для вывода графиков используется библиотека Highcharts JS - http://highsoft.com/ -->
