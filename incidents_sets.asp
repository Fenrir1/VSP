<%@ language="VBScript"%><!-- #include file="const.asp" --><!-- #include file="common.asp" -->
<%
' экран настроек модуля аварий/профработ VSP

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
    '---Check if Role=1(Administrator)---------------------------
	sqlstr = "select * from Users where User_Login='"+Auth_Name+"'"
	Rs.Open sqlstr, Conn
    If not Rs.EOF then
		if (cInt(Rs.Fields("Role"))<>1) then
            Response.Write("<html><body><div style='text-align: center;'><span style='font-size: 14pt; font-weight: 600; color: #800000}'>Только администратры имеют доступ на эту страницу.</span></div></body></html>")
			response.end
		end if
	end if
    Rs.Close

    '------------------------------------------------------
    '---Init configuration params---------------------------
    incidents_config_init = ""
    incidents_config_html = ""
    incidents_config_save = ""
    incidents_config_update = ""

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

function getConfig() 
    incidents_config_init = ""
    incidents_config_html = ""
    incidents_config_save = ""
    incidents_config_update = ""
    n = 0
    group_id = -1
    sqlstr = "SELECT *  FROM Incidents_config order by group_order, [id]"
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        do while (not Rs.EOF)
            if (n=0) then
                isOdd = "class='odd'"
                n = 1
            else
                isOdd = ""
                n = 0
            end if
            if ((group_id = -1) or (group_id <> cInt(Rs.Fields("group_order")))) then
                incidents_config_html = incidents_config_html & "<h4>"&Rs.Fields("group_name")&"</h4>"
                group_id = cInt(Rs.Fields("group_order"))
            end if
            incidents_config_init = incidents_config_init & " IncidentsConfig."&Rs.Fields("symbol_name")&" = "
            incidents_config_save = incidents_config_save & " "&Rs.Fields("symbol_name")&" : "
            incidents_config_html = incidents_config_html & "<div "&isOdd&" ><div class='param_title' ><span>"&Rs.Fields("name")&"</span></div><input type='text' id='"&Rs.Fields("symbol_name")&"' "
            incidents_config_update = incidents_config_update & " UPDATE Incidents_config set "
            if (Rs.Fields("isNumber")=1) then
                incidents_config_init = incidents_config_init & Rs.Fields("int_val") & ";"
                'incidents_config_save = incidents_config_save & " IncidentsConfig."&Rs.Fields("symbol_name")&", "
                incidents_config_save = incidents_config_save & " $('#"&Rs.Fields("symbol_name")&"').val() , "
                incidents_config_html = incidents_config_html & " value='"&Rs.Fields("int_val")&"' /></div>"
                incidents_config_update = incidents_config_update & " int_val=" & Request(Rs.Fields("symbol_name")) &" where symbol_name='"&Rs.Fields("symbol_name")&"' ;"
            else
                incidents_config_init = incidents_config_init & "'" & Rs.Fields("str_val") & "';"
                incidents_config_save = incidents_config_save & " converterhex($('#"&Rs.Fields("symbol_name")&"').val()), "
                incidents_config_html = incidents_config_html & " value='"&Rs.Fields("str_val")&"' /></div>"
                incidents_config_update = incidents_config_update & " str_val='" & URLDecode(Request(Rs.Fields("symbol_name"))) &"' where symbol_name='"&Rs.Fields("symbol_name")&"' ;"
            end if

            Rs.MoveNext
        loop

    end if
    Rs.Close
end function

function  setConfig()
    sqlstr = incidents_config_update
    response.Write sqlstr
    Rs.Open sqlstr, Conn
end function

    getConfig() 

if NOT IsEmpty(Request("todo")) then
	if Request("todo") = "getConfig" then
		getConfig() 
    elseif Request("todo") = "setConfig" then
		setConfig() 
    end if 
	Response.End
end if	


%>
<!DOCTYPE HTML>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1251">

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
            line-height: 20px;
        }

        #header {
            width: 100%;
            background: #000000;
            color: #c2c2c2;
            position: fixed;
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

        .editScreen>div:not(.odd) {
            background: #c2c2c2;
            color: #000000;

        }

        .odd {
            background: #d7d7d7;
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
    //-------Object: IncidentsConfig----------------------------------- 
    var  IncidentsConfig = {};

    function InitIncidentsConfig() {
            <%=incidents_config_init %>
    }

    function SetIncidentsConfig() {
        var r = Math.random();
        $.ajax({
            url: 'incidents_sets.asp',
            type: 'POST',
            data: { 
                todo: 'setConfig',
                <%=incidents_config_save %>
                r:r
                    },
            success: function(result) {
                location.reload();
            }
        });
    }

    function RefreshIncidentsConfig() {
        location.reload();
    }
    //-------Object: IncidentsConfig----------------------------------- 
    //-----------------------------------------------------------------

    
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
        InitIncidentsConfig();

    });

    </script>

</head>
<body>
    <div  class="row" id="header" >
     <div class="col-md-8" id="header-left" ><h1>Экран настроек модуля аварий/профработ</h1></div>
     <div class="col-md-4" id="header-right" >
         <button id="buttonRefr"  type="button" class="btn btn-secondary"  onClick="RefreshIncidentsConfig()">Отмена</button>
         <button id="buttonSave"  type="button" class="btn btn-success"  onClick="SetIncidentsConfig()">Сохранить</button>

     </div>

    </div>
    <div class="editScreen">
<%=incidents_config_html %>
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
