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

    '------------------------------------------------------
    '---get Indicators list----------------------------------
    IndicatorsList = ""
    if (ScreenID>0) then
        Dim fso ' объявляем переменую fso – экземпляр FileSystemObject
        Set fso = CreateObject("Scripting.FileSystemObject") ' создаем экземпляр объекта FileSystemObject
        ScriptFileName="main_variant_"&ScreenID&".js"
        if not fso.FileExists(ScriptFileName)  Then
            sqlstr = "select * from Screen_List where id="&ScreenID
            Rs.Open sqlstr, Conn
            If not Rs.EOF then
                    IndicatorsList = IndicatorsList & Rs.Fields("indicators_list")
                    Background = Rs.Fields("background_path")
            end if
            Rs.Close
        else
            ScriptFileName=""
        end if

    end if

function saveScreen()
    Indicators = Request("Indicators")
	sqlstr = "EXEC sp_SaveScreen @Indicators='"&Indicators&"' "
    response.write Request("Indicators")
 	Rs.Open sqlstr, Conn


end function

    'todo:'saveScreen',Indicators:IndicatorsList
if NOT IsEmpty(Request("todo")) then
	if Request("todo") = "saveScreen" then
		saveScreen()
	end if 
	Response.End
end if	
%>
<!DOCTYPE HTML>
<html>
<head>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
        <meta http-equiv="X-UA-Compatible" content="ie=edge">
		<meta http-equiv='refresh' content='<%=Refresh %>'; url=http://ufa-qos01ow/vsp/main_variant.asp'>
		<!-- 1. Add these JavaScript inclusions in the head of your page -->
		<title>Экран мониторинга</title>
    	<script src="js\jquery-3.2.1.min.js"></script>
    	<script src="js\jquery-ui.min.js"></script>
        <script src="js\json2.js"></script>
<%  

   ' response.write ScriptFileName
   ' response.end
   if ScriptFileName<>"" then
%> <script src="<%=ScriptFileName %>"></script>
<% else 
%> <script src="main_variant.js"></script>
<% end if %>

<style>
        body {
            background: url(img/screen.png) no-repeat #294959;
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
            width: 200px;  /* 2R */
            height: 100px;
            background-image:    url(img/speed-ticks.png);
            background-size:     cover;                 
            background-repeat:   no-repeat;
            font-size: 12pt;
            color: #FFFFFF;
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
            left: 20px;
            bottom: 30px;
        }
        .speed-indicator .tick-label-2 {
            position: absolute;
            left: 45px;
            bottom: 60px;
        }
        .speed-indicator .tick-label-3 {
            position: absolute;
            left: 85px;
            bottom: 70px;
        }
        .speed-indicator .tick-label-4 {
            position: absolute;
            right: 45px;
            bottom: 60px;
        }
        .speed-indicator .tick-label-5 {
            position: absolute;
            right: 20px;
            bottom: 30px;
        }
        .speed-indicator .tick-label-6 {
            position: absolute;
            right: 5px;
            bottom: 0px;
        }
        .speed-indicator .arrow {
            width: 80px;  /* 2R */
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
            width: 200px;  
            font-size: 16pt;
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

     var DrawSpeedDigits = function(minSpeed, maxSpeed, speedId) {
        for(var i=0; i<=6; i++) {
            labelNumber = i;
            speedLabel = Math.round(minSpeed*1 + i*(maxSpeed*1-minSpeed*1)/6);
            $('#'+speedId).append( "<span class='tick-label-"+labelNumber+"' >"+speedLabel+"</span>" );
        }
    }

    var DrawSpeedArrow = function(minSpeed, maxSpeed, curSpeed, speedId) {
        if (curSpeed>maxSpeed) {
            maxSpeed = maxSpeed*10;
            minSpeed = minSpeed*10;
            $('#'+speedId).append( "<span class='multiplier' >x10</span>" );
        }
        $('#'+speedId).append( "<div class='arrow' >&nbsp;</div>" );
        var alpha = 180*curSpeed/(maxSpeed-minSpeed);
         
        $('#'+speedId+" .arrow").css({
        "transform": "rotate("+alpha+"deg)",
        "transform-origin": "right center" });

    }


    var IndicatorsList = [];
//-------Type: Indicator------------------------------    
    function Indicator(indicatorType, indicatorId) {
        this.type = indicatorType;
        this.id = indicatorId;
        this.class = indicatorType;
        this.top = 150;
        this.left = 0;
        this.content = '';
        this.MsgCode = '';
        this.connectedIndicators = [];
        this.isVisible = 1;
        this.style = "position: absolute; top: "+this.top+"px; left: "+this.left+"px;"

        if ((indicatorType == 'simple-indicator')||(indicatorType == 'complex-indicator') ) {
            this.status = 'status-green';
            this.class += ' '+this.status;
            this.content = ''
        } 
        if (indicatorType == 'string-indicator') {
            this.status = 'status-green';
            this.class += ' '+this.status;
            this.content = 'Sample text'
        }
        if (indicatorType == 'speed-indicator') {
                this.minSpeed = 0;
                this.maxSpeed = 180;
                this.curSpeed = 0;
                this.content = ''
        }
            
     
    };


    function getIndicatorStatusID(indicatorid) {
        var resStat = '';
        for(var i=0; i<IndicatorsList.length; i++) {
            if (IndicatorsList[i].id==indicatorid) {
                resStat = IndicatorsList[i].status;
                break;
            }
        }
        return resStat;
    }

    function getIndicatorStatus(Indicator){
        var resStat = '';
        var r = Math.random();
        if ((Indicator.type=='string-indicator')||(Indicator.type=='simple-indicator')) {
            
            resStat = MsgCodeStatus[Indicator.MsgCode];
    
           /* $.ajax({
                url: 'editor.asp',
                type: 'GET',
                async: false,
                data: { 
                    todo:'getIndicatorStatus',
                    id:Indicator.id,
                    type:Indicator.type,
                    msgcode:Indicator.MsgCode,
                    r:r
                     },
                    success: function(res) {
                            resStat = res;
                            //return resStat;
                    }
              });*/

        } 
        if (Indicator.type=='complex-indicator') {
            var tempStat='';
            if (Indicator.connectedIndicators.length>0) {
                Indicator.connectedIndicators.forEach(function(item, i, arr) {
                    tempStat = getIndicatorStatusID(item);
                    if (tempStat=='status-red') {
                        resStat='status-red';
                    } else {
                        if ((resStat!='status-red')&&(tempStat=='status-yellow')) {
                            resStat='status-yellow';    
                        }
                    }
                });
            }

        }

        return resStat;
    }

    function getOperationPerMin() {
        var OperationPerMin = 0;
        var r = Math.random();
         $.ajax({
                url: 'editor.asp',
                type: 'GET',
                async: false,
                data: { 
                    todo:'getOperationPerMin',
                    r:r
                     },
                    success: function(res) {
                            OperationPerMin = res;
                    }
              });

        return OperationPerMin;
    }


    function drawIndicator(Indicator) {
            var curType = Indicator.type;
            var selfObject = Indicator;
            var diametrCSS = '';
            var innerText = '';
            if ((Indicator.type == 'simple-indicator')||(Indicator.type == 'complex-indicator')) {
                diametrCSS = " width: "+Indicator.radius*2+"px; height: "+Indicator.radius*2+"px; ";  
            }
            if (Indicator.type == 'string-indicator') {
                innerText = hexDecode(Indicator.content);  
            }
             var curType = getIndicatorStatus(Indicator)||'status-green';
            Indicator.status = curType;
            isVisibleCSS = '';
            if (Indicator.isVisible == 0) {
                isVisibleCSS = 'display: none;';
            }
            $("body").append("<div style='"+isVisibleCSS+"position: absolute; "+diametrCSS+" top: "+Indicator.top+"px; left: "+Indicator.left+"px;' class='"+Indicator.type+" "+curType+"' id='"+Indicator.id+"' >"+innerText+"</div>");
            if (Indicator.type == 'speed-indicator') {
                Indicator.curSpeed = MsgCodeStatus['OperationPerMin'];
                DrawSpeedDigits(Indicator.minSpeed,Indicator.maxSpeed,Indicator.id);
                DrawSpeedArrow(Indicator.minSpeed*1,Indicator.maxSpeed*1,Indicator.curSpeed*1,Indicator.id);      
            }

              
    };
//-------Type: Indicator------------------------------        


    function addObject() {
        var objectType =  $("#selectObjectType").val();
        var objectNumber = $('.'+objectType).length || 0;
        while ($('#'+objectType+'-'+objectNumber).length) {
            objectNumber++;
        }

        var newInditaor = new Indicator(objectType,objectType+'-'+objectNumber);

        IndicatorsList.push(newInditaor);
        updateObjectsInspectorSelect();
        setActiveObject(newInditaor);
        newInditaor.drawIndicator();
        

    }

    function updateObjectsInspectorSelect() {
        $('#selectActiveObject').html('');
        $('#connectedIndicators').html('');
        for(var i=0; i<IndicatorsList.length; i++) {
            $('#selectActiveObject').append('<option value="'+IndicatorsList[i].id+'">'+IndicatorsList[i].id+'</option>');
            if ((IndicatorsList[i].type == 'simple-indicator')||(IndicatorsList[i].type == 'string-indicator'))  {
                $('#connectedIndicators').append('<option value="'+IndicatorsList[i].id+'">'+IndicatorsList[i].id+'</option>');
            }

            
        }
    }

    function setActiveObject(selectedObject) {
        //console.log('setActiveObject');
        $("#selectActiveObject").val(selectedObject.id);
        //console.log(selectedObject.top,selectedObject.left);
        $("#activeObjectY").val(selectedObject.top);
        $("#activeObjectX").val(selectedObject.left);
        
       if (selectedObject.type == 'simple-indicator')   { 
            $("#activeObjectMsgCode").val(selectedObject.MsgCode);
            $("#activeObjectMsgCode").show();
            $("#activeObjectContent").hide();
            $("#connectedIndicators").hide();
        }
       if (selectedObject.type == 'string-indicator')   { 
            $("#activeObjectContent").val(selectedObject.content);
            $("#activeObjectMsgCode").show();
            $("#activeObjectContent").show();
            $("#connectedIndicators").hide();
        }
        if (selectedObject.type == 'complex-indicator')   { 
            $("#connectedIndicators option").prop("selected", false);
            selectedObject.connectedIndicators.forEach(function(item, i, arr) {
                $("#connectedIndicators option[value='" + item + "']").prop("selected", true);
            });
            $("#activeObjectMsgCode").hide();
            $("#activeObjectContent").hide();
            $("#connectedIndicators").show();
        } 
        if (selectedObject.type == 'speed-indicator')   {
            $("#activeObjectMsgCode").hide();
            $("#activeObjectContent").hide();
            $("#connectedIndicators").hide();
        } 
        
        
    }

    function fselectActiveObject() {
        var activeObjectId =  $("#selectActiveObject").val();
        for(var i=0; i<IndicatorsList.length; i++) {
            if (IndicatorsList[i].id==activeObjectId) {
                setActiveObject(IndicatorsList[i]);
                break;
            }
        }
    }

    function moveActiveObject() {
        var activeObjectId = $("#selectActiveObject").val();

        for(var i=0; i<IndicatorsList.length; i++) {
            if (IndicatorsList[i].id==activeObjectId) {
                IndicatorsList[i].top=$("#activeObjectY").val();
                IndicatorsList[i].left=$("#activeObjectX").val();
                break;
            }
        }

        $("#"+activeObjectId).css({
            'left': $("#activeObjectX").val()+'px',
            'top': $("#activeObjectY").val()+'px'
             });

    }

    function changeActiveObject() {
        var activeObjectId = $("#selectActiveObject").val();
        for(var i=0; i<IndicatorsList.length; i++) {
            if (IndicatorsList[i].id==activeObjectId) {
                IndicatorsList[i].MsgCode=$("#activeObjectMsgCode").val();
                IndicatorsList[i].connectedIndicators=$("#connectedIndicators").val(); 
                IndicatorsList[i].content=$("#activeObjectContent").val(); 
                $("#"+IndicatorsList[i].id).html(IndicatorsList[i].content);
                break;
            }
        }       
    }

    function deleteActiveObject() {
        var activeObjectId = $("#selectActiveObject").val();
        for(var i=0; i<IndicatorsList.length; i++) {
            if (IndicatorsList[i].id==activeObjectId) {
                IndicatorsList.splice(i,1);
                 $("#"+activeObjectId).remove();
                updateObjectsInspectorSelect();
                $("#activeObjectY").val('');
                $("#activeObjectX").val('');
                break;
            }
        }

    }

    function saveScreen() {
        r = Math.random();
        $.get('editor.asp',{todo:'saveScreen',
        r:r,
        Indicators: JSON.stringify(IndicatorsList) 
        },function(){
            //location.reload();
        }); 
    
    }

    function BuildScreen() {
        for(var i=0; i<IndicatorsList.length; i++) {
            drawIndicator(IndicatorsList[i]);
            
        }
    }

    var MsgCodeStatus={<%
        '------------------------------------------------------
    '---fill MsgCode list----------------------------------
    list_MsgCodeStatus = ""
    'sqlstr = "select distinct CategoryCode+cast(msgID as nvarchar(10))  MsgCode, MsgTypeName from Messages_Type"
    sqlstr = " select distinct CategoryCode+cast(msgID as nvarchar(10))  MsgCode, case when (([Status] <> 1)AND([Status] <> 3)) then 'status-grey' "
	sqlstr = sqlstr&" when (([Status] = 1)AND(dateadd(MINUTE,[Period],[LastTime])<GETDATE())) then 'status-green' "
	sqlstr = sqlstr&"  when (([Status] = 3)AND(dateadd(MINUTE,1.2*[Period],[LastTime])<GETDATE())) then 'status-green' "
	sqlstr = sqlstr&"  else case when [ErrorLevel] = 1 then 'status-white' "
	sqlstr = sqlstr&" 		   when [ErrorLevel] = 2 then 'status-yellow' "
	sqlstr = sqlstr&" 		   when [ErrorLevel] = 3 then 'status-red' "
	sqlstr = sqlstr&" 	  end "
    sqlstr = sqlstr&"  end status_color  from Messages_Type"
    sqlstr = sqlstr&" union all select 'OperationPerMin', cast(SUM(OPERATION)/10 as nvarchar(15)) from Log_VO where [TIME]=(select top 1 [TIME] from Log_VO order by [TIME] desc)"

    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        do while (not Rs.EOF)
            if (list_MsgCodeStatus="") then
                response.write "'"&Rs.Fields("MsgCode")&"': '"&Rs.Fields("status_color")&"'"
                list_MsgCodeStatus = "1"
            else
                response.write ",'"&Rs.Fields("MsgCode")&"': '"&Rs.Fields("status_color")&"'"
            end if
            Rs.MoveNext
        loop
    end if
    Rs.Close
    
    %>};

    /*-------------------------------------------------------------------------------------------*/
    /*-------------START: Convrte text to HEX----------------------------------------------------*/
    /*-------------------------------------------------------------------------------------------*/
function hexEncode(str){
    var hex, i;

    var result = "";
    for (i=0; i<str.length; i++) {
        hex = str.charCodeAt(i).toString(16);
        result += ("000"+hex).slice(-4);
    }

    return result
}

function hexDecode(str){
    var j;
    var hexes = str.match(/.{1,4}/g) || [];
    var back = "";
    for(j = 0; j<hexes.length; j++) {
        back += String.fromCharCode(parseInt(hexes[j], 16));
    }

    return back;
}
    /*-------------------------------------------------------------------------------------------*/
    /*-------------END: Convrte text to HEX------------------------------------------------------*/
    /*-------------------------------------------------------------------------------------------*/

     $(function() {

    IndicatorsList=<%

        if (cInt(ScreenID)>0) then
            response.write  IndicatorsList
        else
            response.write "ActiveScreen.indicators"
        end if
     %>;
        
    <%  
    '=IndicatorsList
    '=Background
        if (cInt(ScreenID)>0) then
            %> $("body").css("background","url(<%=Background %>) no-repeat #294959");  <%
        else
            %> $("body").css("background","url("+ActiveScreen.bgpath+") no-repeat #294959"); <%
        end if
    %>;
      
    
        BuildScreen();


    });
</script>
</head>
<body>

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
