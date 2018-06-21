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
    '---fill MsgCode list----------------------------------
    list_MsgCode = ""
    sqlstr = "select distinct CategoryCode+cast(msgID as nvarchar(10))  MsgCode, MsgTypeName from Messages_Type"
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        do while (not Rs.EOF)
            list_MsgCode = list_MsgCode & "<option value='"&Rs.Fields("MsgCode")&"' >"&Rs.Fields("MsgCode")&" - "&Rs.Fields("MsgTypeName")&"</option>"
            Rs.MoveNext
        loop
    end if
    Rs.Close

    '------------------------------------------------------
    '---fill ScreensId list----------------------------------
    list_ScreensId = ""
    sqlstr = "SELECT [id],[background_path],[indicators_list],isnull([name],cast([id] as nvarchar(10))) name from Screen_List"
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        list_ScreensId = "<select  style='width: 200px' id='dialogSelectId' >"
        do while (not Rs.EOF)
            'response.write "<option value='"&Rs.Fields("id")&"' >"&Rs.Fields("name")&"</option>"
            list_ScreensId = list_ScreensId&"<option value='"&Rs.Fields("id")&"' >"&Rs.Fields("name")&"</option>"
            Rs.MoveNext
        loop
        list_ScreensId = list_ScreensId&"</select>"
    else 
        list_ScreensId = "<span>Доступных эканов не обнаружено</span>"
    end if
    Rs.Close

    'response.write list_ScreensId
    'response.end


function saveScreen()
    Indicators = Request("Indicators")
    ScreenID=Request("screenid")
    BackgroundPath=Request("bgpath")
    ScreenName=Request("screenname")
    'ScreenRefresh=Request("refresh")
	sqlstr = "EXEC sp_SaveScreen @Indicators='"&Indicators&"', @Background='"&BackgroundPath&"', @Name='"&ScreenName&"' "
    ', @Refresh="&ScreenRefresh&" "
    if (ScreenID>0) then
        sqlstr = sqlstr&" ,@ID="&ScreenID&" "
    end if
    'response.write Request("Indicators")
 	Rs.Open sqlstr, Conn

end function

function deleteScreen()
    ScreenID=Request("screenid")
    if (ScreenID>0) then
    	sqlstr = "DELETE FROM Screen_List WHERE [id]="&ScreenID&" "
        Rs.Open sqlstr, Conn
    end if
end function

function saveScreenXML()
    Indicators = Request("Indicators")
    ScreenID=Request("screenid")
    if not (ScreenID>0) then
        ScreenID = 0
    end if
    'BackgroundPath=Request("bgpath")
    'ScreenName=Request("screenname")
    Set objFSO=CreateObject("Scripting.FileSystemObject")


   strFolder = Server.MapPath(".")

    ' How to write file
    outFile=strFolder&"\XML\screen"&ScreenID&".xml"

    Set objFile = objFSO.CreateTextFile(outFile,True)
    objFile.Write Indicators & vbCrLf
    objFile.Close


end function

function saveScreenAsActive()
    Indicators = Request("Indicators")
    ScreenID=Request("screenid")
    'if not (ScreenID>0) then
    '    ScreenID = 0
    'end if
    'BackgroundPath=Request("bgpath")
    'ScreenName=Request("screenname")
    Set objFSO=CreateObject("Scripting.FileSystemObject")


   strFolder = Server.MapPath(".")

    ' How to write file
    outFile=strFolder&"\main_variant_"&ScreenID&".js"

    Set objFile = objFSO.CreateTextFile(outFile,True)
    objFile.Write Indicators & vbCrLf
    objFile.Close
end function


function loadScreen()
    IndicatorsList = ""
    ScreenID=Request("screenid")
    if isnumeric(ScreenID) then
           IndicatorsList = ""
            sqlstr = "select * from Screen_List where id="&ScreenID
            Rs.Open sqlstr, Conn
            If not Rs.EOF then
                    IndicatorsList = IndicatorsList &"{ ""Name"": """& Rs.Fields("name")&""", ""Background"": """& Rs.Fields("background_path")&""", ""Indicators"": "& Rs.Fields("indicators_list")&"}"
            end if
            Rs.Close
    end if 
    Response.write IndicatorsList   
end function

function getScreenList()
    list_ScreensId = ""
    sqlstr = "SELECT [id],[background_path],[indicators_list],isnull([name],cast([id] as nvarchar(10))) name from Screen_List"
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        list_ScreensId = "<select  style='width: 200px' id='dialogSelectId' >"
        do while (not Rs.EOF)
            'response.write "<option value='"&Rs.Fields("id")&"' >"&Rs.Fields("name")&"</option>"
            list_ScreensId = list_ScreensId&"<option value='"&Rs.Fields("id")&"' >("&Rs.Fields("id")&") "&Rs.Fields("name")&"</option>"
            Rs.MoveNext
        loop
        list_ScreensId = list_ScreensId&"</select>"
    else 
        list_ScreensId = "<span>Доступных эканов не обнаружено</span>"
    end if
    Rs.Close
    Response.write list_ScreensId
end function

function getIndicatorStatus()
    msgcode=Request("msgcode")
    status_color=""
    sqlstr = " select case when (([Status] <> 1)AND([Status] <> 3)) then 'status-grey' "
	sqlstr = sqlstr&" when (([Status] = 1)AND(dateadd(MINUTE,[Period],[LastTime])<GETDATE())) then 'status-green' "
	sqlstr = sqlstr&"  when (([Status] = 3)AND(dateadd(MINUTE,1.2*[Period],[LastTime])<GETDATE())) then 'status-green' "
	sqlstr = sqlstr&"  else case when [ErrorLevel] = 1 then 'status-white' "
	sqlstr = sqlstr&" 		   when [ErrorLevel] = 2 then 'status-yellow' "
	sqlstr = sqlstr&" 		   when [ErrorLevel] = 3 then 'status-red' "
	sqlstr = sqlstr&" 	  end "
    sqlstr = sqlstr&"  end status_color  from Messages_Type  where CategoryCode+cast(msgID as nvarchar(10)) ='"&msgcode&"' "
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        status_color = Rs.Fields("status_color")
    end if
    Rs.Close
    Response.write status_color
end function

function getOperationPerMin() 
    sqlstr = "select SUM(OPERATION)/10 OperationPerMin from Log_VO where [TIME]=(select top 1 [TIME] from Log_VO order by [TIME] desc)"
    Rs.Open sqlstr, Conn
    If not Rs.EOF then
        OperationPerMin = Rs.Fields("OperationPerMin")
    end if
    Rs.Close
    Response.write OperationPerMin
end function

function saveScreenRefresh()
    newRefresh = Request("refresh")
    ScreenID=Request("screenid")
    if isnumeric(ScreenID) then
        sqlstr = "if not exists(select  * from Screen_Config where screenID is not NULL) "
        sqlstr = sqlstr&" insert Screen_Config (refresh,screenID) values ("&newRefresh&","&ScreenID&") "
        sqlstr = sqlstr&" else if not exists(select  * from Screen_Config where screenID="&ScreenID&") "
        sqlstr = sqlstr&" insert Screen_Config (refresh,screenID) values ("&newRefresh&","&ScreenID&") "
        sqlstr = sqlstr&"else update Screen_Config set refresh="&newRefresh&" where screenID="&ScreenID&" "
        Rs.Open sqlstr, Conn
    end if

end function

    'todo:'saveScreen',Indicators:IndicatorsList
if NOT IsEmpty(Request("todo")) then
	if Request("todo") = "saveScreen" then
		saveScreen() 
    elseif Request("todo") = "deleteScreen" then
		deleteScreen() 
    elseif  Request("todo") = "saveScreenXML" then
		saveScreenXML() 
    elseif  Request("todo") = "loadScreen" then
		loadScreen() 
    elseif  Request("todo") = "getScreenList" then
		getScreenList() 
    elseif  Request("todo") = "getIndicatorStatus" then
		getIndicatorStatus() 
    elseif  Request("todo") = "getOperationPerMin" then
		getOperationPerMin() 
    elseif  Request("todo") = "saveScreenAsActive" then
		saveScreenAsActive() 
    elseif  Request("todo") = "saveScreenRefresh" then
		saveScreenRefresh() 
	end if 
	Response.End
end if	
%>
<!DOCTYPE HTML>
<html>
<head>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
        <meta http-equiv="X-UA-Compatible" content="ie=edge">
		<!-- <meta http-equiv='refresh' content='60; url=http://ufa-qos01ow/vsp/main1.asp'> -->
		<!-- 1. Add these JavaScript inclusions in the head of your page -->
		<title>Редактор экранов</title>
    	<script src="js\jquery-3.2.1.min.js"></script>
        <script src="js\jquery-ui.min.js"></script>
        <script src="js\json2.js"></script>
        <script src="js\json2xml.js"></script>
        <script src="js\xml2json.js"></script>

        <script src="js\jquery-migrate-3.0.0.min.js"></script>
        <link type="text/css" href="js/jquery-ui.min.css" rel="stylesheet" />


<style>
        body {
            background: url(img/screen.png) no-repeat #294959;
            font-family: Arial;
            position: relative;
        }

        .ui-dialog {
            font-size: .8em;
        }

        .tools-panel {
            width: 100%;
            position: fixed;
            top: 0px;
            left: 0px;
            padding-left: 10px;
            padding-top: 10px;
            padding-bottom: 10px;
            background:rgba(0,0,0,0.5);
            z-index: 95;

        }
        .tools-panel>div{
            float: left;
            margin-left: 30px;

        }
        .tools-panel label {
            color: white;
         }
        .tools-panel>button{
            margin-bottom: 5px;

        }
        .tools-panel-switcher {
            position: absolute;
            right: 80px;
            top: 20px;
            color: white;
            padding: 5px;
            cursor: pointer;
            border: 1px solid white;
            -moz-border-radius:  2px; /* R */
            -webkit-border-radius: 2px ; /* R */
            border-radius: 2px; /* R */
            z-index: 99;

        }
        #activeObject>div{
            float: left;
            margin-right: 5px;
        }
        #activeObject span {
            color: white;
         }

        #activeObject input, #activeObject select {
            vertical-align: top;
        }

        #activeObjectX, #activeObjectY, #activeObjectRadius {
            width: 60px;
        }

        #activeObjectContent {
            width: 150px;
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
            -moz-border-radius: 50%; /* R */
            -webkit-border-radius: 50%; /* R */
            border-radius: 50%; /* R */
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
        this.radius = 0;
        this.content = '';
        this.MsgCode = '';
        this.connectedIndicators = [];
        this.isVisible = 1;
        this.style = "position: absolute; top: "+this.top+"px; left: "+this.left+"px;"

        if (indicatorType == 'simple-indicator') {
            this.radius = 5;
        } 
        if (indicatorType == 'complex-indicator') {
            this.radius = 30;
        } 
        if ((indicatorType == 'simple-indicator')||(indicatorType == 'complex-indicator') ) {
            this.status = 'status-green';
            this.class += ' '+this.status;
            this.content = ''
        } 
        if (indicatorType == 'string-indicator') {
            this.status = 'status-green';
            this.class += ' '+this.status;
            this.content = '00530061006d0070006c006500200074006500780074'
        }
        if (indicatorType == 'speed-indicator') {
                this.minSpeed = 0;
                this.maxSpeed = 180;
                this.curSpeed = 0;
                this.content = ''
        }
            
     
    };

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
            $("body").append("<div style='position: absolute; "+diametrCSS+" top: "+Indicator.top+"px; left: "+Indicator.left+"px;' class='"+Indicator.class+"' id='"+Indicator.id+"' >"+innerText+"</div>");
            if (Indicator.type == 'speed-indicator') {
                DrawSpeedDigits(Indicator.minSpeed,Indicator.maxSpeed,Indicator.id);
                DrawSpeedArrow(Indicator.minSpeed,Indicator.maxSpeed,Indicator.curSpeed,Indicator.id);      
            }
            $('#'+Indicator.id).draggable();
            $('#'+Indicator.id).on( "drag", function( event, ui ) {
                selfObject.top = ui.position.top;
                selfObject.left = ui.position.left;
                $("#activeObjectX").val(ui.position.left);
                $("#activeObjectY").val(ui.position.top);
            } );
            $('#'+Indicator.id).on( "mousedown", function( event, ui ) {
               // console.log('mousedown');
                setActiveObject(selfObject);
            } );
            
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
        drawIndicator(newInditaor);
        setActiveObject(newInditaor);
        
        

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

    function updateScreenSelect() {
        r = Math.random();
        $.ajax({
            url: 'editor.asp',
            type: 'GET',
            data: { 
                todo:'getScreenList',
                r:r
                 },
                success: function(result) {
                    $("#selectScreenDialog").html(result);
            }
        });
    }

    function setActiveObject(selectedObject) {
        //console.log('setActiveObject');

        $("#selectActiveObject").val(selectedObject.id);
        //console.log(selectedObject.top,selectedObject.left);
        $("#activeObjectY").val(selectedObject.top);
        $("#activeObjectX").val(selectedObject.left);
        
       if (selectedObject.type == 'simple-indicator')   { 
            $("#activeObjectMsgCode").val(selectedObject.MsgCode);
            $("#divMsgCode").show();
            //var diametr = $("#"+selectedObject.id).css("width");
            //diametr = diametr.slice(0,-2)*1/2;
            $("#activeObjectRadius").val(   selectedObject.radius  );
            $("#divRadius").show();
            $("#divContent").hide();
            $("#divConnected").hide();
            $("#divMin").hide();
            $("#divMax").hide();
            $("#activeisVisible").prop('checked',!!selectedObject.isVisible);
            $("#divisVisible").show();
        }
       if (selectedObject.type == 'string-indicator')   { 
            $("#activeObjectContent").val(hexDecode(selectedObject.content));
            $("#activeObjectMsgCode").val(selectedObject.MsgCode);
            $("#divMsgCode").show();
            $("#divContent").show();
            $("#divRadius").hide();
            $("#divConnected").hide();
            $("#divMin").hide();
            $("#divMax").hide();
            $("#activeisVisible").prop('checked',!!selectedObject.isVisible);
            $("#divisVisible").show();
        }
        if (selectedObject.type == 'complex-indicator')   { 
            $("#connectedIndicators option").prop("selected", false);
            selectedObject.connectedIndicators.forEach(function(item, i, arr) {
                $("#connectedIndicators option[value='" + item + "']").prop("selected", true);
            });
            $("#divMsgCode").hide();
            $("#divContent").hide();
            $("#divConnected").show();
            //var diametr = $("#"+selectedObject.id).css("width");
            //diametr = diametr.slice(0,-2)*1/2;
            $("#activeObjectRadius").val(   selectedObject.radius  );
            //$("#activeObjectRadius").show();
            $("#divRadius").show();
            $("#divMin").hide();
            $("#divMax").hide();
            $("#activeisVisible").prop('checked',!!selectedObject.isVisible);
            $("#divisVisible").show();
        } 
        if (selectedObject.type == 'speed-indicator')   {
            $("#divMsgCode").hide();
            $("#divContent").hide();
            $("#divRadius").hide();
            $("#divConnected").hide();
            $("#activeMin").val(selectedObject.minSpeed );
            $("#activeMax").val(selectedObject.maxSpeed );
            $("#divMin").show();
            $("#divMax").show();
            $("#activeisVisible").prop('checked',!!selectedObject.isVisible);
            $("#divisVisible").show();
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
                if ((IndicatorsList[i].type == 'simple-indicator')||(IndicatorsList[i].type == 'string-indicator'))   {
                    IndicatorsList[i].MsgCode=$("#activeObjectMsgCode").val();
                }
                if (IndicatorsList[i].type == 'string-indicator')   {
                    IndicatorsList[i].content=hexEncode($("#activeObjectContent").val()); 
                    $("#"+IndicatorsList[i].id).html(hexDecode(IndicatorsList[i].content));
                } 
                if ((IndicatorsList[i].type == 'simple-indicator')||(IndicatorsList[i].type == 'complex-indicator'))   {
                    IndicatorsList[i].radius=$("#activeObjectRadius").val();
                    $("#"+IndicatorsList[i].id).css("width",IndicatorsList[i].radius*2+"px");
                    $("#"+IndicatorsList[i].id).css("height",IndicatorsList[i].radius*2+"px");
                }
                if (IndicatorsList[i].type == 'complex-indicator')  {
                    IndicatorsList[i].connectedIndicators=$("#connectedIndicators").val(); 
                }
                if (IndicatorsList[i].type == 'speed-indicator')  {
                    IndicatorsList[i].minSpeed=$("#activeMin").val(); 
                    IndicatorsList[i].maxSpeed=$("#activeMax").val(); 
                }

                IndicatorsList[i].isVisible=+$("#activeisVisible").is(':checked');
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
        var backgroundPath=$("#backgroundPath").val();
        //var screenRefresh=$("#screenRefresh").val();
        var screenName=changeBackslash($("#screenName").val());
        $.ajax({
            url: 'editor.asp',
            type: 'POST',
            data: { 
                todo:'saveScreen',
                bgpath:backgroundPath,
                //refresh: screenRefresh,
                screenname:screenName,
                screenid: screenId,
                r:r,
                Indicators: JSON.stringify(IndicatorsList) 
                 },
            success: function(result) {
                updateScreenSelect();
                alert('Экран сохранен');
                //location.reload();
            }
        });
  
    }

    function deleteScreen() {
        r = Math.random();
        $.ajax({
            url: 'editor.asp',
            type: 'POST',
            data: { 
                todo:'deleteScreen',
                screenid: screenId,
                r:r
                 },
            success: function(result) {
                updateScreenSelect();
                newScreen();
                alert('Экран удален');
            }
        });       
    }

    function saveScreenAsActive() {
        var r = Math.random();
        var xml2 = json2xml(IndicatorsList);
        var backgroundPath=changeBackslash($("#backgroundPath").val());
        //var screenRefresh=$("#screenRefresh").val();
        var screenName=$("#screenName").val();

     var text = 
    'ActiveScreen = {"screenid":'+screenId+','+
    '"bgpath":"'+backgroundPath+'",'+
    '"screenname":"'+screenName+'",'+
    '"indicators":'+ JSON.stringify(IndicatorsList)+'}';
      
    // '"refresh":"'+screenRefresh+'",'+

        $.ajax({
            url: 'editor.asp',
            type: 'POST',
            data: { 
                todo:'saveScreenAsActive',
                bgpath:backgroundPath,
                screenname:screenName,
                screenid: screenId,
                r:r,
                Indicators: text
                 },
            success: function(result) {
                //updateScreenSelect();
                alert('Экран сохранен как main_variant_'+screenId+'.js');
                //location.reload();
            }
        });
         
    }

    function saveScreenXML() {
        var r = Math.random();
        var xml2 = json2xml(IndicatorsList);
        var backgroundPath=changeBackslash($("#backgroundPath").val());
        //var screenRefresh=$("#screenRefresh").val();
        var screenName=$("#screenName").val();

     var text = 
    '<screenid>'+screenId+'</screenid>'+
    '<bgpath>'+backgroundPath+'</bgpath>'+
    '<screenname>'+screenName+'</screenname>'+
    '<indicators>'+ JSON.stringify(IndicatorsList)+'</indicators>';
    
    //'<refresh>'+screenRefresh+'</refresh>'+
      
        $.ajax({
            url: 'editor.asp',
            type: 'POST',
            data: { 
                todo:'saveScreenXML',
                bgpath:backgroundPath,
                screenname:screenName,
                screenid: screenId,
                r:r,
                Indicators: text
                 },
            success: function(result) {
                //updateScreenSelect();
                alert('Экран сохранен как screen'+screenId+'.xml');
                //location.reload();
            }
        });
         
    }

    function saveScreenParams() {
        var tempBG = $("#backgroundPath").val();
        tempBG = changeBackslash(tempBG);
        $("#backgroundPath").val(tempBG);
        $("body").css("background","url("+tempBG+") no-repeat #294959");
        
    }

    function loadScreenXML() { 
        document.getElementById('XMLfileInput').click();
     
    }

    function loadScreenXMLhandler() {    
        if (!window.File || !window.FileReader || !window.FileList || !window.Blob) {
          alert('Ваш браузер не поддерживает работу с файлами.');
          return;
        }   

        input = document.getElementById('XMLfileInput');
        if (!input) {
          alert("Сбой ввода файла.");
        }
        else if (!input.files) {
          alert("Ваш браузер не поддерживает загруку файлов.");
        }
        else if (!input.files[0]) {
          alert("Файл не выбран.");               
        }
        else {
          file = input.files[0];
          fr = new FileReader();
          fr.onload = receivedText;
          fr.readAsText(file);
        }
  }

  function receivedText() {
    var text = fr.result;
    parseXml(text);

  } 

    function parseXml(xml) {
        var regexpBG = /<bgpath>(.*)<\/bgpath>/ig;
        var regexpName = /<screenname>(.*)<\/screenname>/ig;
        var regexpId = /<screenid>(.*)<\/screenid>/ig;
        var regexpIndicators = /<indicators>(.*)<\/indicators>/ig;
        var regexpRefresh = /<refresh>(.*)<\/refresh>/ig;
        var result = regexpBG.exec(xml);
        $("#backgroundPath").val(result[1]);
        result = regexpName.exec(xml);
        $("#screenName").val(result[1]);
        //result = regexpRefresh.exec(xml);
        //$("#screenRefresh").val(result[1]);

        result = regexpId.exec(xml);
        screenId=result[1]*1;
        result = regexpIndicators.exec(xml);
        $(".simple-indicator, .complex-indicator, .string-indicator, .speed-indicator").remove();
        var tempRes=$.parseJSON(result[1]);

        IndicatorsList=tempRes;
        BuildScreen();

    }

  
    function loadScreenDB() {
        $("#selectScreenDialog").dialog("open");
        
    }

    function loadScreenDBbyID() {
        var selectedId=$("#dialogSelectId").val();
        r = Math.random();
        $.ajax({
            url: 'editor.asp',
            type: 'GET',
            data: { 
                todo:'loadScreen',
                screenid:selectedId, 
                r:r
                 },
            success: function(result) {
                newScreen();
                screenId=selectedId;
                $(".simple-indicator, .complex-indicator, .string-indicator, .speed-indicator").remove();
                var tempRes=$.parseJSON(result);
                //console.log(tempRes);
                $("#screenName").val(tempRes.Name);
                $("#backgroundPath").val(tempRes.Background);
                //$("#screenRefresh").val(tempRes.refresh);
                $("body").css("background","url("+tempRes.Background+") no-repeat #294959");

                IndicatorsList=tempRes.Indicators;
                BuildScreen();
            }
        });

    }

    function changeBackslash(str) {
       return str.replace(/\\/g, "/");
    }

    function newScreen() {
       // $("#activeObjectMsgCode").hide();
       // $("#activeObjectRadius").hide();
       // $("#connectedIndicators").hide();
       // $("#activeObjectContent").hide();
        $("#divMsgCode").hide();
        $("#divRadius").hide();
        $("#divConnected").hide();
        $("#divContent").hide();
        $("#divMin").hide();
        $("#divMax").hide();
        $("#divisVisible").hide();
        $(".simple-indicator, .complex-indicator, .string-indicator, .speed-indicator").remove();
        $('#selectActiveObject').html('');
        $('#connectedIndicators').html('');
        IndicatorsList=[];
        screenId=0;
    }

    function BuildScreen() {
        for(var i=0; i<IndicatorsList.length; i++) {
            drawIndicator(IndicatorsList[i]);
            
        }
        updateObjectsInspectorSelect();
    }    

    function switchPanel() {
        var visibiliti = $(".tools-panel").is(":visible");
       if (visibiliti) {
            $(".tools-panel").hide();
            $(".tools-panel-switcher").html("Показать панель");
    
        } else {
            $(".tools-panel").show();
            $(".tools-panel-switcher").html("Свернуть панель");
    
        }
    }

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


    var screenId = 0;

    $(function() {
        updateScreenSelect();
        newScreen();        

        $("#selectScreenDialog").dialog({ title: "Выбор экрана", width: 250, modal: true, autoOpen: false });
        $("#selectScreenDialog").dialog("option",
            "buttons", {
                "Отмена": function () { $(this).dialog("close"); },
                "Загрузить": function () {
                    loadScreenDBbyID();
                    $(this).dialog("close");
                }
            });

        $("#XMLfileInput").change(function() {
            loadScreenXMLhandler();
        });

        $("#activeObjectY").on( "change", function( event, ui ) {  moveActiveObject();  });
        $("#activeObjectX").on( "change", function( event, ui ) {  moveActiveObject();  });

        $(document).on( "keydown", function( event, ui ) {
            var code = (event.keyCode ? event.keyCode : event.which);
            var activeObjectId = $("#selectActiveObject").val();
            if  ((activeObjectId != '')&&(! $("#activeObjectX").is( ":focus" ))&&(! $("#activeObjectY").is( ":focus" )))  {
                if (code == 40) {
                    //console.log("down pressed");
                    $("#activeObjectY").val(1*$("#activeObjectY").val()+5);
                } else if (code == 38) {
                    //console.log("up pressed");
                    $("#activeObjectY").val(1*$("#activeObjectY").val()-5);
                }  else if (code == 37) {
                    //console.log("up pressed");
                    $("#activeObjectX").val(1*$("#activeObjectX").val()-5);
                }   else if (code == 39) {
                    //console.log("up pressed");
                    $("#activeObjectX").val(1*$("#activeObjectX").val()+5);
                } 
                moveActiveObject();


            }
            
        } );

    });
</script>
</head>
<body>
    <div id="selectScreenDialog"></div>
    <div class="tools-panel">
        <button onClick="newScreen()">Создать новый экран</button>
        <button onClick="loadScreenDB()">Загрузить экран из БД</button>
        <input id="XMLfileInput" type="file" style="display:none;" />
        <button onClick="loadScreenXML()">Загрузить экран из XML</button>
        <button onClick="saveScreen()">Сохранить экран в БД</button>
        <button onClick="saveScreenXML()">Сохранить экран в XML</button>
        <button onClick="saveScreenAsActive()">Сохранить экран в JS</button>
        <button onClick="deleteScreen()">Удалить экран</button>
        <br>
        <label for="screenName">Название</label>
        <input type="text" id="screenName" value="Undefined">  
        <!-- <label for="screenRefresh">Частота</label>
        <input type="text" id="screenRefresh" value="600">   -->
        <label for="screenName">Путь к фоновому изображению</label>
        <input type="text" id="backgroundPath" value="img\screen.png">   
        <button onClick="saveScreenParams()">Установить</button>
        <br>
        <select id="selectObjectType">
            <option value="simple-indicator">Простой индикатор</option>
            <option value="complex-indicator">Составной индикатор</option>
            <option value="string-indicator">Текстовый индикатор</option>
            <option value="speed-indicator">Стрелочный индикатор</option>
        </select>
        <button onClick="addObject()">Создать</button>
        <br>
        <form id="activeObject" name="activeObject">
            <div>
                <select id="selectActiveObject" onchange="fselectActiveObject()" ></select>
            </div>
            <div>
                <span>Координаты</span><br>
                <input type="text" id="activeObjectX" value="0">
                <input type="text" id="activeObjectY" value="0">
                <!-- <input type="button" onclick="moveActiveObject()" value="Переместить"> -->
            </div>
            <!--  divContent, divRadius, divMsgCode, divConnected, divMin, divMax, divisVisible -->
                <div id="divContent">
                    <span>Текст</span><br>
                    <input type="text" id="activeObjectContent" value="">
                </div>
                <div id="divRadius">
                    <span>Радиус</span><br>
                    <input type="text" id="activeObjectRadius" value="">
                </div>
                <div id="divMsgCode">
                    <span>Код индикатора</span><br>
                    <select id="activeObjectMsgCode"  ><%=list_MsgCode  %></select>
                </div>
                <div id="divConnected">
                    <span>Связанные индикаторы</span><br>
                    <select class="select" id="connectedIndicators"  multiple="" tabindex="3">
                        <option value=""></option>
                      </select>
                </div>
                <div id="divMin">
                    <span>Минимум</span><br>
                    <input type="text" id="activeMin" value="">
                </div>
                <div id="divMax">
                    <span>Максимум</span><br>
                    <input type="text" id="activeMax" value="">
                </div>
                <div id="divisVisible">
                    <span>Видимый</span><br>
                    <input type="checkbox" id="activeisVisible" checked="checked">
                </div>
            
            <div>
                <input type="button" onclick="changeActiveObject()" value="Сохранить изменения">
            </div>
            <div>
                <input type="button" onclick="deleteActiveObject()" value="Удалить">
            </div>
        </form>

    </div>
    <div class="tools-panel-switcher" onclick="switchPanel()" >Свернуть панель</div>
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
