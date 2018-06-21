<%@ language="VBScript"%><!-- #include file="const.asp" --><!-- #include file="common.asp" -->
<%
' Редактирование справочников

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
	else
		' Права администратора
%>
<!DOCTYPE HTML>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="ie=edge">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<script type="text/javascript" src="js/jquery.min.js"></script>
<script type="text/javascript">
function ChangeVoc(v) {
	jQuery.get('dataset.asp', { ds: 'Voc', prm: v }, function(ds) {
	  $('#VocCont').html(ds);
	});
	document.getElementById("User").style.display="none";
	document.getElementById("Hist").style.display="none";
	document.getElementById("Fin").style.display="none";
	document.getElementById("RC").style.display="none";
	document.getElementById("Tag").style.display="none";
	$("#ChannelGroup").hide();
	document.getElementById("buffercolor").value=0;
}
$(document).ready(function() {
	ChangeVoc(1);
});
function NewUser(id) {
    oldID = document.getElementById("buffercolor").value;
	if (oldID!='') {
      document.getElementById("r"+oldID).style.backgroundColor="#ffffff";
	}
	document.getElementById("buffercolor").value="";
	document.getElementById("usr_name").value="";
	document.getElementById("usr_login").value="";
	document.getElementById("usr_role").value=0;
	document.getElementById("usr_phone").value="";
	document.getElementById("usr_email").value="";
	document.getElementById("User").style.display="block";
	document.getElementById("Hist").style.display="none";
}
function EditUser(id) {
	jQuery.get('dataset.asp', { ds: 'GetUserProp', prm: id }, function(ds) {
		ds = ds.split('^');
		document.getElementById("usr_name").value=ds[1];
		document.getElementById("usr_login").value=ds[2];
		document.getElementById("usr_role").value=ds[3];
		document.getElementById("usr_phone").value=ds[4];
		document.getElementById("usr_email").value=ds[5];
		document.getElementById("User").style.display="block";
		document.getElementById("Hist").style.display="none";
	});
}
function SaveUser() {
	var p1=document.getElementById("buffercolor").value;
	var p2=document.getElementById("usr_name").value;
	var p3=document.getElementById("usr_login").value;
	var p4=document.getElementById("usr_role").value;
	var p5=document.getElementById("usr_phone").value;
	var p6=document.getElementById("usr_email").value;
	jQuery.get('save.asp', { ds: 'SetUserProp', prm1: p1, prm2: p2, prm3: p3, prm4: p4, prm5: p5, prm6: p6}, function(ds) {
		document.getElementById("User").style.display="none";
		ChangeVoc(3);
		alert('Изменения внесены');
	});
}

function DelUser(id, userlogin) {
	if (confirm('Логин '+userlogin+'. Удалить?')) {
	jQuery.get('save.asp', { ds: 'DelUser', prm1: id}, function(ds) {
		document.getElementById("User").style.display="none";
		ChangeVoc(3);
	});
	}
}

function ViewAudit(id) {
	jQuery.get('dataset.asp', { ds: 'GetUserHist', prm: id }, function(ds) {
		$('#Hist').html(ds);
		document.getElementById("User").style.display="none";
		document.getElementById("Hist").style.display="block";
	});
}

function NewFin(id) {
    oldID = document.getElementById("buffercolor").value;
	if (oldID!='') {
      document.getElementById("r"+oldID).style.backgroundColor="#ffffff";
	}
	document.getElementById("buffercolor").value="";
	document.getElementById("fin_code").value="";
	document.getElementById("fin_code").disabled=false;
	document.getElementById("fin_name").value="";
	document.getElementById("Fin").style.display="block";
}
function EditFin(id) {
	jQuery.get('dataset.asp', { ds: 'GetFin', prm: id }, function(ds) {
		ds = ds.split('^');
		document.getElementById("fin_code").disabled=true;
		document.getElementById("fin_code").value=ds[0];
		document.getElementById("fin_name").value=ds[1];
		document.getElementById("Fin").style.display="block";
	});
}
function SaveFin() {
	var p1=document.getElementById("buffercolor").value;
	var p2=document.getElementById("fin_code").value;
	var p3=document.getElementById("fin_name").value;
	jQuery.get('save.asp', { ds: 'SetFin', prm1: p1, prm2: p2, prm3: p3}, function(ds) {
		document.getElementById("Fin").style.display="none";
		ChangeVoc(1);
		alert('Изменения внесены');
	});
}
function DelFin(id) {
	if (confirm('Код '+id+'. Удалить?')) {
	jQuery.get('save.asp', { ds: 'DelFin', prm1: id}, function(ds) {
		document.getElementById("Fin").style.display="none";
		ChangeVoc(1);
	});
	}
}

function NewRC(id) {
    oldID = document.getElementById("buffercolor").value;
	if (oldID!='') {
      document.getElementById("r"+oldID).style.backgroundColor="#ffffff";
	}
	document.getElementById("buffercolor").value="";
	document.getElementById("rc_code").value="";
	document.getElementById("rc_code").disabled=false;
	document.getElementById("rc_name").value="";
	document.getElementById("rc_err").value="0";
	document.getElementById("RC").style.display="block";
}
function EditRC(id) {
	jQuery.get('dataset.asp', { ds: 'GetRC', prm: id }, function(ds) {
		ds = ds.split('^');
		document.getElementById("rc_code").disabled=true;
		document.getElementById("rc_code").value=ds[0];
		document.getElementById("rc_name").value=ds[1];
		document.getElementById("rc_err").value=ds[2];
		document.getElementById("RC").style.display="block";
	});
}
function SaveRC() {
	var p1=document.getElementById("buffercolor").value;
	var p2=document.getElementById("rc_code").value;
	var p3=document.getElementById("rc_name").value;
	var p4=document.getElementById("rc_err").value;
	jQuery.get('save.asp', { ds: 'SetRC', prm1: p1, prm2: p2, prm3: p3, prm4: p4}, function(ds) {
		document.getElementById("RC").style.display="none";
		ChangeVoc(2);
		alert('Изменения внесены');
	});
}
function DelRC(id) {
	if (confirm('Код '+id+'. Удалить?')) {
	jQuery.get('save.asp', { ds: 'DelRC', prm1: id}, function(ds) {
		document.getElementById("RC").style.display="none";
		ChangeVoc(2);
	});
	}
}

function NewTag(id) {
    oldID = document.getElementById("buffercolor").value;
	if (oldID!='') {
      document.getElementById("r"+oldID).style.backgroundColor="#ffffff";
	}
	document.getElementById("buffercolor").value="";
	document.getElementById("tagid").value="";
	document.getElementById("tagid").disabled=false;
	document.getElementById("tagname").value="";
	document.getElementById("sethi").value=0;
	document.getElementById("sethihi").value=0;
	document.getElementById("fileid").value="";
	document.getElementById("groupname").value="";
	document.getElementById("Prop_Crit").value=0;
	document.getElementById("Prop_Active").value=0;
	document.getElementById("Prop_SignOn").value=0;
	document.getElementById("Prop_Time").value=0;
	document.getElementById("Tag").style.display="block";
}
function EditTag(id) {
	jQuery.get('dataset.asp', { ds: 'GetTag', prm: id }, function(ds) {
		ds = ds.split('^');
		document.getElementById("tagid").value=ds[0];
		document.getElementById("tagid").disabled=true;
		document.getElementById("tagname").value=ds[1];
		document.getElementById("sethi").value=ds[2];
		document.getElementById("sethihi").value=ds[3];
		document.getElementById("fileid").value=ds[4];
		document.getElementById("groupname").value=ds[5];
		document.getElementById("Prop_Crit").value=ds[6];
		document.getElementById("Prop_Active").value=ds[7];
		document.getElementById("Prop_SignOn").value=ds[8];
		document.getElementById("Prop_Time").value=ds[9];
		document.getElementById("Tag").style.display="block";
	});
}
function DelTag(id) {
	if (confirm('Параметр '+id+'. Удалить?')) {
	jQuery.get('save.asp', { ds: 'DelTag', prm1: id}, function(ds) {
		document.getElementById("Tag").style.display="none";
		ChangeVoc(4);
	});
	}
}
function SaveTag() {
	var p1=document.getElementById("buffercolor").value;
	var p2=document.getElementById("tagid").value;
	var p3=document.getElementById("tagname").value;
	var p4=document.getElementById("sethi").value;
	var p5=document.getElementById("sethihi").value;
	var p6=document.getElementById("fileid").value;
	var p7=document.getElementById("groupname").value;
	var p8=document.getElementById("Prop_Crit").value;
	var p9=document.getElementById("Prop_Active").value;
	var p10=document.getElementById("Prop_SignOn").value;
	var p11=document.getElementById("Prop_Time").value;
	jQuery.get('save.asp', { ds: 'SetTag', prm1: p1, prm2: p2, prm3: p3, prm4: p4, prm5: p5, prm6: p6, prm7: p7, prm8: p8, prm9: p9, prm10: p10, prm11: p11}, function(ds) {
		document.getElementById("Tag").style.display="none";
		ChangeVoc(3);
		alert('Изменения внесены');
	});
}
//-----------Channel Groups--------------------------------------------------------------
/*function NewChannelGroup(id) {
    oldID = document.getElementById("buffercolor").value;
	if (oldID!='') {
      document.getElementById("r"+oldID).style.backgroundColor="#ffffff";
	}
	document.getElementById("buffercolor").value="";
	document.getElementById("tagid").value="";
	document.getElementById("tagid").disabled=false;
	document.getElementById("tagname").value="";
	document.getElementById("sethi").value=0;
	document.getElementById("sethihi").value=0;
	document.getElementById("fileid").value="";
	document.getElementById("groupname").value="";
	document.getElementById("Prop_Crit").value=0;
	document.getElementById("Prop_Active").value=0;
	document.getElementById("Prop_SignOn").value=0;
	document.getElementById("Prop_Time").value=0;
	document.getElementById("Tag").style.display="block";
}*/

function padLeadingZero(num, size) {
    var s = "000000000" + num;
    return s.substr(s.length-size);
}

function TimeToMinutes(time_str) {
    var tmp = time_str.split(":");
    return (tmp[0] * 1 * 60 + tmp[1] * 1);
}

function MinutesToTime(minutes_val) {
    var tmp='';
    if (minutes_val>0) {
        if (Math.floor(minutes_val/60)>0) {
            tmp = padLeadingZero(Math.floor(minutes_val/60),2)+':'+padLeadingZero(minutes_val-Math.floor(minutes_val/60)*60,2);
        } else {
            tmp = '00:'+padLeadingZero(minutes_val,2);
        }
    } else {
        tmp = '00:00';
    }
    return tmp;
}

function EditChannelGroup(id) {
	jQuery.get('dataset.asp', { ds: 'GetChannelGroup', prm: id }, function(data) {
		ds = JSON.parse(data);
		$("#cg_ftype").val(ds.cg_ftype);
		$("#cg_group").val(ds.cg_group);
		$("#cg_channel").val(ds.cg_channel);
		$("#cg_warning").val(ds.cg_warning);
		$("#cg_error").val(ds.cg_error);
		$("#cg_minimal").val(ds.cg_minimal);

		$("#cg_limit").val(ds.cg_limit);
		$("#cg_lowactivity_start").val(MinutesToTime(ds.cg_lowactivity_start));
		$("#cg_lowactivity_end").val(MinutesToTime(ds.cg_lowactivity_end));

		$("#ChannelGroup").show();
	});
}
/*function DelChannelGroup(id) {
	if (confirm('Параметр '+id+'. Удалить?')) {
	jQuery.get('save.asp', { ds: 'DelChannelGroup', prm1: id}, function(ds) {
		document.getElementById("ChannelGroup").style.display="none";
		ChangeVoc(4);
	});
	}
}*/
function SaveChannelGroup() {
	var p1=$("#cg_group").val();
	var p2=$("#cg_warning").val();
	var p3=$("#cg_error").val();
	var p4 = $("#cg_minimal").val();

	var p5 = $("#cg_limit").val();

	if ((/(\d\d:\d\d)/.test($("#cg_lowactivity_start").val())) && (/(\d\d:\d\d)/.test($("#cg_lowactivity_end").val()))) {
	    var p6 = TimeToMinutes($("#cg_lowactivity_start").val());
	    var p7 = TimeToMinutes($("#cg_lowactivity_end").val());

	    if (((p6 != 0) || (p7 != 0)) && (p6 > p7)) {
	        alert('Границы интервала низкой активности не корректны.');
	    } else {
	        jQuery.get('dataset.asp', {
	            ds: 'SaveChannelGroup', cg_group: p1, cg_warning: p2, cg_error: p3, cg_minimal: p4,
	            cg_limit: p5, cg_lowactivity_start: p6, cg_lowactivity_end: p7
	        }, function (data) {
	            $("#ChannelGroup").hide();
	            ChangeVoc(5);
	            alert('Изменения внесены');
	        });
	    }
	} else {
	    alert('Не коректный формат границы интервала низкой активности.');
	}



}

//---------------------------------------------------------------------------------

function selectRow(newID)
{
    // обесцвечиваем старый
    var oldID
    oldID = document.getElementById("buffercolor").value;
	if (oldID!='') {
      document.getElementById("r"+oldID).style.backgroundColor="#ffffff";
	}
    // подсвечиваем новый
    document.getElementById("r"+newID).style.backgroundColor="#ffd9ba";
    document.getElementById("buffercolor").value = newID;
	
	document.getElementById("User").style.display="none";
	document.getElementById("Hist").style.display="none";
}
</script>
<style type="text/css">
<!--
BODY {
	margin: 20px;
	font-family: Verdana, Arial, helvetica, sans-serif, Geneva;
	font-size: 10pt;
}
TABLE {
	margin-bottom: 0px;
	margin-top: 0px;
	border-top: solid 1px #CCCCCC;
	border-left: solid 1px #CCCCCC;
}
TD {
	font-family: Verdana, Arial, helvetica, sans-serif, Geneva;
	border-bottom: solid 1px #CCCCCC;
	border-right: solid 1px #CCCCCC;
	padding-left: 2px;
	padding-right: 2px;
	font-size: 9pt;
}
TH {
	border-right: solid 1px black;
	border-bottom: solid 1px black;
	font-size: 8pt;
	font-weight: 600;
	background-color: #6886BA;
	color: #FFFFFF;
	padding-left: 2px;
	padding-right: 2px;
}
-->
</style>
</head>
<body>
<div style="margin-bottom: 16px">
Выбор справочника 
<select name="voc" onchange="javascript: ChangeVoc(this.value);">
<option value="1" selected>Финансовые институты</option>
<option value="2">Response Code</option>
<option value="3">Пользователи</option>
<option value="4">Контролируемые параметры</option>
<option value="5">Контролируемые параметры каналов</option>
</select>
</div>
<div id="VocCont" style="margin-bottom: 16px;"></div>

<div id="User" style="display: none; float: left;">
<table>
<tr><th colspan=2>Карточка пользователя:</th></tr>
<tr><td>Фамилия И.О.</td><td><input type=text size=40 name="usr_name" id="usr_name" maxlength="80" value=""/></td></tr>
<tr><td>Логин</td><td><input type=text size=40 name="usr_login" id="usr_login" maxlength="80" value=""/></td></tr>
<tr><td>Роль</td><td><select name="usr_role" id="usr_role"><option value="0">пользователь</option><option value="1">администратор</option></select>
<tr><td>Телефон</td><td><input type=text size=40 name="usr_phone" id="usr_phone" maxlength="10" value=""/></td></tr>
<tr><td>Эл.почта</td><td><input type=text size=40 name="usr_email" id="usr_email" maxlength="100" value=""/></td></tr>
<tr><td colspan=2><input type="button" onclick="SaveUser();" value="Сохранить" /></td></tr>
</table>
</div>
<div id="Hist" style="display: none; float: left;">
</div>

<div id="Fin" style="display: none; float: left;">
<table>
<tr><th colspan=2>Редактирование:</th></tr>
<tr><td>Код</td><td><input type=text size=40 name="fin_code" id="fin_code" maxlength="4" value=""/></td></tr>
<tr><td>Финансовый институт</td><td><input type=text size=40 name="fin_name" id="fin_name" maxlength="50" value=""/></td></tr>
<tr><td colspan=2><input type="button" onclick="SaveFin();" value="Сохранить" /></td></tr>
</table></div>

<div id="RC" style="display: none; float: left;">
<table>
<tr><th colspan=3>Редактирование:</th></tr>
<tr><td>RC Код</td><td><input type=text size=40 name="rc_code" id="rc_code" maxlength="3" value=""/></td></tr>
<tr><td>Описание</td><td><input type=text size=40 name="rc_name" id="rc_name" maxlength="50" value=""/></td></tr>
<tr><td>Критичность</td><td><input type=text size=40 name="rc_err" id="rc_err" maxlength="1" value=""/></td></tr>
<tr><td colspan=3><input type="button" onclick="SaveRC();" value="Сохранить" /></td></tr>
</table></div>

<div id="Tag" style="display: none; float: left;">
<table>
<tr><th colspan=3>Редактирование:</th></tr>
<tr><td>ID параметра</td><td><input type=text size=50 name="tagid" id="tagid" maxlength="50" value=""/></td></tr>
<tr><td>Параметр</td><td><input type=text size=50 name="tagname" id="tagname" maxlength="255" value=""/></td></tr>
<tr><td>Допустимое значение</td><td><input type=text size=50 name="sethi" id="sethi" value="0"/></td></tr>
<tr><td>Критичное значение</td><td><input type=text size=50 name="sethihi" id="sethihi" value="0"/></td></tr>
<tr><td>Тип файла</td><td><input type=text size=2 name="fileid" id="fileid" maxlength="2" value=""/></td></tr>
<tr><td>Группа</td><td><input type=text size=50 name="groupname" id="groupname" maxlength="50" value=""/></td></tr>
<tr><td>Критичность</td><td><input type=text size=50 name="Prop_Crit" id="Prop_Crit" value="0"/></td></tr>
<tr><td>Активность</td><td><input type=text size=50 name="Prop_Active" id="Prop_Active" value="0"/></td></tr>
<tr><td>SignOn</td><td><input type=text size=50 name="Prop_SignOn" id="Prop_SignOn" value="0"/></td></tr>
<tr><td>Период</td><td><input type=text size=50 name="Prop_Time" id="Prop_Time" value="0"/></td></tr>
<tr><td colspan=3><input type="button" onclick="SaveTag();" value="Сохранить" /></td></tr>
</table></div>

<div id="ChannelGroup" style="display: none; float: left;">
<table>
<tr><th colspan=3>Редактирование:</th></tr>
<tr><td>Тип файла</td><td colspan="2" ><input type=text size=50 name="cg_ftype" id="cg_ftype" maxlength="50" readonly value=""/></td></tr>
<tr><td>Группа каналов</td><td colspan="2" ><input type=text size=50 name="cg_group" id="cg_group" maxlength="255" readonly value=""/></td></tr>
<tr><td>Канал</td><td colspan="2" ><input type=text size=50 name="cg_channel" id="cg_channel" readonly value="0"/></td></tr>
<tr><th colspan="3">    Реакция на сбойные операции</th></tr>
<tr><td>Допустимое значение</td><td colspan="2" ><input type=text size=50 name="cg_warning" id="cg_warning" value="0"/></td></tr>
<tr><td>Критичное значение</td><td colspan="2" ><input type=text size=50 name="cg_error" id="cg_error" value="0"/></td></tr>
<tr><td>Минимальный порог по количеству всех операций</td><td colspan="2" ><input type=text size=50 name="cg_minimal" id="cg_minimal" maxlength="50" value=""/></td></tr>
<tr><th colspan="3">Реакция на общий уровень операций</th></tr>
<tr><td>Допустимый порог по количеству всех операций</td><td  colspan="2" ><input type=text size=50 name="cg_limit" id="cg_limit" value="0"/></td></tr>
<tr><td>Интервал низкой активности</td>
    <td><input type=text size=25 name="cg_limit" id="cg_lowactivity_start" value="0"/></td>
    <td><input type=text size=25 name="cg_limit" id="cg_lowactivity_end" value="0"/></td>
</tr>


<tr><td colspan=3><input type="button" onclick="SaveChannelGroup();" value="Сохранить" /></td></tr>
</table></div>

<input type="hidden" id="buffercolor" value="0" /><div id="r0"></div>
</body>
</html>
<%		
		Conn.Close
		set Cmd=Nothing
		set Conn = Nothing
		set Rs = Nothing
	end if
end if
%>