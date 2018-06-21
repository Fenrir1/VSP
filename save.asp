<%@ language="VBScript"%><!-- #include file="const.asp" --><!-- #include file="common.asp" -->
<%
' Вспомогательный модуль для sets.asp
' сохраняет изменения в БД
set Conn=Server.CreateObject("ADODB.Connection")
Conn.ConnectionTimeout=180
Conn.CommandTimeout=10
Conn.Open(ConnectionString)
set Rs=Server.CreateObject("ADODB.Recordset")
set Cmd=Server.CreateObject("ADODB.Command")
Cmd.ActiveConnection=Conn
Cmd.CommandType=adCmdText

ds=Request("ds")

if ds="SetUserProp" then
	p1=Request("prm1")
	p2=Request("prm2")
	p3=Request("prm3")
	p4=Request("prm4")
	p5=Request("prm5")
	p6=Request("prm6")
	if p1="" then
		SQL_="INSERT INTO [Users] ([User_Name],[User_Login],[Role],[Phone],[Email]) VALUES ('"&p2&"','"&p3&"',"&p4&",'"&p5&"','"&p6&"')"
	else
		SQL_="UPDATE [Users] SET [User_Name]='"&p2&"',[User_Login]='"&p3&"',[Role]="&p4&",[Phone]='"&p5&"',[Email]='"&p6&"' WHERE [User_ID]="&p1
	end if
	Cmd.CommandText=SQL_
	Cmd.Execute
end if

if ds="DelUser" then
	p1=Request("prm1")
	SQL_="DELETE FROM [Users] WHERE [User_ID]='"&p1&"'"
	Cmd.CommandText=SQL_
	Cmd.Execute
end if

if ds="SetFin" then
	p1=Request("prm1")
	p2=Request("prm2")
	p3=Request("prm3")
	if p1="" then
		SQL_="INSERT INTO [V_Branch_code] ([Branch_code],[Name]) VALUES ('"&p2&"','"&p3&"')"
	else
		SQL_="UPDATE [V_Branch_code] SET [Name]='"&p3&"' WHERE [Branch_code]='"&p2&"'"
	end if
	Cmd.CommandText=SQL_
	Cmd.Execute
end if

if ds="DelFin" then
	p1=Request("prm1")
	SQL_="DELETE FROM [V_Branch_code] WHERE [Branch_code]='"&p1&"'"
	Cmd.CommandText=SQL_
	Cmd.Execute
end if

if ds="SetRC" then
	p1=Request("prm1")
	p2=Request("prm2")
	p3=Request("prm3")
	p4=Request("prm4")
	if p1="" then
		SQL_="INSERT INTO [V_Resp_code] ([Resp_code], [Resp_text], [IsFailed]) VALUES ('"&p2&"','"&p3&"',"&p4&")"
	else
		SQL_="UPDATE [V_Resp_code] SET [Resp_text]='"&p3&"', [IsFailed]="&p4&" WHERE [Resp_code]='"&p2&"'"
	end if
	Cmd.CommandText=SQL_
	Cmd.Execute
end if

if ds="DelRC" then
	p1=Request("prm1")
	SQL_="DELETE FROM [V_Resp_code] WHERE [Resp_code]='"&p1&"'"
	Cmd.CommandText=SQL_
	Cmd.Execute
end if

if ds="SetTag" then
	p1=Request("prm1")
	p2=Request("prm2")
	p3=Request("prm3")
	p4=Request("prm4")
	p5=Request("prm5")
	p6=Request("prm6")
	p7=Request("prm7")
	p8=Request("prm8")
	p9=Request("prm9")
	p10=Request("prm10")
	p11=Request("prm11")
	if p1="" then
		SQL_="INSERT INTO [Tags] ([TagID],[TagName],[SetHi],[SetHiHi],[FileID],[GroupName],[Prop_Crit],[Prop_Active],[Prop_SignOn],[Prop_Time]) VALUES ('"&p2&"','"&p3&"',"&p4&","&p5&",'"&p6&"','"&p7&"',"&p8&","&p9&","&p10&","&p11&")"
	else
		SQL_="UPDATE [Tags] SET [TagName]='"&p3&"',[SetHi]="&p4&",[SetHiHi]="&p5&",[FileID]='"&p6&"',[GroupName]='"&p7&"',[Prop_Crit]="&p8&",[Prop_Active]="&p9&",[Prop_SignOn]="&p10&",[Prop_Time]="&p11&" WHERE [TagID]='"&p2&"'"
	end if
	Cmd.CommandText=SQL_
	Cmd.Execute
end if

if ds="DelTag" then
	p1=Request("prm1")
	SQL_="DELETE FROM [Tags] WHERE [TagID]='"&p1&"'"
	Cmd.CommandText=SQL_
	Cmd.Execute
end if

Conn.Close
set Cmd = Nothing
set Rs = Nothing
set Conn = Nothing
%>
