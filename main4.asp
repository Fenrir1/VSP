<%@ language="VBScript"%><!-- #include file="const.asp" --><!-- #include file="common.asp" -->
<%
' основной экран мониторинга прохождения операций по эквайрингу

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

' Змена , на .
function d(v)
	d = replace(v,",",".")
end function

'--------------------------------------------------------------------------------
'------START: get main vars---------------------------------------------
'--------------------------------------------------------------------------------
hoursCount = 6 ' количество часов на графике

dim warnings(20,7), series(10), ATMID(10,4), TOP5TABLES(5)
for i=0 to 10 
  series(i)=""
  ATMID(i,0)="" ' NumEmv
  ATMID(i,1)=""	' DeviceId
  ATMID(i,2)=""	' DeviceType
  ATMID(i,3)=0	' DaysCount
next

' Заполнение Warings
	i = 0
	sqlstr = "SELECT Channel_Group, ISNULL(Warning_Count,0) Warning_Count, ISNULL(Error_Count,0) Error_Count, ISNULL(Min_Count,0) Min_Count, ISNULL(Limit_Count,0) Limit_Count,  ISNULL(Lowactivity_start,0) Lowactivity_start, ISNULL(Lowactivity_end,0) Lowactivity_end  FROM  Channel_Config"
	Rs.OPEN sqlstr, CONN
	if not RS.EOF then
	do while (not RS.EOF)
		warnings(i,0)=Rs.Fields("Channel_Group") ' Channel_Group
		warnings(i,1)=Rs.Fields("Warning_Count")	' Warning_Count
		warnings(i,2)=Rs.Fields("Error_Count")	' Error_Count
		warnings(i,3)=Rs.Fields("Min_Count")	' Min_Count
        warnings(i,4)=Rs.Fields("Limit_Count")	' Limit_Count
        warnings(i,5)=Rs.Fields("Lowactivity_start")	' Lowactivity_start
        warnings(i,6)=Rs.Fields("Lowactivity_end")	' Lowactivity_end
		i = i+1
		Rs.MoveNext
	loop
	end if
	RS.CLOSE
	
	
function checkWarning(paramName, failCount, totalCount, minutes_val)
	res = "" 'clWarning clError
	if (totalCount>0) then
		failCount = (failCount*100)/totalCount ' проверяем процент сбойных
		for j=0 to UBound(warnings)
			if (warnings(j,0)=paramName) then
				if ((warnings(j,2)>0)and(totalCount>warnings(j,3))and(failCount>warnings(j,2))) then
					res = clError
				elseif ((warnings(j,1)>0)and(totalCount>warnings(j,3))and(failCount>warnings(j,1))) then
					res = clWarning
				end if

                if (warnings(j,6)>0) then
                    
					if (warnings(j,5)<=warnings(j,6)) then
						if (((minutes_val<warnings(j,5))or(minutes_val>warnings(j,6)))and(totalCount<warnings(j,4))) then
							res = clError
						end If
					else 
						if (((minutes_val<warnings(j,5))and(minutes_val>warnings(j,6)))and(totalCount<warnings(j,4))) then
							res = clError
						end If
					end if
                end if

				checkWarning = res
			end if
		next
	end if
	checkWarning = res
end function	
	
	
'------------------------------------
'--START: Information Tables----------
'------------------------------------
' Table 1
	ISS_VISA = 0
	ACQ_VISA = 0 
	ISS_MC = 0
	ACQ_MC = 0
	ISS_NSPK_VISA = 0
	ACQ_NSPK_VISA = 0
	ISS_NSPK_MC = 0
	ACQ_NSPK_MC = 0
	ISS_MIR = 0
	ACQ_MIR = 0
	
	ISS_VISA_Color = ""
	ACQ_VISA_Color = "" 
	ISS_MC_Color = ""
	ACQ_MC_Color = ""
	ISS_NSPK_VISA_Color = ""
	ACQ_NSPK_VISA_Color = ""
	ISS_NSPK_MC_Color = ""
	ACQ_NSPK_MC_Color = ""
	ISS_MIR_Color = ""
	ACQ_MIR_Color = ""
	
	sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION,SUM(OPERATION_FAIL) OPERATION_FAIL, SOURCE_CHANNEL "
    sqlstr = sqlstr&" ,DATEPART(HOUR,[TIME])*60+DATEPART(MINUTE,[TIME]) timeinminutes FROM LOG_VO "
	sqlstr = sqlstr&" WHERE [TIME]=(select top 1 [TIME] from LOG_VO order by [TIME] desc) "
	sqlstr = sqlstr&" GROUP BY [TIME],SOURCE_CHANNEL"
	RS.OPEN sqlstr, CONN
	IF NOT RS.EOF THEN
	DO WHILE (NOT RS.EOF)
		if (Rs.Fields("SOURCE_CHANNEL")="VISA") then
			ISS_VISA = Rs.Fields("OPERATION")
			ISS_VISA_Color = checkWarning("VISA_ISS", Rs.Fields("OPERATION_FAIL"), ISS_VISA,  Rs.Fields("timeinminutes"))
		elseif (Rs.Fields("SOURCE_CHANNEL")="MasterCard") then
			ISS_MC = Rs.Fields("OPERATION")
			ISS_MC_Color = checkWarning("MC_ISS", Rs.Fields("OPERATION_FAIL"), ISS_MC,  Rs.Fields("timeinminutes"))
		elseif (Rs.Fields("SOURCE_CHANNEL")="NSPK_VISA") then
			ISS_NSPK_VISA = Rs.Fields("OPERATION")
			ISS_NSPK_VISA_Color = checkWarning("NSPK_VISA_ISS", Rs.Fields("OPERATION_FAIL"), ISS_NSPK_VISA,  Rs.Fields("timeinminutes"))
		elseif (Rs.Fields("SOURCE_CHANNEL")="NSPK_MasterCard") then
			ISS_NSPK_MC = Rs.Fields("OPERATION")
			ISS_NSPK_MC_Color = checkWarning("NSPK_MC_ISS", Rs.Fields("OPERATION_FAIL"), ISS_NSPK_MC,  Rs.Fields("timeinminutes"))
		elseif (Rs.Fields("SOURCE_CHANNEL")="NSPK MIR") then
			ISS_MIR = Rs.Fields("OPERATION")
			ISS_MIR_Color = checkWarning("MIR_ISS", Rs.Fields("OPERATION_FAIL"), ISS_MIR,  Rs.Fields("timeinminutes"))
		end if
		
		Rs.MoveNext
	LOOP
	END IF
	RS.CLOSE
	
	ACQ_NSPK_VISA_Fail = 0
	ACQ_VISA_Fail = 0

    ACQ_NSPK_VISA_time = 0
	ACQ_VISA_time = 0

	sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION,SUM(OPERATION_FAIL) OPERATION_FAIL, TARGET_CHANNEL "
    sqlstr = sqlstr&" ,DATEPART(HOUR,[TIME])*60+DATEPART(MINUTE,[TIME]) timeinminutes FROM LOG_VO "
	sqlstr = sqlstr&" WHERE [TIME]=(select top 1 [TIME] from LOG_VO order by [TIME] desc) "
	sqlstr = sqlstr&" GROUP BY [TIME],TARGET_CHANNEL"
	Rs.OPEN sqlstr, CONN
	If not Rs.EOF then
	do while (not Rs.EOF)
		if (Rs.Fields("TARGET_CHANNEL")="VISA") then
		    ACQ_VISA = ACQ_VISA+Rs.Fields("OPERATION")
			ACQ_VISA_Fail = ACQ_VISA_Fail+Rs.Fields("OPERATION_FAIL")
            ACQ_VISA_time=Rs.Fields("timeinminutes")
		elseif (Rs.Fields("TARGET_CHANNEL")="VISA SMS") then
			ACQ_VISA = ACQ_VISA+Rs.Fields("OPERATION")
			ACQ_VISA_Fail = ACQ_VISA_Fail+Rs.Fields("OPERATION_FAIL")
            ACQ_VISA_time=Rs.Fields("timeinminutes")
		elseif (Rs.Fields("TARGET_CHANNEL")="MasterCard") then
			ACQ_MC = Rs.Fields("OPERATION")
			ACQ_MC_Color = checkWarning("MC_ACQ", Rs.Fields("OPERATION_FAIL"), ACQ_MC,  Rs.Fields("timeinminutes"))
		elseif (Rs.Fields("TARGET_CHANNEL")="NSPK_VISA") then    
			ACQ_NSPK_VISA = ACQ_NSPK_VISA+Rs.Fields("OPERATION")
			ACQ_NSPK_VISA_Fail = ACQ_NSPK_VISA_Fail+Rs.Fields("OPERATION_FAIL")
            ACQ_NSPK_VISA_time=Rs.Fields("timeinminutes")
		elseif (Rs.Fields("TARGET_CHANNEL")="NSPK_VISA SMS") then
		    ACQ_NSPK_VISA = ACQ_NSPK_VISA+Rs.Fields("OPERATION")
			ACQ_NSPK_VISA_Fail = ACQ_NSPK_VISA_Fail+Rs.Fields("OPERATION_FAIL")
            ACQ_NSPK_VISA_time=Rs.Fields("timeinminutes")
		elseif (Rs.Fields("TARGET_CHANNEL")="NSPK_MasterCard") then
			ACQ_NSPK_MC = Rs.Fields("OPERATION")
			ACQ_NSPK_MC_Color = checkWarning("NSPK_MC_ACQ", Rs.Fields("OPERATION_FAIL"), ACQ_NSPK_MC,  Rs.Fields("timeinminutes"))
		elseif (Rs.Fields("TARGET_CHANNEL")="NSPK MIR") then
			ACQ_MIR = Rs.Fields("OPERATION")
			ACQ_MIR_Color = checkWarning("MIR_ACQ", Rs.Fields("OPERATION_FAIL"), ACQ_MIR,  Rs.Fields("timeinminutes"))
		end if
			

		Rs.MoveNext
	LOOP
	END IF
	RS.CLOSE
	
			ACQ_VISA_Color = checkWarning("VISA_ACQ", ACQ_VISA_Fail, ACQ_VISA, ACQ_VISA_time)
		    ACQ_NSPK_VISA_Color = checkWarning("NSPK_VISA_ACQ", ACQ_NSPK_VISA_Fail, ACQ_NSPK_VISA, ACQ_NSPK_VISA_time)
			


' Circle 1	
	circleIndicatorColor1 = clNormal
	
	if ((ISS_VISA_Color = clWarning)or(ACQ_VISA_Color = clWarning)or(ISS_MC_Color = clWarning)or(ACQ_MC_Color = clWarning)or(ISS_NSPK_VISA_Color = clWarning)or(ACQ_NSPK_VISA_Color = clWarning)or(ISS_NSPK_MC_Color = clWarning)or(ACQ_NSPK_MC_Color = clWarning)or(ISS_MIR_Color = clWarning)or(ACQ_MIR_Color = clWarning)) then
		circleIndicatorColor1 = clWarning
	end if
	
	if ((ISS_VISA_Color = clError)or(ACQ_VISA_Color = clError)or(ISS_MC_Color = clError)or(ACQ_MC_Color = clError)or(ISS_NSPK_VISA_Color = clError)or(ACQ_NSPK_VISA_Color = clError)or(ISS_NSPK_MC_Color = clError)or(ACQ_NSPK_MC_Color = clError)or(ISS_MIR_Color = clError)or(ACQ_MIR_Color = clError)) then
		circleIndicatorColor1 = clError
	end if
	
	
' Table 2
 All_ATM = 0
 All_BPT = 0
 All_POS = 0
 All_H2H_RBS = 0
 
 All_ATM_Color = ""
 All_BPT_Color = ""
 All_POS_Color = ""
 All_H2H_RBS_Color = ""
 
 	sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION,SUM(OPERATION_FAIL) OPERATION_FAIL, SOURCE_CHANNEL "
    sqlstr = sqlstr&" ,DATEPART(HOUR,[TIME])*60+DATEPART(MINUTE,[TIME]) timeinminutes FROM LOG_VO "
	sqlstr = sqlstr&" WHERE [TIME]=(select top 1 [TIME] from LOG_VO order by [TIME] desc) "
	sqlstr = sqlstr&" GROUP BY [TIME],SOURCE_CHANNEL"
	Rs.OPEN sqlstr, CONN
	If not Rs.EOF then
	do while (not Rs.EOF)
		if (Rs.Fields("SOURCE_CHANNEL")="OUR_ATM") then
			All_ATM = Rs.Fields("OPERATION")
			All_ATM_Color = checkWarning("ATM_ACQ", Rs.Fields("OPERATION_FAIL"), All_ATM,  Rs.Fields("timeinminutes"))
		elseif (Rs.Fields("SOURCE_CHANNEL")="OUR_POS") then
			All_POS = Rs.Fields("OPERATION")
			All_POS_Color = checkWarning("POS_ACQ", Rs.Fields("OPERATION_FAIL"), All_POS,  Rs.Fields("timeinminutes"))
		elseif (Rs.Fields("SOURCE_CHANNEL")="OUR_BPT") then
			All_BPT = Rs.Fields("OPERATION")
			All_BPT_Color = checkWarning("BPT_ACQ", Rs.Fields("OPERATION_FAIL"), All_BPT,  Rs.Fields("timeinminutes"))
		elseif (Rs.Fields("SOURCE_CHANNEL")="H2H_BPCRBS") then
			All_H2H_RBS = Rs.Fields("OPERATION")
			All_H2H_RBS_Color = checkWarning("H2H_RBS", Rs.Fields("OPERATION_FAIL"), All_H2H_RBS,  Rs.Fields("timeinminutes"))
		end if

		Rs.MoveNext
	LOOP
	END IF
	RS.CLOSE
	
' Circle 2	
	circleIndicatorColor2 = clNormal
	
	if ((All_ATM_Color = clWarning)or(All_POS_Color = clWarning)or(All_BPT_Color = clWarning)or(All_H2H_RBS_Color = clWarning)) then
		circleIndicatorColor2 = clWarning
	end if
	
	if ((All_ATM_Color = clError)or(All_POS_Color = clError)or(All_BPT_Color = clError)or(All_H2H_RBS_Color = clError)) then
		circleIndicatorColor2 = clError
	end if	
  
' Table 3
 All_3DS = 0
 All_SOA = 0
 
 All_3DS_Color = ""
 All_SOA_Color = ""
 
 VISA_3DS = ""
 NSPK_VISA_3DS = ""
 MC_3DS = ""
 NSPK_MC_3DS = ""
 NSPK_MIR_3DS = ""

 SOA_AGENT = ""
 SOA_USB = ""
 
 	sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION,SUM(OPERATION_FAIL) OPERATION_FAIL, SOURCE_CHANNEL "
    sqlstr = sqlstr&" ,DATEPART(HOUR,[TIME])*60+DATEPART(MINUTE,[TIME]) timeinminutes FROM LOG_VS WHERE "
	sqlstr = sqlstr&" ((SERVICE='3D-Secure' and SOURCE_CHANNEL in ('NSPK_VISA','NSPK_MasterCard','VISA','MasterCard','NSPK_MIR')) or SERVICE='SOA_AGENT' or SERVICE='SOA_USB') and "
	sqlstr = sqlstr&" [TIME]>=convert(datetime,floor(convert(float,Getdate()))) GROUP BY [TIME], SOURCE_CHANNEL order by [TIME]"
	Rs.Open sqlstr, Conn
	If not Rs.EOF then
	do while (not Rs.EOF)
		if (Rs.Fields("SOURCE_CHANNEL")="NSPK_VISA") then
			NSPK_VISA_3DS = checkWarning("NSPK_VISA_3DS", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"),  Rs.Fields("timeinminutes"))
		elseif (Rs.Fields("SOURCE_CHANNEL")="NSPK_MasterCard") then
			NSPK_MC_3DS = checkWarning("NSPK_MC_3DS", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"),  Rs.Fields("timeinminutes"))
		elseif (Rs.Fields("SOURCE_CHANNEL")="VISA") then
			VISA_3DS = checkWarning("VISA_3DS", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"),  Rs.Fields("timeinminutes"))
		elseif (Rs.Fields("SOURCE_CHANNEL")="MasterCard") then
			MC_3DS = checkWarning("MC_3DS", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"),  Rs.Fields("timeinminutes"))
		elseif (Rs.Fields("SOURCE_CHANNEL")="NSPK_MIR") then
			NSPK_MIR_3DS = checkWarning("NSPK_MIR_3DS", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"),  Rs.Fields("timeinminutes"))
		elseif (Rs.Fields("SOURCE_CHANNEL")="RBS") then
			SOA_USB = checkWarning("SOA_USB", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"),  Rs.Fields("timeinminutes"))
		elseif (Rs.Fields("SOURCE_CHANNEL")="OUR_POS") then
			SOA_AGENT = checkWarning("SOA_AGENT", Rs.Fields("OPERATION_FAIL"), Rs.Fields("OPERATION"),  Rs.Fields("timeinminutes"))
		end if
		
		Rs.MoveNext
	LOOP
	END IF
	RS.CLOSE

 
 	sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION, SERVICE "
    sqlstr = sqlstr&" ,DATEPART(HOUR,[TIME])*60+DATEPART(MINUTE,[TIME]) timeinminutes FROM LOG_VS "
	sqlstr = sqlstr&" WHERE [TIME]=(select top 1 [TIME] from LOG_VS order by [TIME] desc) "
	sqlstr = sqlstr&" and ((SERVICE='3D-Secure' and SOURCE_CHANNEL in ('NSPK_VISA','NSPK_MasterCard','VISA','MasterCard','NSPK_MIR')) or SERVICE='SOA_AGENT' or SERVICE='SOA_USB') "
	sqlstr = sqlstr&" GROUP BY [TIME],SERVICE"
	Rs.OPEN sqlstr, CONN
	If not Rs.EOF then
	do while (not Rs.EOF)
		if (Rs.Fields("SERVICE")="3D-Secure") then
			All_3DS = Rs.Fields("OPERATION")
		elseif (Rs.Fields("SERVICE")="SOA_AGENT") then
			All_SOA = All_SOA+Rs.Fields("OPERATION")
		elseif (Rs.Fields("SERVICE")="SOA_USB") then
			All_SOA = All_SOA+Rs.Fields("OPERATION")
		end if

		Rs.MoveNext
	LOOP
	END IF
	RS.CLOSE	
	
	
	if ((VISA_3DS = clWarning)or(NSPK_VISA_3DS = clWarning)or(MC_3DS = clWarning)or(NSPK_MC_3DS = clWarning)or(NSPK_MIR_3DS = clWarning)) then
		All_3DS_Color = clWarning
	end if
	
	if ((VISA_3DS = clError)or(NSPK_VISA_3DS = clError)or(MC_3DS = clError)or(NSPK_MC_3DS = clError)or(NSPK_MIR_3DS = clError)) then
		All_3DS_Color = clError
	end if 
	
	
	if ((SOA_USB = clWarning)or(SOA_AGENT = clWarning)) then
		All_SOA_Color = clWarning
	end if
	
	if ((SOA_USB = clError)or(SOA_AGENT = clError)) then
		All_SOA_Color = clError
	end if 

' Circle 3	
	circleIndicatorColor3 = clNormal
	
	if ((All_3DS_Color = clWarning)or(All_SOA_Color = clWarning)) then
		circleIndicatorColor3 = clWarning
	end if
	
	if ((All_3DS_Color = clError)or(All_SOA_Color = clError)) then
		circleIndicatorColor3 = clError
	end if	


'------------------------------------
'--END: Information Tables------------
'------------------------------------

CurrentTime = DateTimeFormat(DateAdd("m", -1, Now), "yyyy, mm, dd, hh, nn")
'------------------------------------
'--START: Series for Charts----------
'------------------------------------
' Dots
	ISS_VISA_Dot = ""
	ISS_MC_Dot = ""
	ISS_NSPK_VISA_Dot = ""
	ISS_NSPK_MC_Dot = ""
	ISS_MIR_Dot = ""
	
	ACQ_VISA_Dot = ""
	ACQ_NSPK_VISA_Dot = ""
	ACQ_MC_Dot = ""
	ACQ_NSPK_MC_Dot = ""
	ACQ_MIR_Dot = ""
	
	All_ATM_Dot = ""
	All_BPT_Dot = ""
	All_POS_Dot = ""
	All_H2H_RBS_Dot = ""
	
	All_3DS_Dot = ""
	All_SOA_Dot = ""
	

	sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],Channel_Count , Channel_Group FROM Channel_Fail_Series  WHERE [TIME]>=dateadd(hour,-"&hoursCount&",Getdate()) order by [TIME]"
	Rs.Open sqlstr, Conn
	If not Rs.EOF then
	do while (not Rs.EOF)	
		v = Rs.Fields("Channel_Count")
		v1 = Rs.Fields("TIME")
		if (Rs.Fields("Channel_Group")="VISA_ISS") then
			'if (ISS_VISA_Dot<>"") then 
				ISS_VISA_Dot = ISS_VISA_Dot&", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	ISS_VISA_Dot = ISS_VISA_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		elseif (Rs.Fields("Channel_Group")="MC_ISS") then
			'if (ISS_MC_Dot<>"") then 
				ISS_MC_Dot = ISS_MC_Dot&", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	ISS_MC_Dot = ISS_MC_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		elseif (Rs.Fields("Channel_Group")="NSPK_VISA_ISS") then
			'if (ISS_NSPK_VISA_Dot<>"") then 
				ISS_NSPK_VISA_Dot = ISS_NSPK_VISA_Dot&", { name: '', type: 'scatter', yAxis: 0, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	ISS_NSPK_VISA_Dot = ISS_NSPK_VISA_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		elseif (Rs.Fields("Channel_Group")="NSPK_MC_ISS") then
			'if (ISS_NSPK_MC_Dot<>"") then 
				ISS_NSPK_MC_Dot = ISS_NSPK_MC_Dot&", { name: '', type: 'scatter', yAxis: 0, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	ISS_NSPK_MC_Dot = ISS_NSPK_MC_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		elseif (Rs.Fields("Channel_Group")="MIR_ISS") then
			'if (ISS_MIR_Dot<>"") then 
				ISS_MIR_Dot = ISS_MIR_Dot&", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	ISS_MIR_Dot = ISS_MIR_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		elseif (Rs.Fields("Channel_Group")="VISA_ACQ") then
			'if (ACQ_VISA_Dot<>"") then 
				ACQ_VISA_Dot = ACQ_VISA_Dot&", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	ACQ_VISA_Dot = ACQ_VISA_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		elseif (Rs.Fields("Channel_Group")="NSPK_VISA_ACQ") then
			'if (ACQ_NSPK_VISA_Dot<>"") then 
				ACQ_NSPK_VISA_Dot = ACQ_NSPK_VISA_Dot&", { name: '', type: 'scatter', yAxis: 0, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	ACQ_NSPK_VISA_Dot = ACQ_NSPK_VISA_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		elseif (Rs.Fields("Channel_Group")="MC_ACQ") then
			'if (ACQ_MC_Dot<>"") then 
				ACQ_MC_Dot = ACQ_MC_Dot&", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	ACQ_MC_Dot = ACQ_MC_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		elseif (Rs.Fields("Channel_Group")="NSPK_MC_ACQ") then
			'if (ACQ_NSPK_MC_Dot<>"") then 
				ACQ_NSPK_MC_Dot = ACQ_NSPK_MC_Dot&", { name: '', type: 'scatter', yAxis: 0, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	ACQ_NSPK_MC_Dot = ACQ_NSPK_MC_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		elseif (Rs.Fields("Channel_Group")="MIR_ACQ") then
			'if (ACQ_MIR_Dot<>"") then 
				ACQ_MIR_Dot = ACQ_MIR_Dot&", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	ACQ_MIR_Dot = ACQ_MIR_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		elseif (Rs.Fields("Channel_Group")="ATM_ACQ") then
			'if (All_ATM_Dot<>"") then 
				All_ATM_Dot = All_ATM_Dot&", { name: '', type: 'scatter', yAxis: 0, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	All_ATM_Dot = All_ATM_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		elseif (Rs.Fields("Channel_Group")="BPT_ACQ") then
			'if (All_BPT_Dot<>"") then 
				All_BPT_Dot = All_BPT_Dot&", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	All_BPT_Dot = All_BPT_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		elseif (Rs.Fields("Channel_Group")="POS_ACQ") then
			'if (All_POS_Dot<>"") then 
				All_POS_Dot = All_POS_Dot&", { name: '', type: 'scatter', yAxis: 0, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	All_POS_Dot = All_POS_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		elseif (Rs.Fields("Channel_Group")="H2H_RBS") then
			'if (All_H2H_RBS_Dot<>"") then 
				All_H2H_RBS_Dot = All_H2H_RBS_Dot&", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	All_H2H_RBS_Dot = All_H2H_RBS_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		elseif (Rs.Fields("Channel_Group")="3DS") then
			'if (All_3DS_Dot<>"") then 
				All_3DS_Dot = All_3DS_Dot&", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	All_3DS_Dot = All_3DS_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		elseif (Rs.Fields("Channel_Group")="SOA") then
			'if (All_SOA_Dot<>"") then 
				All_SOA_Dot = All_SOA_Dot&", { name: '', type: 'scatter', yAxis: 0, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'else
			'	All_SOA_Dot = All_SOA_Dot&" { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(v1, "yyyy, mm, dd, hh, nn")&"), y: "&v&"}]}"
			'end if 
		end if
	
		Rs.MoveNext
	loop
	end if
	Rs.Close

' Chart 1
	VISA_Series_Color = "#66ff66"
	MC_Series_Color = "#ffff66"
	NSPK_VISA_Series_Color = "#66ffff"
	NSPK_MC_Series_Color = "#FF66FF"
	MIR_Series_Color = "#3366FF"
	
    ISS_VISA_Series = ""
	ISS_MC_Series = ""
	ISS_NSPK_VISA_Series = ""
	ISS_NSPK_MC_Series = ""
	ISS_MIR_Series = ""
	
	LastDate = DateAdd("m", -1, Now)
	LastDate_ISS_VISA_Series = DateAdd("m", -1, Now)
	LastDate_ISS_MC_Series = DateAdd("m", -1, Now)
	LastDate_ISS_NSPK_VISA_Series = DateAdd("m", -1, Now)
	LastDate_ISS_NSPK_MC_Series = DateAdd("m", -1, Now)
	LastDate_ISS_MIR_Series = DateAdd("m", -1, Now)
	
	
	sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION, SOURCE_CHANNEL FROM LOG_VO WHERE [TIME]>=dateadd(hour,-"&hoursCount&",Getdate()) GROUP BY [TIME], SOURCE_CHANNEL order by [TIME]"
	Rs.Open sqlstr, Conn
	If not Rs.EOF then
	do while (not Rs.EOF)
	
	    v = Rs.Fields("OPERATION")
		if (Rs.Fields("SOURCE_CHANNEL")="VISA") then
			if (ISS_VISA_Series<>"") then 
				ISS_VISA_Series = ISS_VISA_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				ISS_VISA_Series = ISS_VISA_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if 
			LastDate_ISS_VISA_Series = Rs.Fields("TIME")
		elseif (Rs.Fields("SOURCE_CHANNEL")="MasterCard") then
			if (ISS_MC_Series<>"") then 
				ISS_MC_Series = ISS_MC_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				ISS_MC_Series = ISS_MC_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if 
			LastDate_ISS_MC_Series = Rs.Fields("TIME")
		elseif (Rs.Fields("SOURCE_CHANNEL")="NSPK_VISA") then
			if (ISS_NSPK_VISA_Series<>"") then 
				ISS_NSPK_VISA_Series = ISS_NSPK_VISA_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				ISS_NSPK_VISA_Series = ISS_NSPK_VISA_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if 
			LastDate_ISS_NSPK_VISA_Series = Rs.Fields("TIME")
		elseif (Rs.Fields("SOURCE_CHANNEL")="NSPK_MasterCard") then
			if (ISS_NSPK_MC_Series<>"") then 
				ISS_NSPK_MC_Series = ISS_NSPK_MC_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				ISS_NSPK_MC_Series = ISS_NSPK_MC_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if
			LastDate_ISS_NSPK_MC_Series = Rs.Fields("TIME")
		elseif (Rs.Fields("SOURCE_CHANNEL")="NSPK MIR") then
			if (ISS_MIR_Series<>"") then 
				ISS_MIR_Series = ISS_MIR_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				ISS_MIR_Series = ISS_MIR_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if
			LastDate_ISS_MIR_Series = Rs.Fields("TIME")
		end if		

		Rs.MoveNext
	loop
	end if
	Rs.Close
	
	ISS_VISA_Series = "{ name: 'ISS_VISA', color: '"&VISA_Series_Color&"', type: 'line', yAxis: 1, data: ["&ISS_VISA_Series&"]}"
	'ISS_VISA_Dot = ""
	'if (ISS_VISA_Dot = clError) then
	'	ISS_VISA_Dot = ", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_ISS_VISA_Series, "yyyy, mm, dd, hh, nn")&"), y: "&ISS_VISA&"}]}"
	'end if
	ISS_MC_Series = ",{ name: 'ISS_MC', color: '"&MC_Series_Color&"', type: 'line', yAxis: 1, data: ["&ISS_MC_Series&"]}"
	'ISS_MC_Dot = ""
	'if (ISS_MC_Color = clError) then
	'	ISS_MC_Dot = ", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_ISS_MC_Series, "yyyy, mm, dd, hh, nn")&"), y: "&ISS_MC&"}]}"
	'end if
	ISS_NSPK_VISA_Series = ",{ name: 'ISS_NSPK_VISA', color: '"&NSPK_VISA_Series_Color&"', type: 'line', yAxis: 0, data: ["&ISS_NSPK_VISA_Series&"]}"
	'ISS_NSPK_VISA_Dot = ""
	'if (ISS_NSPK_VISA_Color = clError) then
	'	ISS_NSPK_VISA_Dot = ", { name: '', type: 'scatter', yAxis: 0, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_ISS_NSPK_VISA_Series, "yyyy, mm, dd, hh, nn")&"), y: "&ISS_NSPK_VISA&"}]}"
	'end if
	ISS_NSPK_MC_Series = ",{ name: 'ISS_NSPK_MC', color: '"&NSPK_MC_Series_Color&"', type: 'line', yAxis: 0, data: ["&ISS_NSPK_MC_Series&"]}"
	'ISS_NSPK_MC_Dot = ""
	'if (ISS_NSPK_MC_Color = clError) then
	'	ISS_NSPK_MC_Dot = ", { name: '', type: 'scatter', yAxis: 0, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_ISS_NSPK_MC_Series, "yyyy, mm, dd, hh, nn")&"), y: "&ISS_NSPK_MC&"}]}"
	'end if
	ISS_MIR_Series = ",{ name: 'ISS_MIR', color: '"&MIR_Series_Color&"', type: 'line', yAxis: 1, data: ["&ISS_MIR_Series&"]}"
	'ISS_MIR_Dot = ""
	'if (ISS_MIR_Color = clError) then
	'	ISS_MIR_Dot = ", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_ISS_MIR_Series, "yyyy, mm, dd, hh, nn")&"), y: "&ISS_MIR&"}]}"
	'end if

' Chart 2	
	ACQ_VISA_Series = ""
	ACQ_MC_Series = ""
	ACQ_NSPK_VISA_Series = ""
	ACQ_NSPK_MC_Series = ""
	ACQ_MIR_Series = ""
	
	LastDate_ACQ_VISA_Series = DateAdd("m", -1, Now)
	LastDate_ACQ_MC_Series = DateAdd("m", -1, Now)
	LastDate_ACQ_NSPK_VISA_Series = DateAdd("m", -1, Now)
	LastDate_ACQ_NSPK_MC_Series = DateAdd("m", -1, Now)
	LastDate_ACQ_MIR_Series = DateAdd("m", -1, Now)
	
	sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION, case when TARGET_CHANNEL='NSPK_VISA SMS' then 'NSPK_VISA' when TARGET_CHANNEL='VISA SMS' then 'VISA' else TARGET_CHANNEL end as TARGET_CHANNEL FROM LOG_VO "
	sqlstr = sqlstr&"WHERE [TIME]>=dateadd(hour,-"&hoursCount&",Getdate()) GROUP BY [TIME], case when TARGET_CHANNEL='NSPK_VISA SMS' then 'NSPK_VISA' when TARGET_CHANNEL='VISA SMS' then 'VISA' else TARGET_CHANNEL end order by [TIME]"
	Rs.Open sqlstr, Conn
	If not Rs.EOF then
	do while (not Rs.EOF)

	    v = Rs.Fields("OPERATION")
		if (Rs.Fields("TARGET_CHANNEL")="VISA")then
			if (ACQ_VISA_Series<>"") then 
				ACQ_VISA_Series = ACQ_VISA_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				ACQ_VISA_Series = ACQ_VISA_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if 
			LastDate_ACQ_VISA_Series = Rs.Fields("TIME")
		elseif (Rs.Fields("TARGET_CHANNEL")="MasterCard") then
			if (ACQ_MC_Series<>"") then 
				ACQ_MC_Series = ACQ_MC_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				ACQ_MC_Series = ACQ_MC_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if 
			LastDate_ACQ_MC_Series = Rs.Fields("TIME")
		elseif (Rs.Fields("TARGET_CHANNEL")="NSPK_VISA") then
			if (ACQ_NSPK_VISA_Series<>"") then 
				ACQ_NSPK_VISA_Series = ACQ_NSPK_VISA_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				ACQ_NSPK_VISA_Series = ACQ_NSPK_VISA_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if 
			LastDate_ACQ_NSPK_VISA_Series = Rs.Fields("TIME")
		elseif (Rs.Fields("TARGET_CHANNEL")="NSPK_MasterCard") then
			if (ACQ_NSPK_MC_Series<>"") then 
				ACQ_NSPK_MC_Series = ACQ_NSPK_MC_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				ACQ_NSPK_MC_Series = ACQ_NSPK_MC_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if
			LastDate_ACQ_NSPK_MC_Series = Rs.Fields("TIME")
		elseif (Rs.Fields("TARGET_CHANNEL")="NSPK MIR") then
			if (ACQ_MIR_Series<>"") then 
				ACQ_MIR_Series = ACQ_MIR_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				ACQ_MIR_Series = ACQ_MIR_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if
			LastDate_ACQ_MIR_Series = Rs.Fields("TIME")
		end if	
	
		
		Rs.MoveNext
	loop
	end if
	Rs.Close	

			
	ACQ_VISA_Series = "{ name: 'ACQ_VISA', color: '"&VISA_Series_Color&"', type: 'line', yAxis: 1, data: ["&ACQ_VISA_Series&"]}"
	'ACQ_VISA_Dot = ""
	'if (ACQ_VISA_Color = clError) then
	'	ACQ_VISA_Dot = ", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_ACQ_VISA_Series, "yyyy, mm, dd, hh, nn")&"), y: "&ACQ_VISA&"}]}"
	'end if
	ACQ_MC_Series = ",{ name: 'ACQ_MC', color: '"&MC_Series_Color&"', type: 'line', yAxis: 1, data: ["&ACQ_MC_Series&"]}"
	'ACQ_MC_Dot = ""
	'if (ACQ_MC_Color = clError) then
	'	ACQ_MC_Dot = ", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_ACQ_MC_Series, "yyyy, mm, dd, hh, nn")&"), y: "&ACQ_MC&"}]}"
	'end if
	ACQ_NSPK_VISA_Series = ",{ name: 'IACQ_NSPK_VISA', color: '"&NSPK_VISA_Series_Color&"', type: 'line', yAxis: 0, data: ["&ACQ_NSPK_VISA_Series&"]}"
	'ACQ_NSPK_VISA_Dot = ""
	'if (ACQ_NSPK_VISA_Color = clError) then
	'	ACQ_NSPK_VISA_Dot = ", { name: '', type: 'scatter', yAxis: 0, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_ACQ_NSPK_VISA_Series, "yyyy, mm, dd, hh, nn")&"), y: "&ACQ_NSPK_VISA&"}]}"
	'end if
	ACQ_NSPK_MC_Series = ",{ name: 'ACQ_NSPK_MC', color: '"&NSPK_MC_Series_Color&"', type: 'line', yAxis: 0, data: ["&ACQ_NSPK_MC_Series&"]}"
	'ACQ_NSPK_MC_Dot = ""
	'if (ACQ_NSPK_MC_Color = clError) then
	'	ACQ_NSPK_MC_Dot = ", { name: '', type: 'scatter', yAxis: 0, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_ACQ_NSPK_MC_Series, "yyyy, mm, dd, hh, nn")&"), y: "&ACQ_NSPK_MC&"}]}"
	'end if
	ACQ_MIR_Series = ",{ name: 'ACQ_MIR', color: '"&MIR_Series_Color&"', type: 'line', yAxis: 1, data: ["&ACQ_MIR_Series&"]}"
	'ACQ_MIR_Dot = ""
	'if (ACQ_MIR_Color = clError) then
	'	ACQ_MIR_Dot = ", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_ACQ_MIR_Series, "yyyy, mm, dd, hh, nn")&"), y: "&ACQ_MIR&"}]}"
	'end if

' Chart 3
	All_ATM_Series_Color = "#66ff66"
	All_BPT_Series_Color = "#ffff66"
	All_POS_Series_Color = "#66ffff"
	All_H2H_RBS_Series_Color = "#FF66FF"		
	All_ATM_Series = ""
	All_BPT_Series = ""
	All_POS_Series = ""
	All_H2H_RBS_Series = ""	
	
	LastDate = DateAdd("m", -1, Now)
	LastDate_All_ATM_Series = DateAdd("m", -1, Now)
	LastDate_All_BPT_Series = DateAdd("m", -1, Now)
	LastDate_All_POS_Series = DateAdd("m", -1, Now)
	LastDate_All_H2H_RBS_Series = DateAdd("m", -1, Now)
	
	sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION, SOURCE_CHANNEL FROM Log_VO WHERE [TIME]>=dateadd(hour,-"&hoursCount&",Getdate()) GROUP BY [TIME], SOURCE_CHANNEL order by [TIME]"
	Rs.Open sqlstr, Conn
	If not Rs.EOF then
	do while (not Rs.EOF)
	
	    v = Rs.Fields("OPERATION")
		if (Rs.Fields("SOURCE_CHANNEL")="OUR_ATM") then
			if (All_ATM_Series<>"") then 
				All_ATM_Series = All_ATM_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				All_ATM_Series = All_ATM_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if 
			LastDate_All_ATM_Series = Rs.Fields("TIME")
		elseif (Rs.Fields("SOURCE_CHANNEL")="OUR_BPT") then
			if (All_BPT_Series<>"") then 
				All_BPT_Series = All_BPT_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				All_BPT_Series = All_BPT_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if 
			LastDate_All_BPT_Series = Rs.Fields("TIME")
		elseif (Rs.Fields("SOURCE_CHANNEL")="OUR_POS") then
			if (All_POS_Series<>"") then 
				All_POS_Series = All_POS_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				All_POS_Series = All_POS_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if 
			LastDate_All_POS_Series = Rs.Fields("TIME")
		elseif (Rs.Fields("SOURCE_CHANNEL")="H2H_BPCRBS") then
			if (All_H2H_RBS_Series<>"") then 
				All_H2H_RBS_Series = All_H2H_RBS_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				All_H2H_RBS_Series = All_H2H_RBS_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if
			LastDate_All_H2H_RBS_Series = Rs.Fields("TIME")
		end if		
		
		Rs.MoveNext
	loop
	end if
	Rs.Close
	
	All_ATM_Series = "{ name: 'ATM', color: '"&All_ATM_Series_Color&"', type: 'line', yAxis: 0, data: ["&All_ATM_Series&"]}"
	'All_ATM_Dot = ""
	'if (All_ATM_Color = clError) then
	'	All_ATM_Dot = ", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_All_ATM_Series, "yyyy, mm, dd, hh, nn")&"), y: "&All_ATM&"}]}"
	'end if	
	All_BPT_Series = ",{ name: 'BPT', color: '"&All_BPT_Series_Color&"', type: 'line', yAxis: 1, data: ["&All_BPT_Series&"]}"
	'All_BPT_Dot = ""
	'if (All_BPT_Color = clError) then
	'	All_BPT_Dot = ", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_All_BPT_Series, "yyyy, mm, dd, hh, nn")&"), y: "&All_BPT&"}]}"
	'end if	
	All_POS_Series = ",{ name: 'POS', color: '"&All_POS_Series_Color&"', type: 'line', yAxis: 0, data: ["&All_POS_Series&"]}"
	'All_POS_Dot = ""
	'if (All_POS_Color = clError) then
	'	All_POS_Dot = ", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_All_POS_Series, "yyyy, mm, dd, hh, nn")&"), y: "&All_POS&"}]}"
	'end if
	All_H2H_RBS_Series = ",{ name: 'H2H_RBS', color: '"&All_H2H_RBS_Series_Color&"', type: 'line', yAxis: 1, data: ["&All_H2H_RBS_Series&"]}"
	'All_H2H_RBS_Dot = ""
	'if (All_H2H_RBS_Color = clError) then
	'	All_H2H_RBS_Dot = ", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_All_H2H_RBS_Series, "yyyy, mm, dd, hh, nn")&"), y: "&All_H2H_RBS&"}]}"
	'end if

'-----------------------------------------------------------------------------------------
	All_3DS_Series_Color = "#66ff66"
	All_SOA_Series_Color = "#ffff66"
	All_3DS_Series = ""
	All_SOA_Series = ""
	LastDate = DateAdd("m", -1, Now)
	LastDate_All_3DS_Series = DateAdd("m", -1, Now)
	LastDate_All_SOA_Series = DateAdd("m", -1, Now)
	
	'sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION, SERVICE FROM Log_VS WHERE [TIME]>=dateadd(hour,-"&hoursCount&",Getdate()) GROUP BY [TIME], SERVICE order by [TIME]"
	sqlstr = "SELECT DATEADD(MONTH,-1,[TIME]) [TIME],SUM(OPERATION) OPERATION, "
	sqlstr = sqlstr&"case when [SERVICE] in ('SOA_AGENT','SOA_USB') then 'SOA_AGENT'"
	sqlstr = sqlstr&"when [SERVICE] = '3D-Secure' then '3D-Secure' end as [SERVICE] "
	sqlstr = sqlstr&"FROM Log_VS "
	sqlstr = sqlstr&"WHERE "
	sqlstr = sqlstr&"((SERVICE='3D-Secure' and SOURCE_CHANNEL in ('NSPK_VISA','NSPK_MasterCard','VISA','MasterCard')) or SERVICE='SOA_AGENT' or SERVICE='SOA_USB') and "
	sqlstr = sqlstr&"[TIME]>=dateadd(hour,-"&hoursCount&",Getdate()) "
	sqlstr = sqlstr&"GROUP BY [TIME], "
	sqlstr = sqlstr&"(case when [SERVICE] in ('SOA_AGENT','SOA_USB') then 'SOA_AGENT'"
	sqlstr = sqlstr&"when [SERVICE] = '3D-Secure' then '3D-Secure' end) "
	sqlstr = sqlstr&"order by [TIME]"
	Rs.Open sqlstr, Conn
	If not Rs.EOF then
	do while (not Rs.EOF)
	
	    v = Rs.Fields("OPERATION")
		if (Rs.Fields("SERVICE")="3D-Secure") then
			if (All_3DS_Series<>"") then 
				All_3DS_Series = All_3DS_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				All_3DS_Series = All_3DS_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if 
			LastDate_All_3DS_Series = Rs.Fields("TIME")
		elseif (Rs.Fields("SERVICE")="SOA_AGENT") then
			if (All_SOA_Series<>"") then 
				All_SOA_Series = All_SOA_Series&",{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			else
				All_SOA_Series = All_SOA_Series&"{x: Date.UTC("&DateTimeFormat(Rs.Fields("TIME"), "yyyy, mm, dd, hh, nn")&"), y: "&v&"}"
			end if 
			LastDate_All_SOA_Series = Rs.Fields("TIME")
		end if		
		
		Rs.MoveNext
	loop
	end if
	Rs.Close
	
	All_3DS_Series = "{ name: '3DS', color: '"&All_3DS_Series_Color&"', type: 'line', yAxis: 1, data: ["&All_3DS_Series&"]}"
	'All_3DS_Dot = ""
	'if (All_3DS_Color = clError) then
	'	All_3DS_Dot = ", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_All_3DS_Series, "yyyy, mm, dd, hh, nn")&"), y: "&All_3DS&"}]}"
	'end if
	All_SOA_Series = ",{ name: 'SOA', color: '"&All_SOA_Series_Color&"', type: 'line', yAxis: 0, data: ["&All_SOA_Series&"]}"
	'All_SOA_Dot = ""
	'if (All_SOA_Color = clError) then
	'	All_SOA_Dot = ", { name: '', type: 'scatter', yAxis: 1, data: [{color: null, marker: {fillColor: '#FF0000', lineColor: '#FF0000', radius: 5}, x: Date.UTC("&DateTimeFormat(LastDate_All_SOA_Series, "yyyy, mm, dd, hh, nn")&"), y: "&All_SOA&"}]}"
	'end if

'------------------------------------
'--END: Series for Chart2------------
'------------------------------------

'------------------------------------
'--START: Data CircleIndicator------
'------------------------------------
TimeVO = 0
TimeVS = 0

sqlstr = "select FileType,LastUpdate, DATEDIFF(MINUTE,LastUpdate,GETDATE()) IntervalLength from  VIP_files"
Rs.Open sqlstr, Conn
If not Rs.EOF then
do while (not Rs.EOF)
	if (Rs.Fields("FileType")="VO") then
		TimeVO = Rs.Fields("IntervalLength")
		LastUpdateVO = DateTimeFormat(Rs.Fields("LastUpdate"), "dd.mm.yyyy hh:nn")
	elseif (Rs.Fields("FileType")="VS") then
		TimeVS = Rs.Fields("IntervalLength")
		LastUpdateVS = DateTimeFormat(Rs.Fields("LastUpdate"), "dd.mm.yyyy hh:nn")
	end if
	Rs.MoveNext
loop
end if
Rs.Close

	PeriodVO = 10
	PeriodVS = 10

circleIndicatorMarker1 = ""
circleIndicatorMarker2 = ""
circleIndicatorMarker3 = ""

if ((TimeVO>PeriodVO) and (PeriodVO>0)) then
	'circleIndicatorMarker1 = "circleIndicator1.renderer.image('q.gif', 50, 50, 100, 100).add();"
	'circleIndicatorMarker2 = "circleIndicator2.renderer.image('q.gif', 25, 25, 50, 50).add();"

     circleIndicatorMarker1 = " var image1 = document.createElementNS('http://www.w3.org/2000/svg', 'image');"
     circleIndicatorMarker1 = circleIndicatorMarker1&"   image1.setAttributeNS(null, 'x', 50);  image1.setAttributeNS(null, 'y', 50); image1.setAttributeNS(null, 'width', 100); "
     circleIndicatorMarker1 = circleIndicatorMarker1&"   image1.setAttributeNS(null, 'height', 100); image1.setAttributeNS(null, 'href', 'q.gif'); svg1.appendChild(image1); "

     circleIndicatorMarker2 = " var image2 = document.createElementNS('http://www.w3.org/2000/svg', 'image');"
     circleIndicatorMarker2 = circleIndicatorMarker2&"   image2.setAttributeNS(null, 'x', 25);  image2.setAttributeNS(null, 'y', 25); image2.setAttributeNS(null, 'width', 50); "
     circleIndicatorMarker2 = circleIndicatorMarker2&"   image2.setAttributeNS(null, 'height', 50); image2.setAttributeNS(null, 'href', 'q.gif'); svg2.appendChild(image2); "

end if 
if ((TimeVS>PeriodVS) and (PeriodVS>0)) then
	'circleIndicatorMarker3 = "circleIndicator3.renderer.image('q.gif', 25, 25, 50, 50).add();"

     circleIndicatorMarker3 = " var image3 = document.createElementNS('http://www.w3.org/2000/svg', 'image');"
     circleIndicatorMarker3 = circleIndicatorMarker3&"   image3.setAttributeNS(null, 'x', 25);  image3.setAttributeNS(null, 'y', 25); image3.setAttributeNS(null, 'width', 50); "
     circleIndicatorMarker3 = circleIndicatorMarker3&"   image3.setAttributeNS(null, 'height', 50); image3.setAttributeNS(null, 'href', 'q.gif'); svg3.appendChild(image3); "

end if 
'------------------------------------
'--END: Data CircleIndicator--------
'------------------------------------

VIP_Title = ""
sqlstr = "select VIP_Title from VIP_Config"
Rs.Open sqlstr, Conn
If not Rs.EOF then
	VIP_Title = Rs.Fields("VIP_Title")
end if
Rs.Close


'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------


	CurrentStratTime = DateTimeFormat(DateAdd("m", -1, DateAdd("h", -1*hoursCount, Now)), "yyyy, mm, dd, hh, nn")
	CurrentEndTime = DateTimeFormat(DateAdd("m", -1, Now), "yyyy, mm, dd, hh, nn")


'CurrentStratTime = DateTimeFormat(DateAdd("m", -1, Now), "yyyy, mm, dd, 0, 0")
CurrentTime = DateTimeFormat(DateAdd("m", -1, Now), "yyyy, mm, dd, hh, nn")
'CurrentTimeLabel = DateTimeFormat(DateAdd("m", -1, Now), "hh:nn")
CurrentTimeLabel = DateTimeFormat(DateAdd("m", -1, cDate(LastFileTimeFull)), "hh:nn")
'CurrentEndTime = DateTimeFormat(DateAdd("m", -1, Now), "yyyy, mm, dd, 23, 59")
%>
<!DOCTYPE HTML>
<html>
<head>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
                <meta http-equiv="X-UA-Compatible" content="ie=edge">

             <!--  <meta http-equiv='refresh' content='60; url=http://ufa-qos01ow/vsp/main4.asp'>-->

		<!-- 1. Add these JavaScript inclusions in the head of your page -->
		<script type="text/javascript" src="js/jquery.min.js"></script>
		<script type="text/javascript" src="js/highcharts.js"></script>
		<script type="text/javascript" src="js/themes/gray.js"></script>
		<!-- 2. Add the JavaScript to initialize the chart on document ready -->
		<script type="text/javascript">
		
			var chart1;
			var chart2;
			var circleIndicator1;
			//var circleIndicator2;
			var circleIndicator3;
			var circleIndicator4;
			
			var containerISS;
			var containerACQ;
			var containerATM;
			var container3DS;

			var FlagOut=1;
			
			var now = new Date(); 
			var now_utc =  Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(),  now.getHours(), now.getMinutes(), now.getSeconds());			
			

			$(document).ready(function() {

			    // Круговая диграмма 1
			    var svg1 = document.createElementNS("http://www.w3.org/2000/svg", "svg");
			    svg1.setAttributeNS(null, 'width', 200);
			    svg1.setAttributeNS(null, 'height', 200);

			    circleIndicator1 = document.createElementNS("http://www.w3.org/2000/svg", "circle");
			    circleIndicator1.setAttributeNS(null, 'cx', 100);
			    circleIndicator1.setAttributeNS(null, 'cy', 100);
			    circleIndicator1.setAttributeNS(null, 'r', 90);
			    circleIndicator1.setAttributeNS(null, 'style', 'fill: <% =circleIndicatorColor1 %>; stroke: <% =circleIndicatorColor1 %>;' );
			    svg1.appendChild(circleIndicator1);

			    /*circleIndicator1 = new Highcharts.Chart({
			        chart:   {renderTo: 'circleIndicator1', type: 'line', margin: [0, 0, 0, 0] },
			        credits: {enabled: false},	legend:  {enabled: false},	tooltip: {enabled: false},	title:   {text: ''}
			    });*/
			    /*circleIndicator1.renderer.circle(100, 100, 90).attr({
			        fill: '<% =circleIndicatorColor1 %>',
			        stroke: '<% =circleIndicatorColor1 %>'
			    }).add();*/
				
			    <%=circleIndicatorMarker1  %>

                
			    $("#circleIndicator1").html(svg1);
				
			    // Круговая диграмма 2
			    var svg2 = document.createElementNS("http://www.w3.org/2000/svg", "svg");
			    svg2.setAttributeNS(null, 'width', 100);
			    svg2.setAttributeNS(null, 'height', 100);
			    circleIndicator2 = document.createElementNS("http://www.w3.org/2000/svg", "circle");
			    circleIndicator2.setAttributeNS(null, 'cx', 50);
			    circleIndicator2.setAttributeNS(null, 'cy', 50);
			    circleIndicator2.setAttributeNS(null, 'r', 45);
			    circleIndicator2.setAttributeNS(null, 'style', 'fill: <% =circleIndicatorColor2 %>; stroke: <% =circleIndicatorColor2 %>;' );
			    svg2.appendChild(circleIndicator2);
			    /*var circleIndicator2 = new Highcharts.Chart({
			        chart:   {renderTo: 'circleIndicator2', type: 'line', margin: [0, 0, 0, 0] },
			        credits: {enabled: false},	legend:  {enabled: false},	tooltip: {enabled: false},	title:   {text: ''}
			    });
			    circleIndicator2.renderer.circle(50, 50, 45).attr({
			        fill: '<% =circleIndicatorColor2 %>',
			        stroke: '<% =circleIndicatorColor2 %>'
			    }).add();*/
				
			    <%=circleIndicatorMarker2 %>	
                
			    $("#circleIndicator2").html(svg2);
				
			    // Круговая диграмма 3
			    var svg3 = document.createElementNS("http://www.w3.org/2000/svg", "svg");
			    svg3.setAttributeNS(null, 'width', 100);
			    svg3.setAttributeNS(null, 'height', 100);
			    circleIndicator3 = document.createElementNS("http://www.w3.org/2000/svg", "circle");
			    circleIndicator3.setAttributeNS(null, 'cx', 50);
			    circleIndicator3.setAttributeNS(null, 'cy', 50);
			    circleIndicator3.setAttributeNS(null, 'r', 45);
			    circleIndicator3.setAttributeNS(null, 'style', 'fill: <% =circleIndicatorColor3 %>; stroke: <% =circleIndicatorColor3 %>;' );
			    svg3.appendChild(circleIndicator3);
			    /*circleIndicator3 = new Highcharts.Chart({
			        chart:   {renderTo: 'circleIndicator3', type: 'line', margin: [0, 0, 0, 0] },
			        credits: {enabled: false},	legend:  {enabled: false},	tooltip: {enabled: false},	title:   {text: ''}
			    });
			    circleIndicator3.renderer.circle(50, 50, 45).attr({
			        fill: '<% =circleIndicatorColor3 %>',
			        stroke: '<% =circleIndicatorColor3 %>'
			    }).add();*/
				
			    <%=circleIndicatorMarker3 %>	
                
			    $("#circleIndicator3").html(svg3);

		
				// Диграмма ISS
				containerISS = new Highcharts.Chart({
			        chart: {
			            renderTo: 'containerISS',
						ignoreHiddenSeries : false,
						marginTop: 40
			        },
					credits: {enabled: false},
			        legend:  {enabled: false},
			        tooltip: {enabled: false},
			        title:   {align: 'right',
								text: 'Эмиссия',
								y: 10
							},
			        subtitle: {align: 'left', 
								text: 'NSPK', 
								y: 10
							},
			        xAxis: {
			            min: Date.UTC(<%=CurrentStratTime %>),
			            max: Date.UTC(<%=CurrentEndTime %>),
			            type: 'datetime',
			            tickInterval: 1800*1000,
						gridLineWidth: 1,
						gridLineColor: 'rgba(255, 255, 255, 0.1)',
			            dateTimeLabelFormats: { // don't display the dummy year
			                hour: '%H:%M'
			            },
						plotLines: [{
										color: '#C0C0C0',
										width: 1,
										value: Date.UTC(<%=CurrentTime %>)
										
									}]
			        },
					yAxis: [{ // Primary yAxis
							min: 0,
							title: {
								text: ''
							}
						}, { // Secondary yAxis
							min: 0,
							title: {
								text: ''
							},
							opposite: true
					}],
					plotOptions: {
			            scatter: {
			                dataLabels: {
			                    enabled: false,
			                    align: 'right',
			                    style: { font: 'bold 24px Arial' },
			                    formatter: function() {	return this.point.name; }
			                },
			                marker: {
			                    enabled: true, 
			                    symbol: 'circle'
			                },
			                enableMouseTracking: false
			            }
					},
					series: [
<%
	Response.write ISS_VISA_Series
	Response.write ISS_VISA_Dot
	Response.write ISS_MC_Series
	Response.write ISS_MC_Dot
	Response.write ISS_NSPK_VISA_Series
	Response.write ISS_NSPK_VISA_Dot
	Response.write ISS_NSPK_MC_Series
	Response.write ISS_NSPK_MC_Dot
	Response.write ISS_MIR_Series
	Response.write ISS_MIR_Dot
%>					
					]
				});				

				
				// Диграмма ACQ
				containerACQ = new Highcharts.Chart({
			        chart: {
			            renderTo: 'containerACQ',
						ignoreHiddenSeries : false,
						marginTop: 40
			        },
					credits: {enabled: false},
			        legend:  {enabled: false},
			        tooltip: {enabled: false},
			        title:   {align: 'right', text: 'Эквайринг', y: 10},
			        subtitle: {align: 'left', text: 'NSPK', y: 10},
			        xAxis: {
			            min: Date.UTC(<%=CurrentStratTime %>),
			            max: Date.UTC(<%=CurrentEndTime %>),
			            type: 'datetime',
			            tickInterval: 1800*1000,
						gridLineWidth: 1,
						gridLineColor: 'rgba(255, 255, 255, 0.1)',
			            dateTimeLabelFormats: { // don't display the dummy year
			                hour: '%H:%M'
			            },
						plotLines: [{
										color: '#C0C0C0',
										width: 1,
										value: Date.UTC(<%=CurrentTime %>)
										
									}]
			        },
					yAxis: [{ // Primary yAxis
							min: 0,
							title: {
								text: ''
							}
						}, { // Secondary yAxis
							min: 0,
							title: {
								text: ''
							},
							opposite: true
					}],
					plotOptions: {
			            scatter: {
			                dataLabels: {
			                    enabled: false,
			                    align: 'right',
			                    style: { font: 'bold 24px Arial' },
			                    formatter: function() {	return this.point.name; }
			                },
			                marker: {
			                    enabled: true, 
			                    symbol: 'circle'
			                },
			                enableMouseTracking: false
			            }
					},
					series: [
<%
	Response.write ACQ_VISA_Series
	Response.write ACQ_VISA_Dot
	Response.write ACQ_MC_Series
	Response.write ACQ_MC_Dot
	Response.write ACQ_NSPK_VISA_Series
	Response.write ACQ_NSPK_VISA_Dot
	Response.write ACQ_NSPK_MC_Series
	Response.write ACQ_NSPK_MC_Dot
	Response.write ACQ_MIR_Series
	Response.write ACQ_MIR_Dot
%>					
					]
				});			


				
// Диграмма ATM
				containerATM = new Highcharts.Chart({
			        chart: {
			            renderTo: 'containerATM',
						ignoreHiddenSeries : false,
						marginTop: 40
			        },
					credits: {enabled: false},
			        legend:  {enabled: false},
			        tooltip: {enabled: false},
			        title:   {align: 'right', text: 'BPT/H2H_RBS', y: 10},
			        subtitle: {align: 'left', text: 'ATM/POS', y: 10},
			        xAxis: {
			            min: Date.UTC(<%=CurrentStratTime %>),
			            max: Date.UTC(<%=CurrentEndTime %>),
			            type: 'datetime',
			            tickInterval: 1800*1000*2,
						gridLineWidth: 1,
						gridLineColor: 'rgba(255, 255, 255, 0.1)',
			            dateTimeLabelFormats: { // don't display the dummy year
			                hour: '%H:%M'
			            },
						plotLines: [{
										color: '#C0C0C0',
										width: 1,
										value: Date.UTC(<%=CurrentTime %>)
									}]
			        },
					yAxis: [{ // Primary yAxis
							min: 0,
							title: {
								text: ''
							}
						}, { // Secondary yAxis
							min: 0,
							title: {
								text: ''
							},
							opposite: true
					}],
					plotOptions: {
			            scatter: {
			                dataLabels: {
			                    enabled: false,
			                    align: 'right',
			                    style: { font: 'bold 24px Arial' },
			                    formatter: function() {	return this.point.name; }
			                },
			                marker: {
			                    enabled: true, 
			                    symbol: 'circle'
			                },
			                enableMouseTracking: false
			            }
					},
					series: [
<%
	Response.write All_ATM_Series
	Response.write All_ATM_Dot
	Response.write All_BPT_Series
	Response.write All_BPT_Dot
	Response.write All_POS_Series
	Response.write All_POS_Dot
	Response.write All_H2H_RBS_Series
	Response.write All_H2H_RBS_Dot
%>					
					]
				});					
				
				// Диграмма 3DS
				container3DS = new Highcharts.Chart({
			        chart: {
			            renderTo: 'container3DS',
						ignoreHiddenSeries : false,
						marginTop: 40
			        },
					credits: {enabled: false},
			        legend:  {enabled: false},
			        tooltip: {enabled: false},
			        title:   {align: 'right', text: '3DS', y: 10},
			        subtitle: {align: 'left', text: 'SOA', y: 10},
			        xAxis: {
			            min: Date.UTC(<%=CurrentStratTime %>),
			            max: Date.UTC(<%=CurrentEndTime %>),
			            type: 'datetime',
			            tickInterval: 1800*1000*2,
						gridLineWidth: 1,
						gridLineColor: 'rgba(255, 255, 255, 0.1)',
			            dateTimeLabelFormats: { // don't display the dummy year
			                hour: '%H:%M'
			            },
						plotLines: [{
										color: '#C0C0C0',
										width: 1,
										value: Date.UTC(<%=CurrentTime %>)
									}]
			        },
					yAxis: [{ // Primary yAxis
							min: 0,
							title: {
								text: ''
							}
						}, { // Secondary yAxis
							min: 0,
							title: {
								text: ''
							},
							opposite: true
					}],
					plotOptions: {
			            scatter: {
			                dataLabels: {
			                    enabled: false,
			                    align: 'right',
			                    style: { font: 'bold 24px Arial' },
			                    formatter: function() {	return this.point.name; }
			                },
			                marker: {
			                    enabled: true, 
			                    symbol: 'circle'
			                },
			                enableMouseTracking: false
			            }
					},
					series: [
<%
	Response.write All_3DS_Series
	Response.write All_3DS_Dot
	Response.write All_SOA_Series
	Response.write All_SOA_Dot
%>					
					]
				});	
				
				
			});

		</script>
		
	<style type="text/css">

	
	
	.hc-label {
		  color: #00FFFF;
		}
		
	.highcharts-title {
		fill: #ffffff;
		font-weight: bold;
		font-size: 18px;
	}
	
	.highcharts-subtitle {
		font-size: 18px !important;
		font-weight: bold !important;
		fill: #ffffff !important;
	}
	
	A {
	    color: inherit;
		text-decoration: none;
	}
	<!--
	BODY {
		margin: 0px;
		background-color: #242424;
	}
TABLE {
	margin: 0px;
	padding: 0px;
}
TD {
	padding-top: 1px;
	padding-bottom: 1px;
	text-align: center;
	color: #FFFFFF;
	font-family: Verdana, Arial, helvetica, sans-serif, Geneva;
}
TD.Head {
	color: #000000;
	font-size: 24pt;
}
TD.Txt {
	color: #FFFFFF;
	font-size: 36pt;
	font-weight: 700;
}
	
	-->
	</style>
</head>
<body>
<div align="center">
<table border="0" padding="0"  width="1920px" height="1080px" style="border: none;">
<tr style="padding: 0;">
	<td width="300px" style="padding: 0;" ></td>
	<td width="325px" style="padding: 0;"></td>
	<td width="325px" style="padding: 0;" ></td>
	<td width="300px" style="padding: 0;"></td>
	<td width="325px" style="padding: 0;"></td>
	<td width="325px" style="padding: 0;"></td>
</tr>
	<tr>
		<td width="625px" style="border: none;" colspan="2" >
	
		    <table border="0" height="718px" width="100%" cellspacing="0" >
			  <tr>
				<td rowspan="2" height="210px" width="210px" style="border-top: solid 1px #4572A7;  border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;  text-align: center">
					<div id="circleIndicator1"  style="width: 200px; height: 200px; margin-left: 6px; margin-top: 0;"></div>
				</td>
				<td height="30px" colspan="2" style="border-top: solid 1px #4572A7; border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;  text-align: center; font-size: 20pt; font-weight: 400; "><%=LastUpdateVO %></td>
			
			  </tr>
			  <tr>
				<td height="180px" style=" border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;  text-align: center; font-size: 40pt; font-weight: 600;">ISS</td>
				<td style=" border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;  text-align: center;  font-size: 40pt; font-weight: 600;">ACQ</td>
			 
			  </tr>
			  <tr>
				<td style="color: <%=VISA_Series_Color %>; border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;  text-align: center; font-size: 28pt; font-weight: 400;"><a target="_blank" href="detail.asp?T=7&Param=ISS_VISA">VISA</a></td>
				<td style="<% if (ISS_VISA_Color<>"") then 
					response.write "background: "&ISS_VISA_Color&";" 
					if (ISS_VISA_Color=clWarning) then response.write "color: #000000;"  end if
				end if %>  border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;  text-align: center; font-size: 28pt; font-weight: 400;"><% =ISS_VISA %></td>
				<td style="<% if (ACQ_VISA_Color<>"") then 
					response.write "background: "&ACQ_VISA_Color&";" 
					if (ACQ_VISA_Color=clWarning) then response.write "color: #000000;"  end if
				end if %>  border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;  text-align: center; font-size: 28pt; font-weight: 400;"><% =ACQ_VISA %></td>
			  </tr><tr>
				<td style="color: <%=MC_Series_Color %>; border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;   text-align: center; font-size: 28pt; font-weight: 400;"><a target="_blank" href="detail.asp?T=7&Param=ISS_MC">MC</a></td>
				<td style="<% if (ISS_MC_Color<>"") then 
					response.write "background: "&ISS_MC_Color&";"
					if (ISS_MC_Color=clWarning) then response.write "color: #000000;"  end if
				end if %>  border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;   text-align: center; font-size: 28pt; font-weight: 400;"><% =ISS_MC %></td>
				<td style="<% if (ACQ_MC_Color<>"") then 
					response.write "background: "&ACQ_MC_Color&";"
					if (ACQ_MC_Color=clWarning) then response.write "color: #000000;"  end if
				end if %>  border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;   text-align: center; font-size: 28pt; font-weight: 400;"><% =ACQ_MC %></td>
              </tr><tr>
				<td style="color: <%=NSPK_VISA_Series_Color %>; border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;   text-align: center; font-size: 28pt; font-weight: 400;"><a target="_blank" href="detail.asp?T=7&Param=ISS_NSPK_VISA">NSPK_VISA</a></td>
				<td style="<% if (ISS_NSPK_VISA_Color<>"") then 
					response.write "background: "&ISS_NSPK_VISA_Color&";"
					if (ISS_NSPK_VISA_Color=clWarning) then response.write "color: #000000;"  end if
				end if %> border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;   text-align: center; font-size: 28pt; font-weight: 400;"><% =ISS_NSPK_VISA %></td>
				<td style="<% if (ACQ_NSPK_VISA_Color<>"") then 
					response.write "background: "&ACQ_NSPK_VISA_Color&";"
					if (ACQ_NSPK_VISA_Color=clWarning) then response.write "color: #000000;"  end if					
					end if %> border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;   text-align: center; font-size: 28pt; font-weight: 400;"><% =ACQ_NSPK_VISA %></td>
              </tr><tr>
				<td style="color: <%=NSPK_MC_Series_Color %>; border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;   text-align: center; font-size: 28pt; font-weight: 400;"><a target="_blank" href="detail.asp?T=7&Param=ISS_NSPK_MC">NSPK_MC</a></td>
				<td style="<% if (ISS_NSPK_MC_Color<>"") then 
					response.write "background: "&ISS_NSPK_MC_Color&";"
					if (ISS_NSPK_MC_Color=clWarning) then response.write "color: #000000;"  end if
				end if %>  border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;   text-align: center; font-size: 28pt; font-weight: 400;"><% =ISS_NSPK_MC %></td>
				<td style="<% if (ACQ_NSPK_MC_Color<>"") then 
					response.write "background: "&ACQ_NSPK_MC_Color&";"
					if (ACQ_NSPK_MC_Color=clWarning) then response.write "color: #000000;"  end if
				end if %>  border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;   text-align: center; font-size: 28pt; font-weight: 400;"><% =ACQ_NSPK_MC %></td>
              </tr><tr>		
				<td style="color: <%=MIR_Series_Color %>; border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;   text-align: center; font-size: 28pt; font-weight: 400;"><a target="_blank" href="detail.asp?T=7&Param=ISS_MIR">MIR</a></td>
				<td style="<% if (ISS_MIR_Color<>"") then 
					response.write "background: "&ISS_MIR_Color&";"
					if (ISS_MIR_Color=clWarning) then response.write "color: #000000;"  end if
				end if %>  border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;   text-align: center; font-size: 28pt; font-weight: 400;"><% =ISS_MIR %></td>
				<td style="<% if (ACQ_MIR_Color<>"") then 
					response.write "background: "&ACQ_MIR_Color&";"
					if (ACQ_MIR_Color=clWarning) then response.write "color: #000000;"  end if
				end if %>  border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;   text-align: center; font-size: 28pt; font-weight: 400;"><% =ACQ_MIR %></td>
              </tr>				  
			</table>

		</td>
		<td width="1350px" style="border: none;" colspan="4">
			<div id="containerISS"  style="width: 1280px; height: 359px; margin: 0 auto"></div>
			<div id="containerACQ"  style="width: 1280px; height: 359px; margin: 0 auto"></div>
		</td>
	</tr>
	<tr>
		<td style="border: none;" width="300px" >
	
		    <table border="0" height="350px" width="100%" cellspacing="0" >
			  <tr>
				<td height="150px" width="150px"  style="border-top: solid 1px #4572A7;  border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7; text-align: center">
					<div id="circleIndicator2"  style="width: 100px; height: 100px; margin-left: 25px; margin-top: 0;"></div>
				</td>
				<td  style="border-top: solid 1px #4572A7; border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7; text-align: center; font-size: 40pt; font-weight: 600;">All</td>
			  </tr>
			  <tr>
				<td style="color: <%=All_ATM_Series_Color %>;  border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7; text-align: center; font-size: 20pt; font-weight: 400;"><a target="_blank" href="detail.asp?T=8&Param=All_ATM">ATM</a></td>
				<td style="<% if (All_ATM_Color<>"") then 
					response.write "background: "&All_ATM_Color&";"
					if (All_ATM_Color=clWarning) then response.write "color: #000000;"  end if
					end if %>  border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7; text-align: center; font-size: 20pt; font-weight: 400;"><% =All_ATM %></td>
			  </tr><tr>
				<td style="color: <%=All_BPT_Series_Color %>; border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;  text-align: center; font-size: 20pt; font-weight: 400;"><a target="_blank" href="detail.asp?T=8&Param=All_BPT">BPT</a></td>
				<td style="<% if (All_BPT_Color<>"") then 
					response.write "background: "&All_BPT_Color&";"
					if (All_BPT_Color=clWarning) then response.write "color: #000000;"  end if
					end if %> border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;  text-align: center; font-size: 20pt; font-weight: 400;"><% =All_BPT %></td>
			  </tr><tr>
				<td style="color: <%=All_POS_Series_Color %>; border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;  text-align: center; font-size: 20pt; font-weight: 400;"><a target="_blank" href="detail.asp?T=8&Param=All_POS">POS</a></td>
				<td style="<% if (All_POS_Color<>"") then 
					response.write "background: "&All_POS_Color&";"
					if (All_POS_Color=clWarning) then response.write "color: #000000;"  end if
					end if %>  border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;  text-align: center; font-size: 20pt; font-weight: 400;"><% =All_POS %></td>
              </tr><tr>
				<td style="color: <%=All_H2H_RBS_Series_Color %>; border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7;  text-align: center; font-size: 20pt; font-weight: 400;"><a target="_blank" href="detail.asp?T=8&Param=All_H2H_RBS">H2H_RBS</a></td>
				<td style="<% if (All_H2H_RBS_Color<>"") then 
					response.write "background: "&All_H2H_RBS_Color&";"
					if (All_H2H_RBS_Color=clWarning) then response.write "color: #000000;"  end if					
					end if %>  border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7; text-align: center; font-size: 20pt; font-weight: 400;"><% =All_H2H_RBS %></td>
              </tr>				  
			</table>

		</td>
		<td style="border: none;"  width="650px" colspan="2" >
			<div id="containerATM"  style="width: 650px; height: 350px; margin: 0 auto"></div>
		</td>
		<td style="border: none;"  width="300px"  >
	
		    <table border="0" height="350px"  width="100%" cellspacing="0" >
			  <tr>
				<td height="150px" width="150px" style="border-top: solid 1px #4572A7; border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7; text-align: center">
					<div id="circleIndicator3"  style="width: 100px; height: 100px; margin-left: 25px; margin-top: 0;"></div>
				</td>
				<td  style=" border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-top: solid 1px #4572A7; border-bottom: solid 1px #4572A7; text-align: center; font-size: 40pt; font-weight: 600;">All</td>
				
			  </tr>

			  <tr>
				<td style="color: <%=All_3DS_Series_Color %>;border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7; text-align: center; font-size: 20pt; font-weight: 400;"><a target="_blank" href="detail.asp?T=9">3DS</a></td>
				<td style="<% if (All_3DS_Color<>"") then 
					response.write "background: "&All_3DS_Color&";"
					if (All_3DS_Color=clWarning) then response.write "color: #000000;"  end if	
					end if %> border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7; text-align: center; font-size: 20pt; font-weight: 400;"><% =All_3DS %></td>
			  </tr><tr>
				<td style="color: <%=All_SOA_Series_Color %>;border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7; text-align: center; font-size: 20pt; font-weight: 400;"><a target="_blank" href="detail.asp?T=9">SOA</a></td>
				<td style="<% if (All_SOA_Color<>"") then 
					response.write "background: "&All_SOA_Color&";"
					if (All_SOA_Color=clWarning) then response.write "color: #000000;"  end if	
					end if %> border-left: solid 1px #4572A7; border-right: solid 1px #4572A7; border-bottom: solid 1px #4572A7; text-align: center; font-size: 20pt; font-weight: 400;"><% =All_SOA %></td>
			  </tr>				  
			</table>

		</td>
		<td style="border: none;"  width="650px" colspan="2"  >
			<div id="container3DS"  style="width: 650px; height: 350px; margin: 0 auto"></div>
		</td>		
	</tr>

</table>
</div>
</body>
</html>
<%

  Conn.Close
  set Conn = Nothing
  set Rs = Nothing

end if
%>
<!-- Разработка: Берников И.П. -->
<!-- Для вывода графиков используется библиотека Highcharts JS - http://highsoft.com/ -->
