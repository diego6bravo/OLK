<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>
<!--#include file="../chkLogin.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="adminTradSave.asp"-->
<%        
		Select Case Request("pollCmd")
			Case "D"
				sql = "update olkPoll set pollStatus = 'D' where pollIndex = " & Request("pollIndex")
				redirVar = "../pollman.asp"
			Case "delPoll"
				sql = "update olkPoll set pollStatus = 'D' where pollIndex in (" & Request("delID") & ")"
				redirVar = "../pollman.asp"
			Case "addPoll"
		        Set rs = Server.CreateObject("ADODB.RecordSet")
				sql = "select ISNULL(max(pollIndex)+1,1) pollIndex from olkpoll"
				set rs = conn.execute(sql)
				
				If Request("pollTitleTrad") <> "" Then
					SaveNewTrad Request("pollTitleTrad"), "Poll", "pollIndex", "alterPollTitle", rs("pollIndex")	
				End If

				randomize()
				randomNumber=Int(7 * rnd())
				sql = "insert olkPoll(pollIndex, pollName, pollTitle, pollDate, pollStatus) " & _
					  "values(" & rs("pollIndex") & ", N'" & saveHTMLDecode(Request("pollName"), False) & "', N'" & _
					  saveHTMLDecode(Request("pollTitle"), False) & "', Convert(datetime,'" & SaveSqlDate(Request("pollDate")) & "',120), 'C') " & _
					  " insert OLKPollLines(pollIndex, pollLineNum, LineText, colorIndex) " & _
					  "values(" & rs("pollIndex") & ", 1, N'(---)', " & randomNumber & ")"
	        	If Request("btnApply") <> "" Then
					redirVar = "../pollEdit.asp?pollIndex=" & rs("pollIndex")				
				Else
					redirVar = "../pollMan.asp"
				End If
				
				set rs = nothing
			Case "updatePoll"
				If Request("pollStatus") = "O" Then pollStatus = "O" Else pollStatus = "C"
				sql = "update olkPoll set pollTitle = N'" & saveHTMLDecode(Request("pollTitle"), False) & "', pollName = N'" & saveHTMLDecode(Request("pollName"), False) & _
					  "', pollDate = Convert(datetime,'" & SaveSqlDate(Request("pollDate")) & "',120), pollStatus = '" & pollStatus & "' " & _
					  "where pollIndex = " & Request("pollIndex")
		        Set rs = Server.CreateObject("ADODB.RecordSet")
		        sqlt = "select pollLineNum from olkPollLines where pollIndex = " & Request("pollIndex")
		        set rs = conn.execute(sqlt)
	        	do while not rs.eof
	        		sql = sql & " update olkPollLines set LineText = N'" & saveHTMLDecode(Request("opt" & rs("pollLineNum")), False) & "', " & _
	        					"colorIndex = " & Request("cIndex" & rs("pollLineNum")) & ", lineOrder = " & Request("order" & rs("pollLineNum")) & " where pollIndex = " & _
	        					Request("pollIndex") & " and pollLineNum = " & rs("pollLineNum")
	        	rs.movenext
	        	loop
	        	If Request("optNew") <> "" Then
	        		set rs = conn.execute("select IsNull((select Max(pollLineNum)+1 from OLKPollLines where pollIndex = " & Request("pollIndex") & "), 0)")
	        		pollLineNum = rs(0)
	        		sql = sql & "insert OLKPollLines(pollIndex, pollLineNum, LineText, colorIndex, lineOrder) " & _
	        					"values(" & Request("pollIndex") & ", " & pollLineNum & ", N'" & saveHTMLDecode(Request("optNew"), False) & "', " & Request("cIndexNew") & ", " & Request("orderNew") & ") "
	        	End If
	        	If Request("btnApply") <> "" Then
					redirVar = "../pollEdit.asp?pollIndex=" & Request("pollIndex")
				Else
					redirVar = "../pollman.asp"
				End If
			Case "del"
				sql = "delete OLKPollLines where pollIndex = " & Request("pollIndex") & " and pollLineNum = " & Request("pollLineNum") & _
						" delete OLKPollLinesAlterNames where pollIndex = " & Request("pollIndex") & " and pollLineNum = " & Request("pollLineNum")
				redirVar = "../pollEdit.asp?pollIndex=" & Request("pollIndex")
			Case "addLine"
				sql = ""
				randomize()
				randomNumber=Int(7 * rnd())
				If Request("p") = "f" then
					sql = "update OLKPollLines set pollLineNum = pollLineNum+1 where pollIndex = " & Request("pollIndex") & _
						  " insert OLKPollLines(pollIndex, pollLineNum, LineText, colorIndex) " & _
						  "values(" & Request("pollIndex") & ", 1, N'(---)', " & randomNumber & ")"
				Else
					sql = "declare @pollLineNum int set @pollLineNum = ISNULL((select max(pollLineNum)+1 from olkPollLines where pollIndex = " & Request("pollIndex") & "),1) " & _
						  " insert OLKPollLines(pollIndex, pollLineNum, LineText, colorIndex) " & _
						  "values(" & Request("pollIndex") & ", @pollLineNum, N'(---)', " & randomNumber & ")"
				End If			
			redirVar = "../pollEdit.asp?pollIndex=" & Request("pollIndex")
		End Select 
	conn.execute(sql)
	
	If Request("pollCmd") = "updatePoll" and Request("optNewTrad") <> "" Then
		SaveNewTrad Request("optNewTrad"), "PollLines", "pollIndex,pollLineNum", "AlterLineText", Request("pollIndex") & "," & pollLineNum
	End If
	
	conn.close 
	response.redirect redirVar %>