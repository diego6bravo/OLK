<!--#include file="../chkLogin.asp" -->
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="adminTradSave.asp"-->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

Select Case Request("cmd")
	Case "del"
		sql = "update OLKADPoll set Status = 'R' where AdPollID in (" & Request("delID") & ")"
		conn.execute(sql)
	Case "delEnc"
		sql = "update OLKADPoll set Status = 'R' where AdPollID = " & Request("delIndex")
		conn.execute(sql)
	Case "editQC"
		If Request("Description") = "" Then Desc = "NULL" Else Desc = "N'" & saveHTMLDecode(Request("Description"), False) & "'"
		If Request("Filter") = "" Then qFilter = "NULL" Else qFilter = "N'" & saveHTMLDecode(Request("Filter"), False) & "'"
		If Request("Status") = "A" Then Status = "A" Else Status = "N"
		
		If Request("AdPollID") = "" Then
			sql = "declare @AdPollID int set @AdPollID = IsNull((select Max(AdPollID)+1 from OLKADPoll), 0) " & _
					"select @AdPollID AdPollID " & _
					"insert OLKADPoll(AdPollID, Name, Description, StartDate, EndDate, Filter, Status, UserSign) " & _
					"values(@AdPollID, N'" & Request("Name") & "', " & Desc & ", Convert(datetime,'" & SaveSqlDate(Request("StartDate")) & "',120), Convert(datetime,'" & SaveSqlDate(Request("EndDate")) & "',120), " & qFilter & ", '" & Status & "', " & Session("vendid") & ")"
			set rs = Server.CreateObject("ADODB.RecordSet")
			set rs = conn.execute(sql)
			AdPollID = rs(0)
		Else
			sql = "update OLKADPoll set Name = N'" & Request("Name") & "', Description = " & Desc & ", StartDate = Convert(datetime,'" & SaveSqlDate(Request("StartDate")) & "',120), EndDate = Convert(datetime,'" & SaveSqlDate(Request("EndDate")) & "',120), " & _
					"Filter = " & qFilter & ", Status = '" & Status & "' where AdPollID = " & Request("AdPollID")
			conn.execute(sql)
			AdPollID = Request("AdPollID")
		End If
		
		sql = "delete OLKADPollAgents where AdPollID = " & AdPollID
		conn.execute(sql)
		arrAgents = Split(Request("Agents"), ", ")
		sql = ""
		For i = 0 to UBound(arrAgents)
			sql = sql & "insert OLKADPollAgents(AdPollID, SlpCode) values(" & AdPollID & ", " & arrAgents(i) & ") "
		Next
		If sql <> "" Then conn.execute(sql)
		
		If Request("NameTrad") <> "" Then
			SaveNewTrad Request("NameTrad"), "ADPoll", "AdPollID", "alterName", AdPollID
		End If
		If Request("DescriptionTrad") <> "" Then
			SaveNewTrad Request("DescriptionTrad"), "ADPoll", "AdPollID", "alterDescription", AdPollID
		End If

		If Request("btnApply") <> "" Then Response.Redirect "../extPollEdit.asp?AdPollID=" & AdPollID
	Case "editQuestion"
	
		If Request("MandatoryNote") = "Y" Then MandatoryNote = "Y" Else MandatoryNote = "N"
	
		If Request("LineID") = "" Then
			sql = "declare @LineID int set @LineID = IsNull((select Max(LineID)+1 from OLKADPollLines where AdPollID = " & Request("AdPollID") & "), 0) " & _
				"select @LineID LineID " & _
				"insert OLKADPollLines(AdPollID, LineID, Question, Type, MandatoryNote, Ordr) " & _
				"values(" & Request("AdPollID") & ", @LineID, N'" & saveHTMLDecode(Request("Question"), False) & "', '" & Request("Type") & "', '" & MandatoryNote & "', " & Request("Ordr") & ")"
			set rs = Server.CreateObject("ADODB.RecordSet")
			set rs = conn.execute(sql)
			LineID = rs("LineID")
		Else
			sql = "update OLKADPollLines set Question = N'" & saveHTMLDecode(Request("Question"), False) & "', Type = '" & Request("Type") & "', MandatoryNote = '" & MandatoryNote & "', Ordr = " & Request("Ordr") & " " & _
					"where AdPollID = " & Request("AdPollID") & " and LineID = " & Request("LineID")
			conn.execute(sql)
			LineID = Request("LineID")
		End If
		
		If Request("QuestionTrad") <> "" Then
			SaveNewTrad Request("QuestionTrad"), "ADPollLines", "AdPollID,LineID", "alterQuestion", Request("AdPollID") & "," & LineID
		End If
		
		If Request("btnApply") <> "" Then
			Response.Redirect "../extPollEdit.asp?AdPollID=" & Request("AdPollID") & "&editIndex=" & LineID & "#tblEditLine"
		Else
			Response.Redirect "../extPollEdit.asp?AdPollID=" & Request("AdPollID")
		End If
	Case "updateQuestions"
		set rs = Server.CreateObject("ADODB.RecordSet")
		sql = "select LineID from OLKADPollLines where AdPollID = " & Request("AdPollID")
		set rs = conn.execute(sql)
		sql = ""
		do while not rs.eof
			sql = sql & "update OLKAdPollLines set Ordr = " & Request("Ordr" & rs("LineID")) & " where AdPollID = " & Request("AdPollID") & " and LineID = " & rs("LineID") & " "
		rs.movenext
		loop 
		If sql <> "" Then conn.execute(sql)
		Response.Redirect "../extPollEdit.asp?AdPollID=" & Request("AdPollID") & "#tblQuestions"
	Case "delQuestion"
		sql = "delete OLKADPollLines where AdPollID = " & Request("AdPollID") & " and LineID = " & Request("LineID") & " " & _
				"delete OLKADPollLinesChoices where AdPollID = " & Request("AdPollID") & " and LineID = " & Request("LineID") & " " & _
				"delete OLKADPollLinesChoicesAlterNames where AdPollID = " & Request("AdPollID") & " and LineID = " & Request("LineID") & " "
		conn.execute(sql)
		Response.Redirect "../extPollEdit.asp?AdPollID=" & Request("AdPollID") & "#tblQuestions"
	Case "updateChoice"
		set rs = Server.CreateObject("ADODB.RecordSet")
		sql = "select ChoiceID from OLKADPollLinesChoices where AdPollID = " & Request("AdPollID") & " and LineID = " & Request("editChoice")
		set rs = conn.execute(sql)
		sql = ""
		do while not rs.eof
			sql = sql & "update OLKADPollLinesChoices set Choice = N'" & Request("Choice" & rs("ChoiceID")) & "', Ordr = " & Request("choiceOrdr" & rs("ChoiceID")) & ", Color = '" & Request("ForeColor" & rs("ChoiceID")) & "' " & _
					"where AdPollID = " & Request("AdPollID") & " and LineID = " & Request("editChoice") & " and ChoiceID = " & rs("ChoiceID") & " "
		rs.movenext
		loop
		
		If sql <> "" Then conn.execute(sql)
		
		If Request("ChoiceNew") <> "" Then
			sql = "declare @ChoiceID int set @ChoiceID = IsNull((select Max(ChoiceID)+1 from OLKADPollLinesChoices where AdPollID = " & Request("AdPollID") & " and LineID = " & Request("editChoice") & "), 0) " & _
			"select @ChoiceID ChoiceID " & _
			"insert OLKADPollLinesChoices(AdPollID, LineID, ChoiceID, Choice, Color, Ordr) " & _
			"values(" & Request("AdPollID") & ", " & Request("editChoice") & ", @ChoiceID, N'" & Request("ChoiceNew") & "', '" & Request("ForeColorNew") & "', " & Request("choiceOrdrNew") & ") "
			set rs = conn.execute(sql)
			
		
			If Request("ChoiceNewTrad") <> "" Then
				SaveNewTrad Request("ChoiceNewTrad"), "ADPollLinesChoices", "AdPollID,LineID,ChoiceID", "alterChoice", Request("AdPollID") & ", " & Request("editChoice") & ", " & rs("ChoiceID")
			End If
		End If
		
		If Request("btnApply") <> "" Then
			Response.Redirect "../extPollEdit.asp?AdPollID=" & Request("AdPollID") & "&editChoice=" & Request("editChoice") & "&#editChoice"
		Else
			Response.Redirect "../extPollEdit.asp?AdPollID=" & Request("AdPollID") & "#tblQuestions"
		End If
	Case "delChoice"
		sql = "delete OLKADPollLinesChoices where AdPollID = " & Request("AdPollID") & " and LineID = " & Request("LineID") & " and ChoiceID = " & Request("ChoiceID") & " " & _
				"delete OLKADPollLinesChoicesAlterNames where AdPollID = " & Request("AdPollID") & " and LineID = " & Request("LineID") & " and ChoiceID = " & Request("ChoiceID") & " "
		conn.execute(sql)
			Response.Redirect "../extPollEdit.asp?AdPollID=" & Request("AdPollID") & "&editChoice=" & Request("LineID") & "&#editChoice"
End Select


Response.Redirect "../extpollman.asp"

%>
