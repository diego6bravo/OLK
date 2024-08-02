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
<!--#include file="../lcidReturn.inc" -->
<%

LogNum = Session("ItmRetVal")
LineID = Request.Form("LineID")
Field = Request.Form("Field")
FieldType = Request.Form("FieldType")
Value = Request.Form("Value")
ProcType = Request.Form("ProcType")
IgnoreOK = False

Select Case ProcType
	Case "Cmb"
		Select Case Field
			Case "EnableCombo"
				cmd.CommandText = "DBOLKItmEnableCombos" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LogNum") = LogNum
				cmd("@Active") = Value
				cmd.execute()
			Case Else
				cmd.CommandText = "DBOLKItmSaveComboData" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LogNum") = LogNum
				cmd("@Field") = Field
				cmd("@FieldType") = FieldType
				
				If Value <> "" Then
					Select Case FieldType
						Case "S"
							cmd("@ValueText") = Value
						Case "N"
							cmd("@ValueNumber") = CDbl(getNumericOut(Value))
						Case "I"
							cmd("@ValueInt") = CLng(Value)
						Case "D"
							cmd("@ValueDate") = SaveCmdDate(Value)
					End Select
				End If
				
				cmd.execute()
		End Select
	Case "CmbComp"	
		cmd.CommandText = "DBOLKItmSaveComboCompData" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@LineID") = LineID
		cmd("@Field") = Field
		cmd("@FieldType") = FieldType
		
		If Value <> "" Then
			Select Case FieldType
				Case "S"
					cmd("@ValueText") = Value
				Case "N"
					cmd("@ValueNumeric") = CDbl(getNumericOut(Value))
				Case "I"
					cmd("@ValueInt") = CLng(Value)
				Case "D"
					cmd("@ValueDate") = SaveCmdDate(Value)
			End Select
		End If
				
		cmd.execute()
	Case "CheckFilterValue"
		FilterType = CInt(Request("FilterType"))
		If Request("Checked") = "Y" Then Checked = "Y" Else Checked = "N"
		cmd.CommandText = "DBOLKItmSaveFilterVal" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@FilterType") = FilterType
		Select Case FilterType
			Case 1, 2, 3
				cmd("@FilterID") = Value
			Case 4
				cmd("@FilterValue") = Value
		End Select
		cmd("@Checked") = Checked
		cmd.execute()
	Case "CheckUdfFilterValue"
		If Request("Checked") = "Y" Then Checked = "Y" Else Checked = "N"
		
		cmd.CommandText = "DBOLKItmSaveUdfFilterVal" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LogNum") = LogNum
		cmd("@FieldID") = Request.Form("FieldID")
		cmd("@FilterValue") = Value
		cmd("@Checked") = Checked
		cmd.execute()
	Case "GetUDFData"
		cmd.CommandText = "DBOLKGetUDFValues" & Session("ID")
		cmd.Parameters.Refresh()
		cmd("@LanID") = Session("LanID")
		cmd("@TableID") = "OITM"
		cmd("@FieldID") = Request.Form("FieldID")
		rs.open cmd, , 3, 1
		retVal = ""
		do while not rs.eof
			If retVal <> "" Then retVal = retVal & "{V}"
			retVal = retVal & rs(0) & "{S}" & rs(1)
		rs.movenext
		loop
		IgnoreOK = True
		Response.Write retVal
	Case Else
		Select Case Field
			Case "ShowImage"
				cmd.CommandText = "DBOLKItmSaveShowData" & Session("ID")
				cmd.Parameters.Refresh()
				cmd("@LogNum") = LogNum
				cmd("@ShowImage") = Value
				cmd.execute()
			Case "Filter"
				Query = Request.Form("Query")
				
				sql = "select top 1 '' from OITM "
				
				If Query <> "" Then sql = sql & " where " & Query
				
				On Error Resume Next
				set rs = conn.execute(sql)
				
				If Err.Number = 0 Then
					cmd.CommandText = "DBOLKItmSaveFilter" & Session("ID")
					cmd.Parameters.Refresh()
					cmd("@LogNum") = LogNum
					If Query <> "" Then cmd("@Filter") = Query
					cmd.execute() 
					conn.execute(sql)
				Else
					IgnoreOK = True
					Response.Write "Err{S}" & Err.Description
				End If
			Case "ItemCode"
				cmd.CommandText = "DBOLKItmSaveItemCode" & Session("ID")
				cmd("@LogNum") = LogNum
				cmd("@ItemCode") = Value
				set rs = cmd.execute()
				Response.Write rs(0)
				IgnoreOK = True
			Case Else
				cmd.CommandText = "DBOLKProcessAJAX"
				cmd.Parameters.Refresh()
				cmd("@dbID") = Session("ID")
				cmd("@LogNum") = LogNum
				cmd("@TableID") = "TITM"
				cmd("@FieldID") = Field
				cmd("@FieldType") = FieldType
				If Value <> "" Then
					Select Case FieldType
						Case "S"
							cmd("@ValueText") = Value
						Case "N"
							cmd("@ValueNumeric") = CDbl(getNumericOut(Value))
						Case "I"
							cmd("@ValueInt") = CLng(getNumeric(Value))
						Case "D"
							cmd("@ValueDate") = SaveCmdDate(Value)
						Case "T"
							cmd("@ValueDate") = SaveCmdTime(Value)
							cmd("@FieldType") = "D"
					End Select
				End If
				cmd.execute()
			End Select
		
		
End Select

If Not IgnoreOK Then Response.Write "ok"

%>