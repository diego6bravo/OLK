<%@ Language=VBScript %> 
<% If session("OLKDB") = "" Then response.redirect "../lock.asp" %>
<!--#include file="../myHTMLEncode.asp"-->
<!--#include file="../lcidReturn.inc"-->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

set rs = Server.CreateObject("ADODB.recordset")

Select Case Request("cmd")
	Case "cartExtra"
		ObjType = CInt(Request("ObjType"))
		If ObjType = -13 Then ObjType = 13
		If Request("NumAtCard") <> "" Then NumAtCard = "N'" & saveHTMLDecode(Request("NumAtCard"), False) & "'" Else NumAtCard = "NULL"
		If Request("Comments") <> "" Then Comments = "N'" & Request("Comments") & "'" Else Comments = "NULL"
		If Request("DocDueDate") <> "" Then DocDueDate = "Convert(datetime,'" & SaveSqlDate(Request("DocDueDate")) & "',120)" Else DocDueDate = "NULL"
		If Request("PartSupply") = "Y" Then PartSupply = "Y" Else PartSupply = "N"
		If Request("PayToCode") <> "" Then PayToCode = "N'" & saveHTMLDecode(Request("PayToCode"), False) & "'" Else PayToCode = "NULL"
		If Request("ShipToCode") <> "" Then ShipToCode = "N'" & saveHTMLDecode(Request("ShipToCode"), False) & "'" Else ShipToCode = "NULL"
		If Request("Project") <> "" Then Project = "N'" & saveHTMLDecode(Request("Project"), False) & "'" Else Project = "NULL"
		If Request("DocCur") <> "" Then DocCur = "N'" & saveHTMLDecode(Request("DocCur"), False) & "'" Else DocCur = "DocCur"
		
		sqlx = "update r3_obscommon..tdoc set CardName = N'" & saveHTMLDecode(Request("CardName"), False) & "', comments = " & Comments & ", cntctcode = N'" & request.form("CntctCode") & "', " & _
			   "NumAtCard = " & NumAtCard & ", SlpCode = " & Request("SlpCode") & ", GroupNum = " & Request("GroupNum") & ", DocDate = Convert(datetime,'" & SaveSqlDate(Request("DocDate")) & "',120), DocDueDate = " & DocDueDate & ", PartSupply = '" & PartSupply & "', ReserveInvoice = '" & Request("ReserveInvoice") & "', " & _
			   "DocCur = " & DocCur & ", PayToCode = " & PayToCode & ", ShipToCode = " & ShipToCode & ", Project = " & Project & " " & _
			   "where lognum = " & Session("retval") & " " & _
			   "update r3_obscommon..doc1 set SlpCode = " & Request("SlpCode") & " where lognum = " & Session("RetVal") & " " & _
			   "update R3_ObsCommon..TLOG set Object = " & ObjType & " where LogNum = " & Session("RetVal")
		conn.execute(sqlx)
		
		If Request("ChangePList") = "Y" Then

			Session("plist") = CInt(Request("NewPriceList"))
			set cmd = Server.CreateObject("ADODB.Command")
			cmd.ActiveConnection = connCommon
			cmd.CommandType = &H0004
			cmd.CommandText = "DBOLKApplyPListLines" & Session("ID")
			cmd.Parameters.Refresh()
			cmd("@LogNum") = Session("RetVal")
			cmd("@Pricelist") = Session("plist")
			cmd("@CardCode") = Session("UserName")
			cmd("@UserType") = userType
			cmd.execute()

		End If
		sql = "select AliasID, TypeID, (select SDKID collate database_default from r3_obscommon..tcif where companydb = '" & Session("OLKDb") & "')++AliasID As InsertID " & _
			  "from cufd T0 " & _
			  "left outer join OLKCUFD T1 on T1.TableID = T0.TableID and T1.FieldID = T0.FieldID " & _
			  "where T0.TableId = 'OINV' and AType in ('" & userType & "','T') and OP in ('T','P') and Active = 'Y'"
		set rs = server.createobject("ADODB.RecordSet")
		rs.open sql, conn, 3, 1
		If Not rs.eof then
			sql = "update r3_obscommon..tdoc set "
			do while not rs.eof
				If rs.bookmark <> 1 Then sql = sql & ", "
				If Request("U_" & rs("AliasID")) <> "" Then 
					strValue = Request("U_" & rs("AliasID"))
					Select Case rs("TypeID") 
						Case "B" 
							AliasVal = getNumeric(strValue)
						Case "D"
							AliasVal = "Convert(datetime,'" & SaveSqlDate(strValue) & "',120)"
						Case Else
							AliasVal = "N'" & saveHTMLDecode(strValue, False) & "'"
					End Select
				Else 
					AliasVal = "NULL"
				End If
				sql = sql & rs("InsertID") & " = " & AliasVal
			rs.movenext
			loop
			sql = sql & " where lognum = " & Session("RetVal")
			conn.execute(sql)
		End If
	Case "cartExp"
	
		AddItem = True
		TaxCode = ""
		Select Case myApp.LawsSet
			Case "MX", "CL", "CR", "GT", "US", "CA", "BR"
				If Request("TaxCode") <> "" Then
				  	TaxCode = Request("TaxCode")
				Else
				  	TaxCode = getExpTaxCode
				End If
				
				If TaxCode = "Disabled" Then 
					TaxCode = ", NULL"
				ElseIf TaxCode = "" Then
					errMsg = "&err=tax&expItem=Y&tItem=" & Request("Item") & "&document=" & Request("document") & "&page=" & Request("page")
					AddItem = False
				Else
					TaxCode = ", '" & TaxCode & "' "
				End If
		End Select

		sql = "delete R3_ObsCommon..DOC3 where LogNum = " & Session("RetVal") 
		If Request("chkExpns") <> "" Then sql = sql & " and ExpnsCode not in (" & Request("chkExpns") & ")"
		conn.execute(sql)
		
		If Request("chkExpns") <> "" Then
			sql = 	"declare @LogNum int set @LogNum = " & Session("RetVal") & " " & _
					"declare @LineNum int declare @ExpnsCode int "
			
			ArrExp = Split(Request("chkExpns"), ", ")	
			For i = 0 to UBound(ArrExp)
				sql = sql & "set @ExpnsCode = " & ArrExp(i) & " " & _
							"if not exists(select '' from R3_ObsCommon..DOC3 where LogNum = @LogNum and ExpnsCode = @ExpnsCode) begin " & _
							"	set @LineNum = IsNull((select Max(LineNum)+1 from R3_ObsCommon..DOC3 where LogNum = @LogNum),0) "
							
							If AddItem Then
								sql = sql & "EXEC OLKCommon..DBOLKCartAddExp" & Session("ID") & " " & Session("RetVal") & ", " & ArrExp(i) & TaxCode & ", " & getNumeric(Request("Price" & ArrExp(i))) & " "
							End If
							
				sql = sql & "end else begin " & _
							"	update R3_ObsCommon..DOC3 set LineTotal = " & getNumeric(Request("Price" & ArrExp(i))) & " " & _
							"	where LogNum = @LogNum and ExpnsCode = @ExpnsCode " & _
							"End "
			Next
		End If
		
		conn.execute(sql)
End Select

conn.close
response.redirect "../operaciones.asp?cmd=cart"

Function getExpTaxCode
	sql = "select IsNull(TaxCode,'') TaxCode " & _
	"from R3_ObsCommon..TDOC T0 " & _
	"left outer join CRD1 T1 on T1.CardCode = T0.CardCode collate database_default and T1.AdresType = 'S' and T1.Address = T0.ShipToCode collate database_default " & _
	"where T0.LogNum = " & Session("RetVal") & " "
	set rTax = Server.CreateObject("ADODB.RecordSet")
	set rTax = conn.execute(sql)
	If Not rTax.EOF Then
		getExpTaxCode = rTax(0)
	Else
		getExpTaxCode = "Disabled"
	End If
End Function

%>
