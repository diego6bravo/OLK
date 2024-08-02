<!--#include file="../chkLogin.asp" -->
<!--#include file="../clearItem.asp"-->
<!--#include file="../clsApplication.asp"-->
<!--#include file="../clsSession.asp"-->
<%

set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus
%>

<!--#include file="lang/setCart3.asp" -->
<html <% If Session("myLng") = "he" Then %>dir="rtl"<% End If %>>
<!--#include file="../myHTMLEncode.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=getsetCart3LngStr("Lttl3dxLines")%></title>
<link rel="stylesheet" type="text/css" href="../design/0/style/stylePopUp.css">

<%
set rs = Server.CreateObject("ADODB.recordset")

If Request("btnSubmit") <> "" Then
	sql = "select T0.BatchNum from OIBT T0 " & _
	"where T0.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "' and T0.WhsCode = N'" & saveHTMLDecode(Request("WhsCode"), False) & "' " & _
	"and T0.Status = 0 and T0.Quantity-IsNull(T0.IsCommited,0)-OLKCommon.dbo.OLKItmBtchCmmtd(" & Session("RetVal") & ", " & Request("LineNum") & ", T0.BatchNum) > 0 "
	set rs = conn.execute(sql)

	delBatch = ""
	
	sql = "declare @LogNum int set @LogNum = " & Session("RetVal") & " " & _
		"declare @LineNum int set @LineNum = " & Request("LineNum") & " " & _
		"declare @BatchNum nvarchar(32) "
		
	do while not rs.eof
		BatchNum = Replace(rs("BatchNum"), " ", "<:-Space->")
		SelQty = Request("Sel" & BatchNum)
		If SelQty > 0 Then
			sql = sql & "set @BatchNum = N'" & rs("BatchNum") & "' " & _
			"if not exists(select 'A' from R3_ObsCommon..DOC2 where LogNum = @LogNum and LineNum = @LineNum and BatchNum = @BatchNum) begin " & _
			"	insert R3_ObsCommon..DOC2(LogNum, LineNum, BatchNum, Quantity) " & _
			"	values(@LogNum, @LineNum, @BatchNum, " & SelQty & ") " & _
			"end else begin " & _
			"	update R3_ObsCommon..DOC2 set Quantity = " & SelQty & " where LogNum = @LogNum and LineNum = @LineNum and BatchNum = @BatchNum " & _
			"end "
		Else
			If delBatch <> "" Then delBatch = delBatch & ", "
			delBatch = delBatch & "N'" & rs("BatchNum") & "'"
		End If
	rs.movenext
	loop
	If delBatch <> "" Then
		sql = sql & "delete R3_ObsCommon..DOC2 where LogNum = @LogNum and LineNum = @LineNum and BatchNum in (" & delBatch & ") "
	End If
	conn.execute(sql)
	doClose
Else

sql = "select IsNull(T0.ItemName, '') ItemName, IsNull(T0.SalPackMsr, '') SalPackMsr, T0.SalPackUn, IsNull(T0.SalUnitMsr, '') SalUnitMsr, T0.NumInSale, T1.WhsCode, T0.PicturName, " & _
	"IsNull((select WhsName from OWHS where WhsCode = T1.WhsCode collate database_default), '') WhsName, T2.SaleType, T1.Quantity/Case T2.SaleType When 3 Then T0.SalPackUn Else 1 End Quantity, " & _
	"IsNull((select Sum(Quantity) from R3_ObsCommon..doc2 where LogNum = " & Session("RetVal") & " and LineNum = " & Request("LineNum") & "),0) SelQty, " & _
	"T0.DocEntry ItmEntry, " & _
	"(select U_GrpID from [@TM3DXITDE] where U_ItmEntry = T0.DocEntry) GrpID, " & _
	"(select U_GrpNam from [@TM3DXLVLT] where U_GrpID = (select U_GrpID from [@TM3DXITDE] where U_ItmEntry = T0.DocEntry)) GrpName, " & _
	"(select U_grpLvlT from [@TM3DXLVLT] where U_GrpID = (select U_GrpID from [@TM3DXITDE] where U_ItmEntry = T0.DocEntry)) GrpLevel " & _
	"from OITM T0 " & _
	"inner join R3_ObsCommon..DOC1 T1 on T1.LogNum = " & Session("RetVal") & " and T1.LineNum = " & Request("LineNum") & " " & _
	"inner join OLKSalesLines T2 on T2.LogNum = T1.LogNum and T2.LineNum = T1.LineNum " & _
	"where T0.ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "'"
set rs = conn.execute(sql)

If rs("PicturName") <> "" Then
	Pic = rs("PicturName")
Else
	Pic = "n_a.gif"
End If 

lineQty = CDbl(rs("Quantity"))
lineUnit = rs("SaleType")
ItmEntry = rs("ItmEntry")
GrpID = rs("GrpID")
GrpLevel = rs("GrpLevel")
WhsCode = rs("WhsCode")
SalPackUn = CDbl(rs("SalPackUn"))
NumInSale = CDbl(rs("NumInSale"))

Select Case CInt(lineUnit)
	Case 1
		ReqQty = lineQty
	Case 2
		ReqQty = lineQty*NumInSale
	Case 3
		ReqQty = lineQty*NumInSale*SalPackUn
End Select

If Request("LvlTyp1") <> "" Then
	set cm = Server.CreateObject("ADODB.Command")
	cm.CommandText = "DBOLKSetItm3dxView" & Session("ID")
	cm.CommandType = adCmdStoredProc
	cm.ActiveConnection = connCommon
	cm.Parameters.Refresh
	set p = cm.Parameters
	cm("@SlpCode") = Session("vendid")
	cm("@ItemCode") = Request("Item")
	
	cm("@LvlTypID") = 1
	cm("@SelectedValue") = Request("LvlTyp1")
	If Request("ShowDesc1") = "Y" Then cm("@ShowDesc") = "Y" Else cm("@ShowDesc") = "N"
	cm.Execute
	
	If GrpLevel >= 2 Then
		cm("@LvlTypID") = 2
		cm("@SelectedValue") = Request("LvlTyp2")
		If Request("ShowDesc2") = "Y" Then cm("@ShowDesc") = "Y" Else cm("@ShowDesc") = "N"
		cm.Execute
	End If
End If

Dim selLvlIndex(1)
Dim selLvl(5)
Dim selLvlChk(5)
Dim showDesc(5)
set rl = Server.CreateObject("ADODB.RecordSet")
sql = "select LvlTypID, SelectedValue, ShowDesc from OLKItm3dxView where SlpCode = " & Session("vendid") & " and ItemCode = N'" & saveHTMLDecode(Request("Item"), False) & "'"
rl.open sql, conn, 3, 1
do while not rl.eof
	i = rl("LvlTypID")-1
	selLvlIndex(rl.bookmark-1) = i
	selLvl(i) = rl("SelectedValue")
	showDesc(i) = rl("ShowDesc")
rl.movenext
loop

sql = "select T0.U_LvlTypID, T0.U_LvlTypNa from [@TM3DXLVLX] T0 where T0.U_GrpID = " & GrpID
rl.close
rl.open sql, conn, 3, 1
%>
</head>
<script language="javascript" src="../general.js"></script>
<script language="javascript">
var txtValSelDimX = '<%=getsetCart3LngStr("LtxtValSelDimX")%>';
var txtValRepDim = '<%=getsetCart3LngStr("LtxtValRepDim")%>';
var txtValNumVal = '<%=getsetCart3LngStr("DtxtValNumVal")%>';
var txtValNumMinVal = '<%=getsetCart3LngStr("DtxtValNumMinVal")%>';
var txtValMoreThenAvl = '<%=getsetCart3LngStr("LtxtValMoreThenAvl")%>';
var txtValSelOverQty = '<%=getsetCart3LngStr("LtxtValSelOverQty")%>';
</script>
<script language="javascript" src="setCart3.js.asp?GrpLevel=<%=GrpLevel%>"></script>
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" <% If Request("btnShow") <> "" Then %>onload="SumTotal();"<% End If %> onfocus="javascript:if (!OpenWin.closed) OpenWin.focus();">
<table border="0" width="100%" id="table1" cellpadding="0" height="400">
	<form method="POST" action="setCart3.asp" name="frmMain" onsubmit="return valFrm();" webbot-action="--WEBBOT-SELF--">
		<tr>
			<td colspan="5" class="GeneralTblBold2"><%=getsetCart3LngStr("Lttl3dxSel")%> - <%=rs("GrpName")%></td>
		</tr>
		<tr>
			<td class="GeneralTblBold2" height="100" rowspan="5">
			<p align="center"><a href="javascript:Start('../thumb/?item=<%=Replace(Replace(Replace(myHTMLEncode(Request("Item")),"#","%23"),"&","%26"),"""","%22")%>&pop=Y&AddPath=../',529,510,'yes')"><img border="0" src="../pic.aspx?filename=<%=Pic%>&dbName=<%=Session("olkdb")%>"></a></td>
			<td class="GeneralTblBold2" height="15"><%=getsetCart3LngStr("DtxtItem")%>:</td>
			<td class="GeneralTbl" height="15"><%=Request("Item")%>&nbsp;</td>
			<td class="GeneralTblBold2" height="15"><%=getsetCart3LngStr("DtxtWarehouse")%>:</td>
			<td class="GeneralTbl" height="15"><%=rs("WhsName")%>&nbsp;</td>
		</tr>
		<tr>
			<td class="GeneralTblBold2"><%=getsetCart3LngStr("DtxtDescription")%>:</td>
			<td class="GeneralTbl" colspan="3"><%=rs("ItemName")%>&nbsp;</td>
		</tr>
		<tr>
			<td class="GeneralTblBold2"><%=getsetCart3LngStr("LtxtReqQty")%>:</td>
			<td class="GeneralTbl">
			<input class="GeneralTbl" type="text" name="txtReqQty" value="<%=ReqQty%>" size="1" style="text-align: right; width: 100%; border-style: solid; border-width: 0"></td>
			<td class="GeneralTblBold2">&nbsp;</td>
			<td class="GeneralTbl">&nbsp;</td>
		</tr>
		<tr>
			<td class="GeneralTblBold2"><%=getsetCart3LngStr("LtxtSelQty")%>:</td>
			<td class="GeneralTbl"><input class="GeneralTbl" type="text" name="txtSelQty" value="<%=rs("SelQty")%>" size="1" style="text-align: right; width: 100%; border-style: solid; border-width: 0"></td>
			<td class="GeneralTblBold2">&nbsp;</td>
			<td class="GeneralTbl">&nbsp;</td>
		</tr>
		<tr>
			<td class="GeneralTblBold2"><%=getsetCart3LngStr("LtxtPendQty")%>:</td>
			<td class="GeneralTbl"><input class="GeneralTbl" type="text" name="txtOpenQty" value="<%=ReqQty-CDbl(rs("SelQty"))%>" size="1" style="text-align: right; width: 100%; border-style: solid; border-width: 0"></td>
			<td class="GeneralTblBold2">&nbsp;</td>
			<td class="GeneralTbl">&nbsp;</td>
		</tr>
		<tr>
			<td colspan="5" valign="top" class="GeneralTblBold2">
				<table cellpadding="0" cellspacing="2" border="0" width="100%">
					<% If GrpLevel >= 2 Then %>
					<tr class="GeneralTblBold2">
						<td width="120">&nbsp;</td>
						<td>
						<table border="0" id="table4" cellpadding="0" bgcolor="#FFFFFF">
							<tr>
								<td class="GeneralTblBold2"><%=getsetCart3LngStr("LtxtDimension")%> 2</td>
								<td class="GeneralTbl">
								<select size="1" name="LvlTyp2">
								<option></option>
								<% do while not rl.eof %>
								<option <% If Request("LvlTyp2") = CStr(rl(0)) or Request.Form.Count = 0 and CStr(selLvl(1)) = CStr(rl(0)) Then %>selected<% End If %> value="<%=rl(0)%>"><%=myHTMLEncode(rl(1))%></option>
								<% rl.movenext
								loop
								If rl.RecordCount > 0 Then rl.movefirst %>
								</select></td>
								<td class="GeneralTbl">
								<input style="border-style:solid; border-width:0; background:background-image" type="checkbox" <% If Request("ShowDesc2") = "Y" or Request.Form.Count = 0 and showDesc(1) = "Y" Then %>checked<% End If %> name="ShowDesc2" id="ShowDesc2" value="Y"><label for="ShowDesc2"><%=getsetCart3LngStr("DtxtDescription")%></label></td>
							</tr>
						</table>
						</td>
					</tr>
					<% End If %>
					<tr>
						<td valign="top" width="120" class="GeneralTblBold2">
						<table border="0" width="100%" id="table3" cellpadding="0" bgcolor="#FFFFFF">
							<tr>
								<td class="GeneralTblBold2"><%=getsetCart3LngStr("LtxtDimension")%>&nbsp;1</td>
							</tr>
							<tr>
								<td class="GeneralTbl"><select size="1" name="LvlTyp1">
								<option></option>
								<% do while not rl.eof %>
								<option <% If Request("LvlTyp1") = CStr(rl(0)) or Request.Form.Count = 0 and CStr(selLvl(0)) = CStr(rl(0)) Then %>selected<% End If %> value="<%=rl(0)%>"><%=myHTMLEncode(rl(1))%></option>
								<% rl.movenext
								loop
								If rl.RecordCount > 0 Then rl.movefirst %>
								</select></td>
							</tr>
							<tr>
								<td class="GeneralTbl">
								<input style="border-style:solid; border-width:0; background:background-image" type="checkbox" name="ShowDesc1" <% If Request("ShowDesc1") = "Y" or Request.Form.Count = 0 and showDesc(0) = "Y" Then %>checked<% End If %> id="ShowDesc1" value="Y"><label for="ShowDesc1"><%=getsetCart3LngStr("DtxtDescription")%></label></td>
							</tr>
							<% 
						
							If Request("LvlTyp1") <> "" Then
								chkLvlTyp1 = CInt(Request("LvlTyp1"))
							Else
								chkLvlTyp1 = CInt(selLvl(0))
							End If
													
							If GrpLevel >= 2 Then
								If Request("LvlTyp2") <> "" Then
									chkLvlTyp2 = CInt(Request("LvlTyp2"))
								Else
									chkLvlTyp2 = CInt(selLvl(1))
								End If
							Else
								chkLvlTyp2 = chkLvlTyp1
							End If
						
							Dim rvCount(5)
							Dim rvItems()
							Dim rvItem(1)
							If selLvl(0) <> "" and GrpLevel > 2 Then
							set rv = Server.CreateObject("ADODB.RecordSet")

							do while not rl.eof 
							If CStr(rl(0)) <> CStr(chkLvlTyp1) and CStr(rl(0)) <> CStr(chkLvlTyp2) Then
							sql = "select U_ValSign ValSign, U_ValSign + ' - ' + U_ValDesc ValDesc " & _
							"from [@TM3DXLVLV] " & _
							"where U_GrpID = " & GrpID & " and U_LvlTypID = " & rl(0) & " order by U_ValOrd asc"
							rv.open sql, conn, 3, 1
							Redim rvItems(rv.RecordCount-1) %>
							<tr>
								<td class="GeneralTblBold2">
								<%=rl(1)%>&nbsp;</td>
							</tr>
							<tr>
								<td class="GeneralTbl">
								<select size="1" name="LvlSelVal<%=rl(0)%>" onchange="chOtherDim()">
								<% do while not rv.eof
								rvItem(0) = rl(0)
								rvItem(1) = rv(0)
								rvItems(rv.bookmark-1) = rvItem
								If rv.bookmark = 1 Then
									selLvlChk(CInt(rl(0))) = rv(0)
								End If %>
								<option value="<%=rv(0)%>"><%=myHTMLEncode(rv(1))%></option>
								<% rv.movenext
								loop
								rvCount(CInt(rl(0))) = rvItems %>
								</select></td>
							</tr>
							<% 
							rv.close
							End If
							rl.movenext
							loop
							rl.movefirst
							set rv = nothing
							End If %>
							</table>
						</td>
						<td valign="top" rowspan="2">
						<% If selLvl(0) <> "" Then
						sql = "select U_LvlTypID LvlTypID, U_ValSign ValSign, U_ValDesc ValDesc, U_ValOrd ValOrd from [@TM3DXLVLV] where U_GrpID = " & GrpID & " "
						If GrpLevel = 1 Then
							sql = sql & " union " & _
										"select 1 LvlTypID, null ValSign, null ValDesc, 0 ValOrd"
						End If
						sql = sql & " order by LvlTypID, ValOrd"
						rs.close
						rs.open sql, conn, 3, 1
		
						set rb = Server.CreateObject("ADODB.RecordSet")
						sql = "SELECT X0.U_ValSign + IsNull('-'+X1.U_ValSign,'') + IsNull('-'+X2.U_ValSign,'') + IsNull('-'+X3.U_ValSign,'') + IsNull('-'+X4.U_ValSign,'') + IsNull('-'+X5.U_ValSign,'') N'Sign', " & _
						"X0.U_ValDesc + IsNull('-'+X1.U_ValDesc,'') + IsNull('-'+X2.U_ValDesc,'') + IsNull('-'+X3.U_ValDesc,'') + IsNull('-'+X4.U_ValDesc,'') + IsNull('-'+X5.U_ValDesc,'') N'Descripción',  " & _
						"IsNull(T0.Quantity,0)-IsNull(T0.IsCommited,0)-OLKCommon.dbo.OLKItmBtchCmmtd(" & Session("RetVal") & ", " & Request("LineNum") & ", T0.BatchNum) AvlQty,  " & _
						"IsNull(T1.Quantity,0) SelQty  " & _
						"from [@TM3DXLVLV] X0 " & _
						"left outer join [@TM3DXLVLV] X1 on X1.U_GrpID = X0.U_GrpID and X1.U_LvlTypID = 1 " & _
						"left outer join [@TM3DXLVLV] X2 on X2.U_GrpID = X0.U_GrpID and X2.U_LvlTypID = 2 " & _
						"left outer join [@TM3DXLVLV] X3 on X3.U_GrpID = X0.U_GrpID and X3.U_LvlTypID = 3 " & _
						"left outer join [@TM3DXLVLV] X4 on X4.U_GrpID = X0.U_GrpID and X4.U_LvlTypID = 4 " & _
						"left outer join [@TM3DXLVLV] X5 on X5.U_GrpID = X0.U_GrpID and X5.U_LvlTypID = 5 " & _
						"left outer join OIBT T0 on T0.ItemCode = N'" & Request("Item") & "' and T0.WhsCode = N'" & saveHTMLDecode(Request("WhsCode"), False) & "' and T0.BatchNum = X0.U_ValSign + IsNull('-'+X1.U_ValSign,'') + IsNull('-'+X2.U_ValSign,'') + IsNull('-'+X3.U_ValSign,'') + IsNull('-'+X4.U_ValSign,'') + IsNull('-'+X5.U_ValSign,'') " & _
						"left outer join R3_ObsCommon..DOC2 T1 on T1.LogNum = " & Session("RetVal") & " and T1.LineNum = " & Request("LineNum") & "  " & _
						"	and T1.BatchNum = X0.U_ValSign + IsNull('-'+X1.U_ValSign,'') + IsNull('-'+X2.U_ValSign,'') + IsNull('-'+X3.U_ValSign,'') + IsNull('-'+X4.U_ValSign,'') + IsNull('-'+X5.U_ValSign,'') collate database_default " & _
						"where X0.U_GrpID = (select U_GrpID from [@TM3DXITDE] where U_ItmEntry = (select DocEntry from OITM where ItemCode = N'" & Request("Item") & "')) and X0.U_LvlTypID = 0 order by "
						
						sql = sql & "X" & chkLvlTyp1 & ".U_LvlTypID, X" & chkLvlTyp1 & ".U_ValOrd "
						
						If GrpLevel >= 2 Then
							sql = sql & ", X" & chkLvlTyp2 & ".U_LvlTypID, X" & chkLvlTyp2 & ".U_ValOrd "
							For o = 0 to GrpLevel-1
								If o <> chkLvlTyp1 and o <> chkLvlTyp2 Then
									sql = sql & ", X" & o & ".U_LvlTypID, X" & o & ".U_ValOrd "
								End If
							Next
						End If
						rb.open sql, conn, 3, 1
						 %>
						<ilayer name="scroll1" width=460 height=210 clip="0,0,170,150">
						<layer name="scroll2" width=460 height=210 bgColor="white">
						<div id="scroll3" style="width:460px;height:210px;background-color:white;overflow:scroll">
							<table border="0" cellspacing="2" cellpadding="0" width="100%" id="tblMatrix">
								<tr>
									<td></td>
									<% 
									If Request("LvlTyp2") <> "" Then
										rs.Filter = "LvlTypID = " & Request("LvlTyp2")
									Else
										If GrpLevel >= 2 Then
											rs.Filter = "LvlTypID = " & selLvl(1)
										Else
											rs.Filter = "LvlTypID = 1"
										End If
									End If
									Dim colCount()
									Redim colCount(rs.RecordCount-1) %>
									<script language="javascript">
									var colCount = new Array(<%=UBound(colCount)+1%>);
									</script>
									<% For i = 1 to rs.RecordCount
									colCount(i-1) = rs("ValSign") %>
									<script language="javascript">colCount[<%=i-1%>] = '<%=rs("ValSign")%>'</script>
									<td class="GeneralTblBold2" colspan="2" id="col<%=rs("valSign")%>">
									<p align="center"><%
									If Request("ShowDesc2") = "Y" or Request.Form.Count = 0 and showDesc(1) = "Y" Then
										Response.Write rs("ValDesc")
									Else
										Response.Write rs("ValSign")
									End If %>&nbsp;</td>
									<% rs.movenext
									Next
									rs.movefirst %>
								</tr>
								<tr>
									<td></td>
									<% For i = 0 to UBound(colCount) %>
									<td class="GeneralTblBold2" id="col<%=colCount(i)%>">
									<%=getsetCart3LngStr("LtxtAvl")%></td>
									<td class="GeneralTblBold2" id="col<%=colCount(i)%>">
									<%=getsetCart3LngStr("LtxtSelection")%></td>
									<% Next %>
								</tr>
								<% 
								rs.Filter = "LvlTypID = " & chkLvlTyp1
								
								last = ""
								do while not rb.eof
								arrSign = Split(rb("Sign"), "-")
								If arrSign(chkLvlTyp1) <> last Then
									If Request("ShowDesc1") = "Y" or Request.Form.Count = 0 and showDesc(0) = "Y" Then
										rowDesc = rs("ValDesc")
									Else
										rowDesc = rs("ValSign")
									End If
									If rb.bookmark <> 1 Then Response.Write "</tr>"
									Response.Write "<tr id = ""row" & rs("valSign") & """><td class=""GeneralTblBold2"">" & rowDesc & "</td>" 
									last = arrSign(chkLvlTyp1)
									rs.movenext
								End If
								
								AvlQty = CDbl(rb("AvlQty"))
								If Request.Form.Count = 0 or Request("Sel" & rb("Sign")) = "" Then
									SelQty = CDbl(rb("SelQty"))
								Else
									SelQty = CDbl(Request("Sel" & rb("Sign")))
								End If
								
								hide = False
								If GrpLevel > 2 Then
									For i = 0 to GrpLevel - 1
										If i <> chkLvlTyp1 and i <> chkLvlTyp2 Then
											If Split(rb("Sign"), "-")(i) <> selLvlChk(i) Then
												hide = true
											End If
										End If
									Next
								End If
								%>
								<td class="GeneralTbl" align="right" id="<%=myHTMLEncode(rb("Sign"))%>" <% If hide Then %>style="display: none;"<% End If %>><%=AvlQty%></td>
								<td class="GeneralTbl" id="<%=myHTMLEncode(rb("Sign"))%>" <% If hide Then %>style="display: none;"<% End If %>>
								<input type="text" <% If AvlQty = 0 Then %>readonly<% End If %> id="SelQty" name="Sel<%=myHTMLEncode(rb("Sign"))%>" size="6" value="<%=SelQty%>" style="text-align: right<% If AvlQty = 0 Then %>; background-color:#E8E8E8<% End If %>; width:100%" onfocus="this.select();" onchange="chkQty(this, <%=SelQty%>, <%=AvlQty%>)">
								</td>
								<% rb.movenext
								loop
								Response.Write "</tr>"
								%>
							</table> 
						</div>
						</layer>
						</ilayer>
						<% End If %>
						</td>
					</tr>
					<tr>
					<% 
					If GrpLevel >= 2 Then 
						tdShowH = 138 - ((GrpLevel-2)*40)
					Else 
						rdShowH = 164
					End If %>
						<td valign="bottom" width="120" class="GeneralTblBold2" height="<%=tdShowH%>">
						<input type="submit" value="<%=getsetCart3LngStr("LtxtShow")%>" name="btnShow" onclick="return valShow();"></td>
					</tr>
				</table>
			</td>
		</tr>
	<tr>
		<td colspan="5" class="GeneralTblBold2">
		<table border="0" cellspacing="0" width="100%" id="table3">
			<tr>
				<td>
				<input type="submit" value="<%=getsetCart3LngStr("DtxtSave")%>" name="btnSubmit" <% If Request("LvlTyp1") = "" and Request.Form.Count = 0 and selLvl(0) = "" Then %>onclick="alert('<%=getsetCart3LngStr("LtxtValSelDim")%>');return false;"<% End If %>></td>
				<td>
				<p align="right">
				<input type="button" value="<%=getsetCart3LngStr("DtxtCancel")%>" name="btnCancel" onclick="javascript:if(confirm('<%=getsetCart3LngStr("LtxtValCloseWin")%>'))window.close();"></td>
			</tr>
		</table>
		</td>
	</tr>
	<input type="hidden" name="LineNum" value="<%=Request("LineNum")%>">
	<input type="hidden" name="Quantity" value="<%=lineQty%>">
	<input type="hidden" name="SelUn" value="<%=lineUnit%>">
	<input type="hidden" name="WhsCode" value="<%=myHTMLEncode(WhsCode)%>">
	<input type="hidden" name="Item" value="<%=myHTMLEncode(Request("Item"))%>">
	<input type="hidden" name="order" value="<%=order%>">
	<input type="hidden" name="orderBy" value="<%=orderBy%>">
	</form>
</table>
<% If selLvl(0) <> "" Then %>
<script language="javascript">chOtherDim();</script>
<% End If %>
</body>
<% rl.close
set rl = nothing
set rb = nothing
End If %>
<% conn.close
set rs = nothing %>
</html>
<% Sub doClose %>
<script>
<% 
If Request("txtSelQty") = 0 Then
	rCmd = 0
ElseIf Request("txtSelQty") = Request("txtReqQty") Then
	rCmd = 1
Else 
	rCmd = 2
End If
%>
opener.setS3Img(<%=rCmd%>);
window.close();
</script>
<% End Sub %>