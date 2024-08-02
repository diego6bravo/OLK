<% addLngPathStr = "" %>
<!--#include file="lang/submitControl.asp" -->

<% Class SubmitControl

Dim enableBG

Dim bgRedir

Dim submitNum

Dim submitNum2

Dim submitNumID

Dim submitNumID2

Dim enableRestore

Dim goEndDesc

Dim goEndFunc

Dim goSecondDesc

Dim goSecondFunc

Dim goThirdDesc

Dim goThirdFunc

Dim transOkMsg

Public Property Let EndButtonDescription(p_Value)
	goEndDesc = p_Value
End Property

Public Property Let EndButtonFunction(p_Value)
	goEndFunc = p_Value
End Property

Public Property Let SecondButtonDescription(p_Value)
	goSecondDesc = p_Value
End Property

Public Property Let SecondButtonFunction(p_Value)
	goSecondFunc = p_Value
End Property

Public Property Let ThirdButtonDescription(p_Value)
	goThirdDesc = p_Value
End Property

Public Property Let ThirdButtonFunction(p_Value)
	goThirdFunc = p_Value
End Property

Public Property Let EnableRunInBackground(p_Value)
	enableBG = p_Value
End Property

Public Property Let RunInBackgroundRedir(p_Value)
	bgRedir = p_Value
End Property

Public Property Let LogNum(p_Value)
	submitNum = p_Value
End Property

Public Property Let LogNum2(p_Value)
	submitNum2 = p_Value
End Property

Public Property Let LogNumID(p_Value)
	submitNumID = p_Value
End Property

Public Property Let LogNumID2(p_Value)
	submitNumID2 = p_Value
End Property

Public Property Let Restore(p_Value)
	enableRestore = p_Value
End Property

Public Property Let TransactionOkMessage(p_Value)
	transOkMsg = p_Value
End Property

Sub GenerateSubmit

If submitNum = "" Then
	Response.Write "LogNum Field is Mandatory"
	Exit Sub
End If

set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "DBOLKGetSubmitControl" & Session("ID")
cmd.Parameters.Refresh()
cmd("@LogNum") = submitNum
set rs = cmd.execute()
poolCount = rs(0)
objCode = rs(1)
hasPrint = rs(2) = "Y"

If userType = "V" Then 
	poolCSS = "CanastaTblExpense"
Else 
	poolCSS = "CanastaTblResaltada"
End If

If submitNum2 = "" Then submitNum2 = "null"

poolCount = "<span class=""" & poolCSS & """>" & poolCount & "</span>" %>
<div style="height: 200px;">&nbsp;</div>
<table border="0" cellpadding="0" width="300" align="center">
	<tr>
		<td>
		<p align="center">
		<img border="0" src="design/<%=SelDes%>/images/gear.gif" id="submitImgWait">
		<img border="0" src="design/<%=SelDes%>/images/error_icon.gif" width="234" height="211" id="submitImgError" style="display: none; ">
		<img border="0" src="design/<%=SelDes%>/images/confirmIcon.gif" width="68" height="65" id="submitImgOK" style="display: none; "></td>
	</tr>
	<tr class="FirmTlt">
		<td>
		<p align="center"><span id="submitTitle"><%=getsubmitControlLngStr("DtxtWait")%>...</span></td>
	</tr>
	<tr class="FirmTbl" id="trSubmitTransPool">
		<td>
		<p align="center"><span id="submitStatus"><%=Replace(getsubmitControlLngStr("LtxtTransactinoPool"), "{0}", poolCount)%></span></td>
	</tr>
	<% If enableBG Then %>
	<tr class="GeneralTblBold2" id="trSubmitRunInBG">
		<td>
		<p align="center"><input type="button" name="runInBg" value="<%=getsubmitControlLngStr("DtxtRunInBG")%>" onclick="javascript:runInBG();"></p>
		</td>
	</tr>
	<% End If %>
	<tr class="FirmTbl" id="trSubmitErr" style="display: none; ">
		<td>
		<p align="center"><% If enableRestore Then %>
			<input type="button" name="btnRestore" value="<%=getsubmitControlLngStr("DtxtRestore")%>" onclick="javascript:restore();"> 
			-<% End If %>
		<input type="button" name="btnRetry" value="<%=getsubmitControlLngStr("DtxtRetry")%>" onclick="javascript:retryAction();"></td>
	</tr>
	<tr class="FirmTbl" id="trSubmitOk" style="display: none; ">
		<td>
		<p align="center"><input type="button" name="btnOk" value="<%=goEndDesc%>" onclick="javascript:<%=goEndFunc%>"></td>
	</tr>
	<% If goSecondDesc <> "" Then %>
	<tr class="FirmTbl" id="trSubmitOk2" style="display: none; ">
		<td>
		<p align="center"><input type="button" name="btnOk2" id="btnOk2" value="<%=goSecondDesc%>"></td>
	</tr>
	<% End If %>
	<% If goThirdDesc <> "" Then %>
	<tr class="FirmTbl" id="trSubmitOk3" style="display: none; ">
		<td>
		<p align="center"><input type="button" name="btnOk3" id="btnOk3" value="<%=goThirdDesc%>"></td>
	</tr>
	<% End If %>
	<% If hasPrint Then
	strPrint = ""
	set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = connCommon
	cmd.CommandType = &H0004
	cmd.CommandText = "DBOLKGetObjectPrint" & Session("ID")
	cmd.Parameters.Refresh()
	cmd("@ObjCode") = objCode
	cmd("@UserType") = userType
	cmd("@LanID") = Session("LanID")
	set rp = Server.CreateObject("ADODB.RecordSet")
	set rp = cmd.execute()
	do while not rp.eof
	secID = rp("SecID")
	If strPrint <> "" Then strPrint = strPrint & ", "
	strPrint = strPrint & secID %>
	<tr class="FirmTbl" id="trSubmitPrint<%=secID%>" style="display: none; ">
		<td>
		<p align="center"><input type="button" name="btnPrint<%=secID%>" id="btnPrint<%=secID%>" value="<%=rp("SecName")%>" onclick="javascript:openPrint(<%=secID%>, '<%=rp("LinkData")%>');"></td>
	</tr>
	<% rp.movenext
	loop
	set rp = nothing  
	End If %>
</table>
<script type="text/javascript">
var strPrint = '<%=strPrint%>';
var hasPrint = <%=JBool(hasPrint)%>;
var LogNum = <%=submitNum%>;
var LogNumID = '<%=submitNumID%>';
var LogNum2 = <%=submitNum2%>;
var LogNumID2 = '<%=submitNumID2%>';
var goEndDesc = '<%=Replace(goEndDesc, "'", "\'")%>';
var goEndFunc = '<%=Replace(goEndFunc, "'", "\'")%>';
var goSecondDesc = '<%=Replace(goSecondDesc, "'", "\'")%>';
var goSecondFunc = '<%=Replace(goSecondFunc, "'", "\'")%>';
var goThirdDesc = '<%=Replace(goThirdDesc, "'", "\'")%>';
var goThirdFunc = '<%=Replace(goThirdFunc, "'", "\'")%>';
var enableBG = <%=JBool(enableBG)%>;
var bgRedir = '<%=Replace(bgRedir, "'", "\'")%>';
var transOkMsg = '<%=Replace(transOkMsg, "'", "\'")%>';
var txtTransactinoPool = '<%=getsubmitControlLngStr("LtxtTransactinoPool")%>';
var poolCSS = '<%=poolCSS%>';
var txtTransComp = '<%=getsubmitControlLngStr("LtxtTransComp")%>';
var txtError = '<%=getsubmitControlLngStr("DtxtError")%>';
var txtProcessing = '<%=getsubmitControlLngStr("LtxtProcessing")%>';
var txtWait = '<%=getsubmitControlLngStr("DtxtWait")%>';
</script>
<script type="text/javascript" src="submitControl.js"></script>
<% End Sub

End Class %>