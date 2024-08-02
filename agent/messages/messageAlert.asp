<% addLngPathStr = "messages/" %>
<!--#include file="lang/messageAlert.asp" -->
<div id="divAlert" style="display: none; position: absolute; z-index: 1; background-color: #FFFFFF; width: 110px;">
<table cellpadding="0" cellspacing="0" style="width: 100%; height: 20px;">
	<tr style="cursor: hand;" onclick="openMsgAlert();">
		<td class="MsgAlertTitle" style="padding-top:1px;padding-left:1px;padding-right:1px;">
		<img alt="" src="images/arrow_up_white.gif" id="imgAlertOpen">
		<span id="txtMsgAlert"></span></td>
	</tr>
</table>
<div id="divAlertMsg">
<ilayer name="scrollAlert1" width=100% height=100 clip="0,0,600,100">
<layer name="scrollAlert2" width=100% height=100 bgColor="white">
<div id="scrollAlert3" style="width:100%;height:100px;overflow:auto">
<table style="width: 100%" cellpadding="0" cellspacing="2" id="tblAlert">
</table>
</div>
</layer>
</ilayer>
</div>
</div>
<script language="javascript">
var txtNewMsg = '<%=getmessageAlertLngStr("LtxtNewMsg")%>';
var txtNewMsgs = '<%=getmessageAlertLngStr("LtxtNewMsgs")%>';
var txtGo2Inbox = '<%=getmessageAlertLngStr("LtxtGo2Inbox")%>';
var DtxtAlert = '<%=getmessageAlertLngStr("DtxtAlert")%>';
var txtClient = '<%=txtClient%>';
var DtxtSupplier = '<%=getmessageAlertLngStr("DtxtSupplier")%>';
var DtxtLead = '<%=getmessageAlertLngStr("DtxtLead")%>';
var txtAgent = '<%=txtAgent%>';
var DtxtSystem = '<%=getmessageAlertLngStr("DtxtSystem")%>';
var DtxtError = '<%=getmessageAlertLngStr("DtxtError")%>';
</script>
<script langauge="javascript" src="messages/messageAlert.js">
</script>