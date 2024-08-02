<% addLngPathStr = "" %><!--#include file="lang/flowControl.asp" -->
<% 
Class FlowControl

Sub GenerateFlow %>
<div id="flowDialog" title="" style="height: 280px; min-height: 280px;">&nbsp;
<div style="height: 268px; min-height: 268px; overflow: auto;" id="flowDialogContent"></div>
</div>
<script type="text/javascript">
var SelDes = '<%=SelDes%>';

var altError = '<%=getflowControlLngStr("LaltError")%>';
var altConf = '<%=getflowControlLngStr("LaltConf")%>';
var altFlow = '<%=getflowControlLngStr("LaltFlow")%>';

var txtClose = '<%=getflowControlLngStr("DtxtClose")%>';
var txtYes = '<%=getflowControlLngStr("DtxtYes")%>';
var txtNo = '<%=getflowControlLngStr("DtxtNo")%>';
var txtConfirm = '<%=getflowControlLngStr("DtxtConfirm")%>';
var txtCancel = '<%=getflowControlLngStr("DtxtCancel")%>';
var txtNote = '<%=getflowControlLngStr("DtxtNote")%>';
var txtConfFlow = '<%=getflowControlLngStr("LtxtConfFlow")%>';
var txtContinue = '<%=getflowControlLngStr("DtxtContinue")%>';

</script>
<script type="text/javascript" src="flowControl.js"></script>
<% End Sub

End Class 

%>