<% addLngPathStr = "" %><!--#include file="lang/flowViewControl.asp" -->
<% Class FlowViewControl
Sub GenerateFlowView %>
<div id="flowViewDialog" title="<%=getflowViewControlLngStr("LtxtFlowViewControl")%>" style="height: 380px; min-height: 380px;">
<div style="height: 368px; overflow: auto;" id="flowViewDialogContent"></div>
</div>
<script language="javascript">
var txtFlowViewControl = '<%=getflowViewControlLngStr("LtxtFlowViewControl")%>';
var txtCanceled = 'Canceled';
var txtRejected = 'Rejected';
var txtReOpened = 'Reopened';
var txtProcessed = 'Processed';
var txtWaiting = 'Waiting';
var txtReqNum = 'Request #';
var txtReqDate = 'Request Date';
var txtReqUser = 'Request User';
var txtRespUser = 'Response User';
var txtRespNote = 'Response Note';
var txtRespDate = 'Response Date';
var txtFlow = 'Flow Name';
var txtFlowNote = 'Flow Message';
var txtUserReq = 'User Request';
var txtDetails = 'Details';
</script>
<script type="text/javascript" src="flowViewControl.js"></script>
<% End Sub
End Class %>