<% addLngPathStr = "inv/" %>
<!--#include file="lang/delOrderTitle.asp" -->
<% Select Case Session("Type") %>
<% Case "I" %>
<%=getdelOrderTitleLngStr("LtxtPurchaseOrderChec")%>
<% Case "O" %>
<%=getdelOrderTitleLngStr("LtxtSalesOrderCheck")%>
<% End Select %>