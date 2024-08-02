<!--#include file="chkLogin.asp"-->
<!--#include file="myHTMLEncode.asp"-->
<!--#include file="clsApplication.asp"-->
<!--#include file="clsSession.asp"-->
<%
set myApp = New clsApplication
myApp.CheckApplicationStatus

set mySession = New clsSession
mySession.CheckSessionStatus

			
sql = 	"select T0.varVar, T0.varDataType, T0.varNotNull " & _
		"from OLKRSVars T0 " & _
		"where T0.rsIndex = " & Request("rsIndex")
set rs = conn.execute(sql) %>
<script language="javascript">
var rsVars = new Array();
<% do while not rs.eof %>
addVar(new ReportVariable('<%=rs("varVar")%>', '<%=rs("varDataType")%>', '<%=rs("varNotNull")%>'));
<% rs.movenext
loop %>
parent.setRSVars(rsVars);

function ReportVariable(varVar, varDataType, varNotNull)
{
	this.varVar = varVar;
	this.varDataType = varDataType;
	this.varNotNull = varNotNull;
}

function addVar(newVariable) {
	this.rsVars[this.rsVars.length] = newVariable;
}
</script>
<% conn.close 
set rs = nothing %>