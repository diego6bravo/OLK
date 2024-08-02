<%
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = &H0004
cmd.CommandText = "OLKCreatePDFAccess"
cmd.Parameters.Refresh
cmd("@dbID") = Session("ID")
cmd.execute

myRnd = cmd("@rnd").value
%>