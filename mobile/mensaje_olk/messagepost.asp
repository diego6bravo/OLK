<% addLngPathStr = "mensaje_olk/" %>
<!--#include file="lang/messagepost.asp" -->
<%
If Request.Form("Urgent") = "Y" Then Urgent = "Y" Else Urgent = "N"
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "DBOLKIntMsg" & Session("ID")
cmd.Parameters.Refresh()
cmd("@OlkUFrom") = Session("vendid")
cmd("@OlkUFromType") = "V"
cmd("@OlkMSG") = Request("Message")
cmd("@OlkSubject") = Request("Subject")
cmd("@OlkUrgent") = Urgent
cmd.execute()
olklog = cmd("@OlkLog")

sql = "SELECT T0.SlpCode AS User_Code, T0.SlpName AS U_Name " & _
	  "FROM OSLP T0 " & _
	  "inner join OLKAgentsAccess T1 on T1.SlpCode = T0.SlpCode and T1.Access <> 'D' " & _
	  "WHERE T0.SlpCode not in (-1," & Session("vendid") & ") "
set rs = conn.execute(sql)
set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = connCommon
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "DBOlkIntMSGU" & Session("ID")
cmd.Parameters.Refresh()
cmd("@OlkLog") = olklog
cmd("@OlkStatus") = "N"
cmd("@UserType") = "V"
cmd("@LanID") = Session("LanID")

do while not rs.eof
	varx = "U" & CStr(RS("User_Code"))
	If Request.Form(varx) = "ON" Then
		cmd("@User") = rs("U_Name")
		cmd.execute()
	end if
rs.movenext
loop
%>
<center>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111" id="table1">
    <!-- fwtable fwsrc="Z:\topmanage\logos\originales\pocket_art.png" fwbase="pocket_artpieza1.gif" fwstyle="FrontPage" fwdocid = "742308039" fwnested=""0" -->
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="table2">
        <tr>
          <td width="100%">
          <p align="<% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>"><b><font face="Verdana" size="1">&nbsp;</font></b></td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center"><b><font size="1" face="Verdana"><%=getmessagepostLngStr("LtxtMsgOK")%></font></b></td>
        </tr>
        <tr>
          <td width="100%" style="font-size: 10px">&nbsp;</td>
        </tr>
        <tr>
          <td width="100%"><font SIZE="1" face="Verdana">
          <p align="center"><b>[<a href="operaciones.asp?cmd=home"><font color="#000000"><%=getmessagepostLngStr("LtxtRetOP")%></font></a>]</b></font></td>
        </tr>
        <tr>
          <td width="100%">
          <p align="center"><font SIZE="1" face="Verdana"><b>[<a href="operaciones.asp?cmd=mssolk"><font color="#000000"><%=getmessagepostLngStr("LtxtSendNewMSG")%></font></a>]</b></font></td>
        </tr>
      </table>
      </td>
    </tr>
    </table>
  </center>
</div>
