<% addLngPathStr = "mensaje_sbo/" %>
<!--#include file="lang/messagepost.asp" -->
<%
set conn3 = Server.CreateObject("ADODB.Connection")
set Cmd = Server.CreateObject("ADODB.Command")
conn3.Provider = olkSqlProv
conn3.Open  "Provider=SQLOLEDB;charset=utf8;" & _
          "Data Source=" & olkip & ";" & _
          "Initial Catalog=R3_ObsCommon;" & _
          "Uid=" & olklogin & ";" & _
          "Pwd=" & olkpass & ""
      set rs = Server.CreateObject("ADODB.recordset")
      set var1 = Server.CreateObject("ADODB.recordset")
      
db = Session("olkdb")
Set cmd.ActiveConnection = conn3
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "OBSSp_Request"
cmd.Parameters.Refresh
cmd.Execute , Array(0, db, Null, 81, "A", Null)
RetVal2 = cmd.Parameters.Item(0).Value
OBSSp_Request = RetVal2

If Request("Urgent") <> "Y" Then Priority = 1 Else Priority = 2
sql = "insert into tmsg(LogNum, Priority, Subject, UserText) " & _
"values('" & RetVal2 & "', " & Priority & ",N'" & Request.Form("Subject") & "', N'" & Request.Form("Message") & "')"
conn3.execute(sql)

sql = ""
ArrVal = Split(Request("SapTo"),", ")
For i = 0 to UBound(ArrVal)
	sql = sql & "insert into msg1(LogNum, UserCode, SendIntrnl) values('" & RetVal2 & "', N'" & ArrVal(i) & "', 'Y') "
next
'response.redirect "../../query.asp?query=" & sql
conn3.execute(sql)
sql = "update tlog set status = 'C', ErrLng = '" & GetLangErrCode() & "' where lognum = " & RetVal2
conn3.execute(sql)

conn3.close
%>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
    <!-- fwtable fwsrc="Z:\topmanage\logos\originales\pocket_art.png" fwbase="pocket_artpieza1.gif" fwstyle="FrontPage" fwdocid = "742308039" fwnested=""0" -->
    <tr>
      <td bgcolor="#9BC4FF">
      <table border="0" cellpadding="0" bordercolor="#111111" width="100%" id="AutoNumber1">
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
          <p align="center"><font SIZE="1" face="Verdana"><b>[<a href="operaciones.asp?cmd=msssbo"><font color="#000000"><%=getmessagepostLngStr("LtxtSendNewMSG")%></font></a>]</b></font></td>
        </tr>
      </table>
      </td>
    </tr>
    </table>
  </center>
</div>