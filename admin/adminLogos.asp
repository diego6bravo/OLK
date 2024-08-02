<!--#include file="top.asp" -->
<!--#include file="lang/adminLogos.asp" -->
<head>
<% 
conn.execute("use [" & Session("OLKDB") & "]")
set rd = Server.CreateObject("ADODB.recordset")
set rs = Server.CreateObject("ADODB.recordset")

If Not IsNull(myApp.TopLogo) Then TopLogo = myApp.TopLogo Else TopLogo = "n_a.gif"
If Not IsNull(myApp.MailLogo) Then MailLogo = myApp.MailLogo Else MailLogo = "n_a.gif"
If Not IsNull(myApp.AgentLogo) Then AgentLogo = "imagenes/" & Session("olkdb") & "/" & myApp.AgentLogo Else AgentLogo = "pic.aspx?FileName=n_a.gif&maxSize=70&dbName=" & Session("olkdb")
%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript">
var changeImg;
var changeFld;
function setPic(img, fld)
{
	changeImg = img;
	changeFld = fld;
	Start('upload/fileupload.aspx?ID=<%=Session("ID")%>&Source=admin&style=admin/style/style_pop.css',400,100,'no');
}
function Start(page, w, h, s) {
OpenWin = this.open(page, "DatePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=no, width="+w+",height="+h);
}
function changepic(img_src) {
	if (changeImg.id != 'AgentLogoImg')
		changeImg.src = "pic.aspx?FileName="+img_src+'&maxSize=250&dbName=<%=Session("olkdb")%>';
	else
		changeImg.src = 'imagenes/<%=Session("olkdb")%>/' + img_src;
		
	changeFld.value = img_src;
	switch (changeImg.id)
	{
		case 'TopLogoImg':
			document.Form1.btnRemTopLogo.disabled = false;
			break;
		case 'MailLogoImg':
			document.Form1.btnRemMailLogo.disabled = false;
			break;
		case 'AgentLogoImg':
			document.Form1.btnRemAgentLogo.disabled = false;
			break;
	}
}
</script>
</head>

<% If Session("style") = "nc" Then %>
<br>
<% End If %>
<form method="POST" action="adminsubmit.asp" name="Form1">
<table border="0" cellpadding="0" width="100%" id="table3">
	<tr>
		<td bgcolor="#E1F3FD"><b><font face="Verdana" size="2">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminLogosLngStr("LtxtLogo")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
		<font color="#4783C5"><%=getadminLogosLngStr("LtxtLogoNote")%></font></font></td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD"><b><font face="Verdana" size="2">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminLogosLngStr("LtxtClientSession")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<table border="0" cellpadding="0">
			<tr>
				<td align="center" width="300"><img src="pic.aspx?FileName=<%=TopLogo%>&maxSize=250&dbName=<%=Session("olkdb")%>" id="TopLogoImg" border="0"></td>
			</tr>
			<tr>
				<td align="center" width="300">
				<input type="button" value="<%=getadminLogosLngStr("DtxtChange")%>" name="B2" class="OlkBtn" onclick="javascript:setPic(TopLogoImg, TopLogo);">
				<input type="button" value="X" name="btnRemTopLogo" <% If TopLogo = "n_a.gif" Then %>disabled<% End If %> style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; width:27; height:23; font-weight:bold" onclick="javascript:if(confirm('<%=getadminLogosLngStr("LtxtConfRemLogo")%>')){document.Form1.TopLogo.value='';document.Form1.TopLogoImg.src='pic.aspx?FileName=n_a.gif&MaxSize=250&dbName=<%=Session("dbName")%>';this.disabled=true;}"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
	<tr>
		<td bgcolor="#E1F3FD"><b><font face="Verdana" size="2">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminLogosLngStr("LtxtSendMail")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<table border="0" cellpadding="0">
			<tr>
				<td align="center" width="300"><img src="pic.aspx?FileName=<%=MailLogo%>&maxSize=250&dbName=<%=Session("olkdb")%>" id="MailLogoImg" border="0"></td>
			</tr>
			<tr>
				<td align="center" width="300">
				<input type="button" value="<%=getadminLogosLngStr("DtxtChange")%>" name="B3" class="OlkBtn" onclick="javascript:setPic(MailLogoImg, MailLogo);">
				<input type="button" value="X" name="btnRemMailLogo" <% If MailLogo = "n_a.gif" Then %>disabled<% End If %> style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; width:27; height:23; font-weight:bold" onclick="javascript:if(confirm('<%=getadminLogosLngStr("LtxtConfRemLogo")%>')){document.Form1.MailLogo.value='';document.Form1.MailLogoImg.src='pic.aspx?FileName=n_a.gif&MaxSize=250&dbName=<%=Session("dbName")%>';this.disabled=true;}"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E1F3FD"><b><font face="Verdana" size="2">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminLogosLngStr("LtxtAgentSession")%></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<table border="0" cellpadding="0">
			<tr>
				<td align="center" width="300">
				<table cellpadding="0" cellspacing="0" border="1" width="300" height="70">
					<tr>
						<td bgcolor="#0073E6" valign="middle" align="center">
						<img src="<%=AgentLogo%>" id="AgentLogoImg" border="0">
						</td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td align="center" width="300">
				<input type="button" value="<%=getadminLogosLngStr("DtxtChange")%>" name="B3" class="OlkBtn" onclick="javascript:setPic(AgentLogoImg, AgentLogo);">
				<input type="button" value="X" name="btnRemAgentLogo" <% If AgentLogo = "n_a.gif" Then %>disabled<% End If %> style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; width:27; height:23; font-weight:bold" onclick="javascript:if(confirm('<%=getadminLogosLngStr("LtxtConfRemLogo")%>')){document.Form1.AgentLogo.value='';document.Form1.AgentLogoImg.src='pic.aspx?FileName=n_a.gif&MaxSize=70&dbName=<%=Session("dbName")%>';this.disabled=true;}"></td>
			</tr>
			<tr>
				<td align="center" width="300">
				<font face="Verdana" size="1" color="#4783C5"><%=getadminLogosLngStr("LtxtHomeLogo")%> 300x70</font></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminLogosLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
</table>
<input type="hidden" id="TopLogo" name="TopLogo" value="<%=myApp.TopLogo%>">
<input type="hidden" id="MailLogo" name="MailLogo" value="<%=myApp.MailLogo%>">
<input type="hidden" id="AgentLogo" name="AgentLogo" value="<%=myApp.AgentLogo%>">
<input type="hidden" name="submitCmd" value="adminLogos">
</form>
<!--#include file="bottom.asp" -->