<!--#include file="top.asp" -->
<!-- #INCLUDE file="FCKeditor/fckeditor.asp" -->
<!--#include file="lang/adminNewsEdit.asp" -->
<!--#include file="adminTradSubmit.asp"-->

<script type="text/javascript" src="scr/calendar.js"></script>
<script type="text/javascript" src="scr/lang/calendar-<%=Left(Session("myLng"), 2)%>.js"></script>
<script type="text/javascript" src="scr/calendar-setup.js"></script>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="javascript">
function Start(page, w, h, s) {
OpenWin = this.open(page, "DatePicker", "toolbar=no,menubar=no,location=no,scrollbars="+s+",resizable=no, width="+w+",height="+h);
}
function setTimeStamp(vardate) {
document.form1.newsDate.value = vardate
}
function changepic(img_src) {
document['newsImage'].src="pic.aspx?FileName="+img_src+'&maxSize=80&dbName=<%=Session("olkdb")%>';
document.form1.newsImg.value = img_src
document.form1.btnRemImg.disabled = false;
}
</script>
<% 
conn.execute("use [" & Session("OLKDB") & "]")
set rs = Server.CreateObject("ADODB.recordset")
If Request("newsIndex") <> "" Then
	sql = "select newsIndex, newsTitle, newsSmallText, newsText, newsSource, newsDate, Status, " & _
		"Convert(char(10),newsDate,108) newsTime, newsImg " & _
	     " from olknews where newsIndex = " & Request("newsIndex")
	set rs = conn.execute(sql)
	submitCmd = "updateNews"
	If IsNull(rs("newsImg")) or rs("newsImg") = "" Then
		showImg = "n_a.gif"
	Else
		showImg = rs("newsImg")
		newsImg = rs("newsImg")
	End If
	newsTitle = rs("newsTitle")
	newsText = rs("newsText")
	newsSmallText = rs("newsSmallText")
	newsSource = rs("newsSource")
	newsDate = rs("newsDate")
	Status = rs("Status")
Else
	submitCmd = "addNews"
	showImg = "n_a.gif"
	newsTitle = ""
	newsText = ""
	newsSmallText = ""
	newsSource = ""
End If
%>
<script language="javascript">
function valFrm()
{
	if (document.form1.newsTitle.value == '')
	{
		alert("<%=getadminNewsEditLngStr("LtxtValNewsNam")%>");
		document.form1.newsTitle.focus();
		return false;
	}
	else if (document.form1.newsDate.value == '')
	{
		alert("<%=getadminNewsEditLngStr("LtxtValNewsDate")%>");
		document.form1.newsDate.focus();
		return false;
	}
	else if (document.form1.newsSmallText.value == '')
	{
		alert("<%=getadminNewsEditLngStr("LtxtValSmallText")%>");
		document.form1.newsSmallText.focus();
		return false;
	}
	/*else if (document.form1.newsText.value == '')
	{
		alert("<%=getadminNewsEditLngStr("LtxtValNewsCont")%>");
		//document.form1.newsText.focus();
		return false;
	}*/
	return true;
}
</script>
<style type="text/css">
.style1 {
	background-color: #E2F3FC;
}
.style2 {
	background-color: #F3FBFE;
}
</style>
</head>

<form method="POST" action="adminSubmit.asp" name="form1" onsubmit="javascript:return valFrm();">
<% If Request("newsIndex") = "" Then %>
<input type="hidden" name="newsTitleTrad">
<input type="hidden" name="newsSourceTrad">
<input type="hidden" name="newsSmallTextTrad">
<input type="hidden" name="newsTextTrad">
<%
End If
strFormName = "form1"
strTextAreaName = "newsText"
%>
<table border="0" cellpadding="0" width="100%" id="table3">
	<tr>
		<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><% If Request("newsIndex") <> "" Then %><%=getadminNewsEditLngStr("LttlEditNews")%><% Else %><%=getadminNewsEditLngStr("LttlAddNews")%><% End If %></font></b></td>
	</tr>
	<tr>
		<td bgcolor="#F5FBFE">
		<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1"> </font>
		<font face="Verdana" size="1" color="#4783C5"><%=getadminNewsEditLngStr("LttlNewsNote")%></font></td>
	</tr>	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table6">
			<tr>
				<td width="127" class="style1">
				<p align="justify">
				<font face="Verdana" size="1" color="#4783C5"><strong><%=getadminNewsEditLngStr("DtxtTitle")%></strong></font></td>
				<td class="style2">
				<p align="justify">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input class="input" type="text" name="newsTitle" size="70" value="<%=Server.HTMLEncode(newsTitle)%>" onkeydown="return chkMax(event, this, 254);"></td>
						<td><a href="javascript:doFldTrad('News', 'newsIndex', '<%=Request("newsIndex")%>', 'alterNewsTitle', 'T', <% If Request("newsIndex") <> "" Then %>null<% Else %>document.form1.newsTitleTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminNewsEditLngStr("DtxtTranslate")%>" border="0"></a></td>
						<td><input type="checkbox" name="chkActive" value="A" <% If Status = "A" Then %>checked<% End If %> id="chkActive" class="noborder"><font color="#4783C5" face="Verdana" size="1"><label for="chkActive"><%=getadminNewsEditLngStr("DtxtActive")%></label></font></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td width="127" class="style1">
				<p align="justify">
								<font color="#4783C5" face="Verdana" size="1">
								<strong><%=getadminNewsEditLngStr("LtxtSource")%></strong></font></td>
				<td class="style2">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input class="input" type="text" name="newsSource" size="33" value="<% If Not IsNull(newsSource) Then %><%=Server.HTMLEncode(newsSource)%><% End If %>" onkeydown="return chkMax(event, this, 254);"></td>
						<td><a href="javascript:doFldTrad('News', 'newsIndex', '<%=Request("newsIndex")%>', 'alterNewsSource', 'T', <% If Request("newsIndex") <> "" Then %>null<% Else %>document.form1.newsSourceTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminNewsEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td width="127" class="style1">
				<font face="Verdana" size="1" color="#4783C5"><strong><%=getadminNewsEditLngStr("DtxtDateOfPub")%></strong></font></td>
				<td class="style2">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td width="16">
						<img border="0" src="images/cal.gif" id="btnNewsDate" width="16" height="16" style="float:left;padding-left:1px;padding-top:1px"></td>
						<td>
						<input readonly class="input" type="text" name="newsDate" id="newsDate" size="11" value="<%=FormatDate(newsDate, False)%>" onclick="btnNewsDate.click()"></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td width="127" valign="top" class="style1">
				<p align="justify">
				<font face="Verdana" size="1" color="#4783C5"><strong><%=getadminNewsEditLngStr("LtxtMainImg")%></strong></font></td>
				<td class="style2">
				<table border="0" id="table7" cellpadding="0">
					<tr>
						<td>
				<img border="1" name="newsImage" id="newsImage" src="pic.aspx?FileName=<%=showImg%>&MaxSize=80&dbName=<%=Session("olkdb")%>"></td>
						<td valign="bottom">
						<input type="button" value="<%=getadminNewsEditLngStr("LtxtUpload")%>" name="B2" class="OlkBtn" onclick="javascript:Start('upload/fileupload.aspx?ID=<%=Session("ID")%>&style=admin/style/style_pop.css&Source=Admin',300,100,'no')">
							<input type="button" value="X" name="btnRemImg" <% If showImg = "n_a.gif" Then %>disabled<% End If %> style="color: #68A6C0; font-family: Tahoma; border: 1px solid #68A6C0; background-color: #E5F1FF; font-size:10px; width:27; height:23; font-weight:bold" onclick="javascript:if(confirm('<%=getadminNewsEditLngStr("LtxtConfRemImg")%>')){document.form1.newsImg.value='';document.form1.newsImage.src='pic.aspx?FileName=n_a.gif&MaxSize=80&dbName=<%=Session("dbName")%>';this.disabled=true;}">
						</td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td width="127" valign="top" class="style1">
				<font face="Verdana" size="1" color="#4783C5"><strong><%=myHTMLDecode(getadminNewsEditLngStr("LtxtSmallText"))%></strong></font></td>
				<td class="style2">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><textarea rows="2" name="newsSmallText" cols="20" style="width:253px; height:63px" onkeydown="return chkMax(event, this, 155);" onchange="if(this.value.length > 155) this.value = this.value.substring(1, 155);"><%=Server.HTMLEncode(newsSmallText)%></textarea></td>
						<td valign="bottom"><a href="javascript:doFldTrad('News', 'newsIndex', '<%=Request("newsIndex")%>', 'alterNewsSmallText', 'M', <% If Request("newsIndex") <> "" Then %>null<% Else %>document.form1.newsSmallTextTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminNewsEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td width="127" valign="top" class="style1">
				<p align="justify">
				<font color="#4783C5" face="Verdana" size="1"><strong><%=getadminNewsEditLngStr("LtxtContent")%></strong></font></td>
				<td class="style2">
				<table cellpadding="0" cellspacing="0" border="0" width="100%">
				<tr>
					<td><%
							Dim oFCKeditor
							Set oFCKeditor = New FCKeditor
							oFCKeditor.BasePath = "FCKeditor/"
							oFCKeditor.Height = 300
							oFCKEditor.ToolbarSet = "Custom"
							oFCKEditor.Value = newsText
							oFCKEditor.Config("AutoDetectLanguage") = False
							If Session("myLng") <> "pt" Then
								oFCKEditor.Config("DefaultLanguage") = Session("myLng")
							Else
								oFCKEditor.Config("DefaultLanguage") = "pt-br"
							End If
							oFCKeditor.Create "newsText"
							%>
						</td>
						<td width="16" valign="bottom">
						<a href="javascript:doFldTrad('News', 'newsIndex', '<%=Request("newsIndex")%>', 'alterNewsText', 'R', <% If Request("newsIndex") <> "" Then %>null<% Else %>document.form1.newsTextTrad<% End If %>);"><img src="images/trad.gif" alt="<%=getadminNewsEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
			</td>
			</tr>
			</table>
		</td>
	</tr>	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%" id="table5">
			<tr>
				<td width="77">
				<input type="submit" value="<%=getadminNewsEditLngStr("DtxtApply")%>" name="btnApply" class="OlkBtn"></td>
				<td width="77">
				<input type="submit" value="<%=getadminNewsEditLngStr("DtxtSave")%>" name="btnSave" class="OlkBtn"></td>
				<td><hr color="#0D85C6" size="1"></td>
				<td width="77">
				<input type="button" value="<%=getadminNewsEditLngStr("DtxtCancel")%>" name="btnCancel" class="OlkBtn" onclick="javascript:if(confirm('<%=getadminNewsEditLngStr("DtxtConfCancel")%>'))window.location.href='adminNews.asp'"></td>
			</tr>
		</table>
		</td>
	</tr>
	</table>
<input type="hidden" name="newsIndex" value="<%=Request("newsIndex")%>">
<input type="hidden" name="newsImg" value="<%=newsImg%>">
<input type="hidden" name="submitCmd" value="<%=submitCmd%>">
</form>
<script type="text/javascript">
    Calendar.setup({
        inputField     :    "newsDate",     // id of the input field
        ifFormat       :    "<%=GetCalendarFormatString%>",      // format of the input field
        button         :    "btnNewsDate",  // trigger for the calendar (button ID)
        align          :    "Bl",           // alignment (defaults to "Bl")
        singleClick    :    true
    });
</script>
<!--#include file="bottom.asp" -->