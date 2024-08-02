<!--#include file="clientInc.asp"-->
<!--#include file="agentTop.asp"-->
<% If userType <> "V" Then Response.Redirect "default.asp" %>
<% If Not (myAut.HasAuthorization(4)) Then  Response.Redirect "unauthorized.asp" %>
<% addLngPathStr = "" %>
<!--#include file="lang/newsEdit.asp" -->
<!--#include file="genman/adminTradForm.asp"-->

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
document['newsImage'].src="pic.aspx?FileName="+img_src+'&maxSize=80';
document.form1.newsImg.value = img_src
}
</script>
<% 

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
		alert("<%=getnewsEditLngStr("LtxtValNewsNam")%>");
		document.form1.newsTitle.focus();
		return false;
	}
	else if (document.form1.newsDate.value == '')
	{
		alert("<%=getnewsEditLngStr("LtxtValNewsDate")%>");
		document.form1.newsDate.focus();
		return false;
	}
	else if (document.form1.newsSmallText.value == '')
	{
		alert("<%=getnewsEditLngStr("LtxtValSmallText")%>");
		document.form1.newsSmallText.focus();
		return false;
	}
	/*else if (document.form1.newsText.value == '')
	{
		alert("<%=getnewsEditLngStr("LtxtValNewsCont")%>");
		//document.form1.newsText.focus();
		return false;
	}*/
	return true;
}
</script>
</head>

<form method="post" action="genman/newsSubmit.asp" name="form1" onsubmit="javascript:return valFrm();">
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
<table border="0" cellpadding="0" width="100%">
	<tr>
		<td class="GeneralTlt"><% If Request("newsIndex") <> "" Then %><%=getnewsEditLngStr("LttlEditNews")%><% Else %><%=getnewsEditLngStr("LttlAddNews")%><% End If %></td>
	</tr>
	<tr>
		<td>
		<table border="0" cellpadding="0" width="100%">
			<tr>
				<td width="127" class="CanastaTblResaltada"><%=getnewsEditLngStr("DtxtTitle")%></td>
				<td class="CanastaTbl">
				<p align="justify">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input class="input" type="text" name="newsTitle" size="70" value="<%=Server.HTMLEncode(newsTitle)%>" onkeydown="return chkMax(event, this, 254);"></td>
						<td><a href="javascript:doFldTrad('News', 'newsIndex', '<%=Request("newsIndex")%>', 'alterNewsTitle', 'T', <% If Request("newsIndex") <> "" Then %>null<% Else %>document.form1.newsTitleTrad<% End If %>);">
						<img src="images/trad.gif" alt="<%=getnewsEditLngStr("DtxtTranslate")%>" border="0"></a></td>
						<td><input type="checkbox" name="chkActive" value="A" <% If Status = "A" Then %>checked<% End If %> id="chkActive" class="noborder"><font color="#4783C5" face="Verdana" size="1"><label for="chkActive"><%=getnewsEditLngStr("DtxtActive")%></label></font></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td width="127" class="CanastaTblResaltada"><%=getnewsEditLngStr("LtxtSource")%></td>
				<td class="CanastaTbl">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><input class="input" type="text" name="newsSource" size="33" value="<% If Not IsNull(newsSource) Then %><%=Server.HTMLEncode(newsSource)%><% End If %>" onkeydown="return chkMax(event, this, 254);"></td>
						<td><a href="javascript:doFldTrad('News', 'newsIndex', '<%=Request("newsIndex")%>', 'alterNewsSource', 'T', <% If Request("newsIndex") <> "" Then %>null<% Else %>document.form1.newsSourceTrad<% End If %>);">
						<img src="images/trad.gif" alt="<%=getnewsEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td width="127" class="CanastaTblResaltada"><%=getnewsEditLngStr("DtxtDateOfPub")%></td>
				<td class="CanastaTbl">
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
				<td width="127" valign="top" class="CanastaTblResaltada"><%=getnewsEditLngStr("LtxtMainImg")%></td>
				<td class="CanastaTbl">
				<table border="0" id="table7" cellpadding="0">
					<tr>
						<td>
				<img border="1" name="newsImage" id="newsImage" src="pic.aspx?FileName=<%=showImg%>&amp;MaxSize=80"></td>
						<td valign="bottom">
						<input type="button" value="<%=getnewsEditLngStr("DtxtAddImg")%>" name="B2" class="OlkBtn" onclick="javascript:Start('../upload/fileupload.aspx?ID=<%=Session("ID")%>&style=../design/<%=SelDes%>/style/stylePopUp.css',400,100,'no')">
						<img border="0" src="images/<%=Session("rtl")%>remove.gif" width="16" height="16" onclick="javascript:javascript:if(confirm('<%=getnewsEditLngStr("LtxtConfRemImg")%>')){document.form1.newsImg.value='';document.form1.newsImage.src='../pic.aspx?FileName=n_a.gif&MaxSize=80&dbName=<%=Session("dbName")%>'};" style="cursor: hand">
						</td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td width="127" valign="top" class="CanastaTblResaltada"><%=myHTMLDecode(getnewsEditLngStr("LtxtSmallText"))%></td>
				<td class="CanastaTbl">
				<table cellpadding="0" cellspacing="0" border="0">
					<tr>
						<td><textarea rows="2" name="newsSmallText" cols="20" style="width:253px; height:63px" onkeydown="return chkMax(event, this, 155);" onchange="if(this.value.length > 155) this.value = this.value.substring(1, 155);"><%=Server.HTMLEncode(newsSmallText)%></textarea></td>
						<td valign="bottom"><a href="javascript:doFldTrad('News', 'newsIndex', '<%=Request("newsIndex")%>', 'alterNewsSmallText', 'M', <% If Request("newsIndex") <> "" Then %>null<% Else %>document.form1.newsSmallTextTrad<% End If %>);">
						<img src="images/trad.gif" alt="<%=getnewsEditLngStr("DtxtTranslate")%>" border="0"></a></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td width="127" valign="top" class="CanastaTblResaltada"><%=getnewsEditLngStr("LtxtContent")%></td>
				<td class="CanastaTbl">
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
						<a href="javascript:doFldTrad('News', 'newsIndex', '<%=Request("newsIndex")%>', 'alterNewsText', 'R', <% If Request("newsIndex") <> "" Then %>null<% Else %>document.form1.newsTextTrad<% End If %>);">
						<img src="images/trad.gif" alt="<%=getnewsEditLngStr("DtxtTranslate")%>" border="0"></a></td>
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
				<td width="1">
				<input type="submit" value="<%=getnewsEditLngStr("DtxtApply")%>" name="btnApply"></td>
				<td width="1">
				<input type="submit" value="<%=getnewsEditLngStr("DtxtSave")%>" name="btnSave"></td>
				<td>&nbsp;</td>
				<td width="1">
				<input type="button" value="<%=getnewsEditLngStr("DtxtCancel")%>" name="btnCancel"onclick="javascript:if(confirm('<%=getnewsEditLngStr("DtxtConfCancel")%>'))window.location.href='newsman.asp'"></td>
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
<!--#include file="agentBottom.asp"-->