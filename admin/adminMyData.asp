<!--#include file="top.asp" -->
<!-- #INCLUDE file="FCKeditor/fckeditor.asp" -->
<!--#include file="lang/adminMyData.asp" -->
<!--#include file="adminTradSubmit.asp"-->
<% conn.execute("use [" & Session("olkdb") & "]")
sql = "select DataReadOnlyNote from OLKCommon"
set rs = conn.execute(sql) %>
<form method="POST" action="adminsubmit.asp" name="Form1">
	<table border="0" cellpadding="0" width="100%">
		<tr>
			<td bgcolor="#E1F3FD"><b><font color="#FFFFFF">&nbsp;</font><font face="Verdana" size="1" color="#31659C"><%=getadminMyDataLngStr("LttlMyData")%></font></b></td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<p align="justify"><img src="images/lentes.gif"><font face="Verdana" size="1">
			</font><font face="Verdana" size="1" color="#4783C5"><%=getadminMyDataLngStr("LttlMyDataNote")%></font></p>
			</td>
		</tr>
		<tr>
			<td bgcolor="#F5FBFE">
			<div align="left">
				<table border="0" cellpadding="0" width="100%">
					<tr>
						<td bgcolor="#F7FBFF" colspan="2">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font><font face="Verdana" size="1" color="#4783C5">
						<input class="noborder" type="checkbox" <% If myApp.MyDataReadOnly Then %>checked<% End If %> name="MyDataReadOnly" value="Y" id="MyDataReadOnly"><label for="MyDataReadOnly"><%=getadminMyDataLngStr("DtxtReadOnly")%></label></font></td>
					</tr>
					<tr>
						<td bgcolor="#F7FBFF" colspan="2">
						<img src="images/ganchito.gif"><font color="#4783C5">
						</font><font face="Verdana" size="1" color="#4783C5">
						<input class="noborder" type="checkbox" <% If myApp.EnableDROnlyNote Then %>checked<% End If %> name="EnableDROnlyNote" value="Y" id="EnableDROnlyNote"><label for="EnableDROnlyNote"><%=getadminMyDataLngStr("LtxtEnableDROnlyNote")%></label></font></td>
					</tr>
					<tr>
						<td width="140" bgcolor="#F7FBFF" valign="top" style="padding-top: 2px;">
						<img src="images/ganchito.gif"><font face="Verdana" size="1" color="#4783C5">&nbsp;<%=getadminMyDataLngStr("DtxtNote")%></font></td>
						<td bgcolor="#F7FBFF">
						<table cellpadding="0" cellspacing="0" border="0" width="100%">
							<tr>
								<td>
								<%
								Dim oFCKeditor
								Set oFCKeditor = New FCKeditor
								oFCKeditor.BasePath = "FCKeditor/"
								oFCKeditor.Height = 300
								oFCKEditor.ToolbarSet = "Custom"
								If Not IsNull(rs("DataReadOnlyNote")) Then oFCKEditor.Value = rs("DataReadOnlyNote")
								oFCKEditor.Config("AutoDetectLanguage") = False
								If Session("myLng") <> "pt" Then
									oFCKEditor.Config("DefaultLanguage") = Session("myLng")
								Else
									oFCKEditor.Config("DefaultLanguage") = "pt-br"
								End If
								oFCKeditor.Create "DataReadOnlyNote"
								%>
								</td>
								<td width="16" valign="bottom">
								<a href="javascript:doFldTrad('Common', '', '', 'AlterDataReadOnlyNote', 'R', null);"><img src="images/trad.gif" alt="<%=getadminMyDataLngStr("DtxtTranslate")%>" border="0"></a>
								</td>
							</tr>
						</table>
						</td>
					</tr>
					</table>
			</div>
			</td>
		</tr>
		<tr>
			<td>
			<table border="0" cellpadding="0" width="100%">
				<tr>
					<td width="77">
					<input type="submit" value="<%=getadminMyDataLngStr("DtxtSave")%>" name="B1" class="OlkBtn"></td>
					<td><hr color="#0D85C6" size="1"></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
	</table>
	<input type="hidden" name="submitCmd" value="adminMyData">
</form>

<!--#include file="bottom.asp" -->