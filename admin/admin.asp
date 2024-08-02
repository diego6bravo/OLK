<!--#include file="top.asp" -->
<!--#include file="lang/admin.asp" -->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<style type="text/css">
.style1 {
	color: #31659C;
	background-color: #F0F9FE;
	font-family: Verdana;
	font-size: xx-small;
	font-weight: bold;
	text-align: center; 
}
.style2 {
	color: #4783C5;
	background-color: #F5FBFE;
	text-align: justify;
	font-family: Verdana;
	font-size: xx-small;
}
.style3 {
	color: #31659C;
	background-color: #E4F4FC;
	font-family: Verdana;
	font-size: xx-small;
	font-weight: bold;
	text-align: <% If Session("rtl") = "" Then %>right<% Else %>left<% End If %>; 
}

.style4 {
	border-collapse: collapse;
	border-color: #111111;
}

</style>
</head>
<% 
cmd.ActiveConnection = connCommon

If Request("searchDB") = "True" Then
	cmd.CommandText = "OLKSearchDB"
	cmd.Parameters.Refresh()
	cmd.Execute()
End If

set rs = Server.CreateObject("ADODB.RecordSet")
cmd.CommandText = "OLKGetDBList"
cmd.Parameters.Refresh()
cmd("@UserType") = "N"
cmd("@curURL") = ""
set rs = cmd.Execute()

If Request("dbErr") = "Y" Then
%>
<script type="text/javascript">alert('<%=getadminLngStr("LtxtActiveDBErr")%>');</script>
<% End If %>
<br/>
	<div style="text-align: center">
		<table border="0" cellpadding="0" style="width: 100%; ">
			<tr>
				<td>
				<div style="text-align: center; ">
                  <center>
					<table border="0" cellpadding="0" width="600">
						<tr>
							<td class="style1">
							<%=getadminLngStr("LMdlAdmin")%></td>
						</tr>
						<tr>
							<td class="style2">
							<%=getadminLngStr("LMdlDesc")%></td>
						</tr>
						</table>
					</center>
				</div>
				</td>
			</tr>
			<tr>
				<td>
				<div style="text-align: center; ">
                  <center>
                  <table border="0" cellpadding="0" width="600" cellspacing="2" class="style4">
                    <tr>
                      <td class="style1"><%=getadminLngStr("LAdminDB")%></td>
                    </tr>
                    <tr>
                      <td style="background-color: #F5FBFE;">
                      <form method="post" action="adminsubmit.asp" name="frmChooseDB" onsubmit="return (document.frmChooseDB.dbName.selectedIndex!=-1);">
                        <div style="text-align: center; ">
                          <center>
                          <table border="0" cellpadding="0" width="95%" class="style4">
                            <tr>
                              <td colspan="2" style="font-size: 10px; width: 302px; ">&nbsp;</td>
                            </tr>
                            <tr>
                              <td rowspan="5" style="background-color: #E4F4FC; ">
                              <select size="4" name="dbName" class="input" style="height:194px; width: 429px;" ondblclick="if(this.selectedIndex!=-1)document.frmChooseDB.submit()">
                              <% do while not rs.eof %>
                              <option dir="ltr" <% If rs("ID") = Session("ID") then %>selected<% end if %> value="<%=rs("ID")%>">(<%=rs("dbName")%>)&nbsp;<%=myHTMLEncode(rs("cmpName"))%><% If rs("Verfy") = "Y" Then Response.Write " (" & getadminLngStr("LSmallDef") & ")" %> 
								v<%=rs("Version")%></option>
                              <% rs.movenext
                              loop %>
                              </select></td>
                              <td class="style3" style="height: 38px; width: 120px;">
                              <img alt="" src="images/<%=Session("rtl")%>admin_home_arrow_left.jpg"/>&nbsp;<%=getadminLngStr("DtxtDB")%></td>
                            </tr>
                          	<tr>
                              <td class="style3" style="text-align: center; height: 38px; width: 120px;">
                              <input class="OlkBtn" type="button" value="<%=getadminLngStr("LSearch")%>" name="B5" onclick="javascript:window.location.href='?cmd=home&searchDB=True'" style="width: 120px; "/></td>
                            </tr>
                          	<tr>
                              <td class="style3" style="text-align: center; height: 38px; width: 120px;">
                              <input class="OlkBtn" type="submit" value="<%=getadminLngStr("LChoose")%>" name="B5" style="width: 120px; "/></td>
                            </tr>
                          	<tr>
                              <td class="style3" style="text-align: center; height: 38px; width: 120px;">
                              <input class="OlkBtn" type="button" value="<%=getadminLngStr("LbtnSmallDef")%>" name="B2" style="width: 120px; " onclick="javascript:window.location.href='adminSubmit.asp?submitCmd=changeDefDb&dbID='+document.frmChooseDB.dbName.value"/></td>
                            </tr>
                          	<tr>
                              <td class="style3" style="text-align: center; height: 38px; width: 120px;">
                              <input class="OlkBtn" type="button" value="<%=getadminLngStr("LUnInstall")%>" name="B4" style="width: 120px; " onclick="javascript:if(confirm('<%=getadminLngStr("LUnInstConf")%>')){this.disabled=true;window.location.href='remDb/adminRemDb.asp?dbName='+document.frmChooseDB.dbName.value+'&AddPath=';}"/></td>
                            </tr>
                          </table>
                          <br/>
                          </center>
                        </div>
                      	<input type="hidden" name="submitCmd" value="changeDb"/>
                      </form>
                      </td>
                    </tr>
                    <tr>
                      <td class="style1"><%=getadminLngStr("LOlkActivTtl")%></td>
                    </tr>
                    <tr>
                      <td style="background-color: #F5FBFE; ">
                      <form method="post" action="addDb/adminAddDb.asp" name="frmActiveDB">
                      <input type="hidden" name="AddPath" value="" />
                        <div style="text-align: center; ">
                        <center>
                          <table border="0" cellpadding="0" width="95%" class="style4">
                            <tr>
                              <td colspan="2" style="font-size: 10px; ">&nbsp;</td>
                            </tr>
                            <tr>
                              <td style="background-color: #E4F4FC; " rowspan="2">
                              <select size="4" name="dbName" class="input" style="width: 429px; height:100px; ">
                              <% 
                              cmd.CommandText = "OLKGetDBListUA"
                              cmd.Parameters.Refresh()
                              set rs = cmd.execute()
                              do while not rs.eof %>
                              <option value="<%=rs("dbName")%>">(<%=rs("dbName")%>)&nbsp;<%=myHTMLEncode(rs("cmpName"))%></option>
                              <% rs.movenext
                              loop %>
                              </select></td>
                              <td class="style3" style="width: 120px"><img src="images/<%=Session("rtl")%>admin_home_arrow_left.jpg">&nbsp;<%=getadminLngStr("DtxtDB")%></td>
                            </tr>
                            <tr>
                              <td style="background-color: #E4F4FC; text-align: center; width: 120px;">
                              <input class="OlkBtn" type="button" value="<%=getadminLngStr("LActivate")%>" style="width: 120px; " name="B1" onclick="this.disabled=true;document.frmActiveDB.submit();"></td>
                            </tr>
                          </table>
                        </center>
                        </div>
                      </form>
                      </td>
                    </tr>
                  </table>
                  </center>
                </div>
				</td>
			</tr>
		</table>
	</div>
	<% If Request("dbVMErr") = "Y" Then
	rs.Filter = "ID = " & Request("ID") %>
	<script language="javascript">alert('<%=getadminLngStr("LDBSupVer")%>'.replace('{0}', '<%=Replace(rs("cmpName"), "'", "\'")%>').replace('{1}', '<%=rs("dbName")%>'));</script>
	<% ElseIf Request("LineMemoErr") = "Y" Then %>
	<script language="javascript">
	alert('<%=getadminLngStr("LtxtLineMemoErr")%>'.replace('{0}', '<%=Request("dbName")%>'));
	</script>
	<script language="ecmascript"></script>
	<% End If %><!--#include file="bottom.asp" -->