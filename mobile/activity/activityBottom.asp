<% 
If Not ReadOnly Then
Select Case Request("cmd")
Case "activityGeneral"
	btnSave = "btnGeneral"
Case "activityAddress"
	btnSave = "btnAddress"
Case "activityContent"
	btnSave = "btnContent"
Case "activityUDF"
	btnSave = "btnUDF"
Case Else
	btnSave = "btnMain"
End Select %>
        <TR>
        <td>
          <div align="center">
            <center>
            <table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">
              <tr>
              	<td>&nbsp;</td>
                <td width="25">
                <p align="center">
                <input border="0" src="images/save_icon.gif" name="<%=btnSave%>" type="image"></td>
                <td width="31">&nbsp;</td>
                <td width="37">
                <p align="center"><input type="image" name="btnAdd" value="btnAdd" border="0" src="images/ok_icon.gif" onclick="javascript:return valFrm();"></td>
                <% If IsNull(ClgCode) Then %>
				<td width="29">
                <p align="center"><a href="operaciones.asp?cmd=actCancel"><img border="0" src="images/x_icon.gif"></a></td><% End If %>
              </tr>
            </table>
            </center>
          </div>
        </td></TR>
<% End If %>