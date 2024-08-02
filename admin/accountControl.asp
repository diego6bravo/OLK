<%
Class AccountControl
	Dim accID
	Dim	accValue
	Dim accDisp
	Dim accDesc
	Dim accType
	
	Public Property Let ID(p_Value)
		accID = p_Value
	End Property
	
	Public Property Let Value(p_Value)
		accValue = p_Value
	End Property
	
	Public Property Let DisplayValue(p_Value)
		accDisp = p_Value
	End Property
	
	Public Property Let Description(p_Value)
		accDesc = p_Value
	End Property
	
	Public Property Let AccountType(p_Value)
		accType = p_Value
	End Property

	Sub GenerateAccount %>
<table border="0" cellspacing="0" width="420" cellpadding="0" class="TblCombo">
	<tr>
		<td style="cursor: default; font-size: 10px;" onclick="return !showSelectAccount('imgCmb<%=accID%>', '<%=accID%>', '<%=accType%>', event);"><span id="txtSel<%=accID%>"><%=accDisp%> - <%=accDesc%></span></td>
		<td width="12"><img src="images_picker/select_arrow_small.gif" id="imgCmb<%=accID%>" onmouseover="this.src='images_picker/select_arrow_over_small.gif'" onmouseout="this.src='images_picker/select_arrow_small.gif'" onclick="return !showSelectAccount(this, '<%=accID%>', '<%=accType%>', event);"></td>
	</tr>
</table>
<input type="hidden" name="<%=accID%>" id="<%=accID%>" value="<%=accValue%>">
<% End Sub
End Class %>
