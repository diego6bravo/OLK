<% addLngPathStr = "" %><!--#include file="lang/viewType.asp" -->
<%
Dim viewTypeCount

viewTypeCount = 0

Class clsViewType
	Dim varViewTypeID
	Dim varViewTypeValue
	Dim varAlterColor
	Dim varOnClick
	Dim varHandCursor
	
	Public Property Let ID(p_Value)
		varViewTypeID = p_Value
	End Property
	
	Public Property Let Value(p_Value)
		varViewTypeValue = p_Value
	End Property
	
	Public Property Let AlterColor(p_Value)
		varAlterColor = p_Value
	End Property
	
	Public Property Let OnClick(p_Value)
		varOnClick = p_Value
	End Property
	
	Public Property Let HandCursor(p_Value)
		varHandCursor = p_Value
	End Property

	Public Sub doViewType
	
	If Not varAlterColor Then Color = "black" Else Color = "white" %>
	<table cellpadding="2" cellspacing="2" border="0">
		<tr>
			<td id="tdViewTypeT<%=viewTypeCount%>" style="border-bottom: 2px solid <% If varViewTypeValue = "T" Then %><%=Color%><% Else %>transparent<% End If %>; "><img src="images/<%=Session("rtl")%>searchStoreIcon<% If varAlterColor Then %>White<% End If %>.gif" alt="<%=getviewTypeLngStr("DtxtStore")%>" onclick="selViewType('T', <%=viewTypeCount%>);<%=Replace(varOnClick, "{Type}", "T")%>" <% If varHandCursor Then %>style="cursor: pointer; "<% End If %>></td>
			<td id="tdViewTypeC<%=viewTypeCount%>" style="border-bottom: 2px solid <% If varViewTypeValue = "C" Then %><%=Color%><% Else %>transparent<% End If %>; "><img src="images/searchCatalogIcon<% If varAlterColor Then %>White<% End If %>.gif" alt="<%=getviewTypeLngStr("DtxtCat")%>" onclick="selViewType('C', <%=viewTypeCount%>);<%=Replace(varOnClick, "{Type}", "C")%>" <% If varHandCursor Then %>style="cursor: pointer; "<% End If %>></td>
			<td id="tdViewTypeL<%=viewTypeCount%>" style="border-bottom: 2px solid <% If varViewTypeValue = "L" Then %><%=Color%><% Else %>transparent<% End If %>; "><img src="images/searchListIcon<% If varAlterColor Then %>White<% End If %>.gif" alt="<%=getviewTypeLngStr("DtxtList")%>" onclick="selViewType('L', <%=viewTypeCount%>);<%=Replace(varOnClick, "{Type}", "L")%>" <% If varHandCursor Then %>style="cursor: pointer; "<% End If %>></td>
		</tr>
	</table>
	<input type="hidden" id="varViewTypeID<%=viewTypeCount%>" name="<%=varViewTypeID%>" value="<%=varViewTypeValue%>">
	<input type="hidden" id="varAlterColor<%=viewTypeCount%>" value="<%=Color%>">
	
<%	viewTypeCount = viewTypeCount + 1
	End Sub
End Class

%>