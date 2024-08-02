<%
Class clsLock
	Dim ctlLockID
	Dim ctlValue
	Dim ctlOnClick
	
	Public Property Let ID(p_Value)
		ctlLockID = p_Value
	End Property
	
	Public Property Let Value(p_Value)
		ctlValue = p_Value
	End Property
	
	Public Property Let OnClick(p_Value)
		ctlOnClick = p_Value
	End Property

	
	Sub GenerateLock
	If Not ctlValue Then addImgStr = "un" Else addImgStr = ""
	 %><img id="imgLock<%=ctlLockID%>" src="images/icon_<%=addImgStr%>lock.jpg" onclick="ClickUnlock('<%=ctlLockID%>');<%=ctlOnClick%>" style="cursor: pointer;">
	<input type="hidden" id="hdLockVal<%=ctlLockID%>" name="Lock<%=ctlLockID%>" value="<%=GetYN(ctlValue)%>"><%
	End Sub 
End Class 
%>