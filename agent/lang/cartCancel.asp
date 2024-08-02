<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cartCancel.xml"
set doccartCancel = server.CreateObject("MSXML2.DOMDocument")
doccartCancel.async = False
DoccartCancel.Load(server.MapPath(xmlfilename)) 
doccartCancel.setProperty "SelectionLanguage", "XPath"
set selectedcartCancelnode = doccartCancel.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartCancelnodes=doccartCancel.documentElement.selectNodes("/languages/language")
function getcartCancelLngStr(instring)
	temp = selectedcartCancelnode.selectSingleNode(instring).text
	getcartCancelLngStr = temp
end function
%>
