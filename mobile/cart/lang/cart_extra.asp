<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cart_extra.xml"
set doccart_extra = server.CreateObject("MSXML2.DOMDocument")
doccart_extra.async = False
Doccart_extra.Load(server.MapPath(xmlfilename)) 
doccart_extra.setProperty "SelectionLanguage", "XPath"
set selectedcart_extranode = doccart_extra.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcart_extranodes=doccart_extra.documentElement.selectNodes("/languages/language")
function getcart_extraLngStr(instring)
	temp = selectedcart_extranode.selectSingleNode(instring).text
	getcart_extraLngStr = temp
end function
%>
