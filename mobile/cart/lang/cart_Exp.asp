<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cart_Exp.xml"
set doccart_Exp = server.CreateObject("MSXML2.DOMDocument")
doccart_Exp.async = False
Doccart_Exp.Load(server.MapPath(xmlfilename)) 
doccart_Exp.setProperty "SelectionLanguage", "XPath"
set selectedcart_Expnode = doccart_Exp.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcart_Expnodes=doccart_Exp.documentElement.selectNodes("/languages/language")
function getcart_ExpLngStr(instring)
	temp = selectedcart_Expnode.selectSingleNode(instring).text
	getcart_ExpLngStr = temp
end function
%>
