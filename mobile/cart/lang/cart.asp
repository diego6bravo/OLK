<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cart.xml"
set doccart = server.CreateObject("MSXML2.DOMDocument")
doccart.async = False
Doccart.Load(server.MapPath(xmlfilename)) 
doccart.setProperty "SelectionLanguage", "XPath"
set selectedcartnode = doccart.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartnodes=doccart.documentElement.selectNodes("/languages/language")
function getcartLngStr(instring)
	temp = selectedcartnode.selectSingleNode(instring).text
	getcartLngStr = temp
end function
%>
