<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "addcart.xml"
set docaddcart = server.CreateObject("MSXML2.DOMDocument")
docaddcart.async = False
Docaddcart.Load(server.MapPath(xmlfilename)) 
docaddcart.setProperty "SelectionLanguage", "XPath"
set selectedaddcartnode = docaddcart.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedaddcartnodes=docaddcart.documentElement.selectNodes("/languages/language")
function getaddcartLngStr(instring)
	temp = selectedaddcartnode.selectSingleNode(instring).text
	getaddcartLngStr = temp
end function
%>
