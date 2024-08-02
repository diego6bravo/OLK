<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cart.js.xml"
set doccartjs = server.CreateObject("MSXML2.DOMDocument")
doccartjs.async = False
Doccartjs.Load(server.MapPath(xmlfilename)) 
doccartjs.setProperty "SelectionLanguage", "XPath"
set selectedcartjsnode = doccartjs.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartjsnodes=doccartjs.documentElement.selectNodes("/languages/language")
function getcartjsLngStr(instring)
	temp = selectedcartjsnode.selectSingleNode(instring).text
	getcartjsLngStr = temp
end function
%>
