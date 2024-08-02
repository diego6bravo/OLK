<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "cartApp.xml"
set doccartApp = server.CreateObject("MSXML2.DOMDocument")
doccartApp.async = False
DoccartApp.Load(server.MapPath(xmlfilename)) 
doccartApp.setProperty "SelectionLanguage", "XPath"
set selectedcartAppnode = doccartApp.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcartAppnodes=doccartApp.documentElement.selectNodes("/languages/language")
function getcartAppLngStr(instring)
	temp = selectedcartAppnode.selectSingleNode(instring).text
	getcartAppLngStr = temp
end function
%>
