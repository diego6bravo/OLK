<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "clientdel.xml"
set docclientdel = server.CreateObject("MSXML2.DOMDocument")
docclientdel.async = False
Docclientdel.Load(server.MapPath(xmlfilename)) 
docclientdel.setProperty "SelectionLanguage", "XPath"
set selectedclientdelnode = docclientdel.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedclientdelnodes=docclientdel.documentElement.selectNodes("/languages/language")
function getclientdelLngStr(instring)
	temp = selectedclientdelnode.selectSingleNode(instring).text
	getclientdelLngStr = temp
end function
%>
