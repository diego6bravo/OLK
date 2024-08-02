<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "clientTop.xml"
set docclientTop = server.CreateObject("MSXML2.DOMDocument")
docclientTop.async = False
DocclientTop.Load(server.MapPath(xmlfilename)) 
docclientTop.setProperty "SelectionLanguage", "XPath"
set selectedclientTopnode = docclientTop.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedclientTopnodes=docclientTop.documentElement.selectNodes("/languages/language")
function getclientTopLngStr(instring)
	temp = selectedclientTopnode.selectSingleNode(instring).text
	getclientTopLngStr = temp
end function
%>
