<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "clientSearch.xml"
set docclientSearch = server.CreateObject("MSXML2.DOMDocument")
docclientSearch.async = False
DocclientSearch.Load(server.MapPath(xmlfilename)) 
docclientSearch.setProperty "SelectionLanguage", "XPath"
set selectedclientSearchnode = docclientSearch.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedclientSearchnodes=docclientSearch.documentElement.selectNodes("/languages/language")
function getclientSearchLngStr(instring)
	temp = selectedclientSearchnode.selectSingleNode(instring).text
	getclientSearchLngStr = temp
end function
%>
