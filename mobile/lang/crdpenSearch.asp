<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "crdpenSearch.xml"
set doccrdpenSearch = server.CreateObject("MSXML2.DOMDocument")
doccrdpenSearch.async = False
DoccrdpenSearch.Load(server.MapPath(xmlfilename)) 
doccrdpenSearch.setProperty "SelectionLanguage", "XPath"
set selectedcrdpenSearchnode = doccrdpenSearch.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcrdpenSearchnodes=doccrdpenSearch.documentElement.selectNodes("/languages/language")
function getcrdpenSearchLngStr(instring)
	temp = selectedcrdpenSearchnode.selectSingleNode(instring).text
	getcrdpenSearchLngStr = temp
end function
%>
