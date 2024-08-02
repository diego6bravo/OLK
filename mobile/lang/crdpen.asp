<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "crdpen.xml"
set doccrdpen = server.CreateObject("MSXML2.DOMDocument")
doccrdpen.async = False
Doccrdpen.Load(server.MapPath(xmlfilename)) 
doccrdpen.setProperty "SelectionLanguage", "XPath"
set selectedcrdpennode = doccrdpen.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcrdpennodes=doccrdpen.documentElement.selectNodes("/languages/language")
function getcrdpenLngStr(instring)
	temp = selectedcrdpennode.selectSingleNode(instring).text
	getcrdpenLngStr = temp
end function
%>
