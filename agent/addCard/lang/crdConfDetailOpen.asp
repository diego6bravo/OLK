<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "crdConfDetailOpen.xml"
set doccrdConfDetailOpen = server.CreateObject("MSXML2.DOMDocument")
doccrdConfDetailOpen.async = False
DoccrdConfDetailOpen.Load(server.MapPath(xmlfilename)) 
doccrdConfDetailOpen.setProperty "SelectionLanguage", "XPath"
set selectedcrdConfDetailOpennode = doccrdConfDetailOpen.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedcrdConfDetailOpennodes=doccrdConfDetailOpen.documentElement.selectNodes("/languages/language")
function getcrdConfDetailOpenLngStr(instring)
	temp = selectedcrdConfDetailOpennode.selectSingleNode(instring).text
	getcrdConfDetailOpenLngStr = temp
end function
%>
