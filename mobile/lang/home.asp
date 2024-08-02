<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "home.xml"
set dochome = server.CreateObject("MSXML2.DOMDocument")
dochome.async = False
Dochome.Load(server.MapPath(xmlfilename)) 
dochome.setProperty "SelectionLanguage", "XPath"
set selectedhomenode = dochome.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedhomenodes=dochome.documentElement.selectNodes("/languages/language")
function gethomeLngStr(instring)
	temp = selectedhomenode.selectSingleNode(instring).text
	gethomeLngStr = temp
end function
%>
