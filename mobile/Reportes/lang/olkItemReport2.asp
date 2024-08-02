<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "olkItemReport2.xml"
set docolkItemReport2 = server.CreateObject("MSXML2.DOMDocument")
docolkItemReport2.async = False
DocolkItemReport2.Load(server.MapPath(xmlfilename)) 
docolkItemReport2.setProperty "SelectionLanguage", "XPath"
set selectedolkItemReport2node = docolkItemReport2.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedolkItemReport2nodes=docolkItemReport2.documentElement.selectNodes("/languages/language")
function getolkItemReport2LngStr(instring)
	temp = selectedolkItemReport2node.selectSingleNode(instring).text
	getolkItemReport2LngStr = temp
end function
%>
