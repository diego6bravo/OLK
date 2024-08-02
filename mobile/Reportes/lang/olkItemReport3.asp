<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "olkItemReport3.xml"
set docolkItemReport3 = server.CreateObject("MSXML2.DOMDocument")
docolkItemReport3.async = False
DocolkItemReport3.Load(server.MapPath(xmlfilename)) 
docolkItemReport3.setProperty "SelectionLanguage", "XPath"
set selectedolkItemReport3node = docolkItemReport3.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedolkItemReport3nodes=docolkItemReport3.documentElement.selectNodes("/languages/language")
function getolkItemReport3LngStr(instring)
	temp = selectedolkItemReport3node.selectSingleNode(instring).text
	getolkItemReport3LngStr = temp
end function
%>
