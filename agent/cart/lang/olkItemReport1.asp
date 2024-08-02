<%
scriptname=Request.ServerVariables("SCRIPT_NAME")
xmlfilename=mid(scriptname, 1, InStrRev(scriptname, "/")) & addLngPathStr & "lang/" & "olkItemReport1.xml"
set docolkItemReport1 = server.CreateObject("MSXML2.DOMDocument")
docolkItemReport1.async = False
DocolkItemReport1.Load(server.MapPath(xmlfilename)) 
docolkItemReport1.setProperty "SelectionLanguage", "XPath"
set selectedolkItemReport1node = docolkItemReport1.selectSingleNode("/languages/language[@xml:lang='" & Session("myLng") & "']") 
set selectedolkItemReport1nodes=docolkItemReport1.documentElement.selectNodes("/languages/language")
function getolkItemReport1LngStr(instring)
	temp = selectedolkItemReport1node.selectSingleNode(instring).text
	getolkItemReport1LngStr = temp
end function
%>
